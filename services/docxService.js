import fs from 'fs';
import path from 'path';
import { createRequire } from 'module';
import { createCanvas } from '@napi-rs/canvas';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import {
  Document,
  Paragraph,
  TextRun,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  Packer,
} from 'docx';

const require = createRequire(import.meta.url);
const pdfjsMjsPath = require.resolve('pdfjs-dist/legacy/build/pdf.mjs');
const standardFontsPath = path.join(path.dirname(pdfjsMjsPath), '../standard_fonts/').replace(/\\/g, '/');

class NodeCanvasFactory {
  create(width, height) {
    const canvas = createCanvas(width, height);
    const context = canvas.getContext('2d');
    return { canvas, context };
  }
  reset(canvasAndContext, width, height) {
    if (canvasAndContext && canvasAndContext.canvas) {
      canvasAndContext.canvas.width = width;
      canvasAndContext.canvas.height = height;
    }
  }
  destroy(canvasAndContext) {
    if (canvasAndContext && canvasAndContext.canvas) {
      canvasAndContext.canvas.width = 0;
      canvasAndContext.canvas.height = 0;
      canvasAndContext.canvas = null;
      canvasAndContext.context = null;
    }
  }
}

/**
 * Safely converts an image object (RGBA, RGB, or Grayscale) to a PNG Buffer.
 */
function convertImgObjToPngBuffer(imgObj) {
  if (!imgObj || !imgObj.width || !imgObj.height || !imgObj.data) return null;
  const w = imgObj.width;
  const h = imgObj.height;
  if (w <= 0 || h <= 0) return null;

  try {
    const canvas = createCanvas(w, h);
    const ctx = canvas.getContext('2d');
    const imgData = ctx.createImageData(w, h);
    const src = imgObj.data;
    const dst = imgData.data;

    if (src.length === w * h * 4) {
      dst.set(src);
    } else if (src.length === w * h * 3) {
      for (let p = 0, q = 0; p < src.length; p += 3, q += 4) {
        dst[q] = src[p];
        dst[q + 1] = src[p + 1];
        dst[q + 2] = src[p + 2];
        dst[q + 3] = 255;
      }
    } else if (src.length === w * h) {
      for (let p = 0, q = 0; p < src.length; p++, q += 4) {
        const v = src[p];
        dst[q] = v;
        dst[q + 1] = v;
        dst[q + 2] = v;
        dst[q + 3] = 255;
      }
    } else {
      for (let p = 0, q = 0; p < src.length && q < dst.length; p++, q++) {
        dst[q] = src[p];
      }
    }

    ctx.putImageData(imgData, 0, 0);
    return canvas.toBuffer('image/png');
  } catch (err) {
    console.warn('Image conversion to PNG failed:', err.message);
    return null;
  }
}

/**
 * Extracts embedded image objects safely with EXACT relative position and proportions.
 */
async function extractPageImages(page, viewport) {
  const images = [];
  try {
    const opList = await page.getOperatorList();
    const fnArray = opList.fnArray;
    const argsArray = opList.argsArray;

    const printableWidthDocxPt = 468; // Standard 8.5" page with 0.75" margins = 468 pt printable

    for (let i = 0; i < fnArray.length; i++) {
      const fn = fnArray[i];
      const args = argsArray[i];

      if (
        fn === pdfjsLib.OPS.paintImageXObject ||
        fn === pdfjsLib.OPS.paintInlineImageXObject ||
        fn === pdfjsLib.OPS.paintImageMaskXObject
      ) {
        const imgTarget = args[0];
        if (!imgTarget) continue;

        let transform = [100, 0, 0, 100, 0, 0];
        for (let j = i - 1; j >= Math.max(0, i - 8); j--) {
          if (fnArray[j] === pdfjsLib.OPS.transform) {
            transform = argsArray[j];
            break;
          }
        }

        const scaleX = Math.abs(transform[0]);
        const scaleY = Math.abs(transform[3]);
        const xPt = transform[4] || 0;
        const pdfY = transform[5] || 0;
        const yPt = Math.max(0, viewport.height - pdfY - (scaleY || 50));

        let imgObj = null;
        if (typeof imgTarget === 'object' && imgTarget !== null) {
          imgObj = imgTarget;
        } else if (typeof imgTarget === 'string') {
          try {
            if (page.objs.has(imgTarget)) {
              imgObj = page.objs.get(imgTarget);
            } else if (page.commonObjs && page.commonObjs.has(imgTarget)) {
              imgObj = page.commonObjs.get(imgTarget);
            } else {
              imgObj = await new Promise((resolve) => {
                page.objs.get(imgTarget, (res) => resolve(res));
              });
            }
          } catch (e) {
            console.warn(`Could not load image ${imgTarget}:`, e.message);
          }
        }

        if (imgObj) {
          const pngBuffer = convertImgObjToPngBuffer(imgObj);
          if (pngBuffer && pngBuffer.length > 0) {
            const nativeW = imgObj.width || scaleX || 100;
            const nativeH = imgObj.height || scaleY || 100;
            const aspectRatio = (nativeW > 0 && nativeH > 0) ? (nativeW / nativeH) : (scaleX / scaleY || 1.0);

            // Proportional width relative to page viewport
            const pageRatio = viewport.width > 0 ? (scaleX > 0 ? scaleX / viewport.width : nativeW / viewport.width) : 0.8;
            let displayWidth = Math.max(24, Math.min(printableWidthDocxPt, Math.round(pageRatio * printableWidthDocxPt)));

            // If scaleX was large (e.g. hero banner), make it full printable width
            if (scaleX >= viewport.width * 0.6) {
              displayWidth = printableWidthDocxPt;
            }

            let displayHeight = Math.round(displayWidth / aspectRatio);

            if (displayHeight > 500) {
              displayHeight = 500;
              displayWidth = Math.round(displayHeight * aspectRatio);
            }

            images.push({
              type: 'image',
              buffer: pngBuffer,
              widthPt: displayWidth,
              heightPt: displayHeight,
              xPt,
              yPt,
            });
          }
        }
      }
    }
  } catch (err) {
    console.warn('Error extracting page images:', err.message);
  }
  return images;
}

/**
 * Parses raw PDF text items into positioned text tokens.
 */
function parseTextItems(textContent, viewport) {
  const items = [];

  for (const item of textContent.items) {
    if (!item.str || item.str.length === 0) continue;

    const transform = item.transform || [1, 0, 0, 1, 0, 0];
    const scaleX = Math.abs(transform[0]);
    const scaleY = Math.abs(transform[3]);
    const fontSizePt = Math.max(8, Math.round(scaleY || scaleX || 11));

    const xPt = transform[4];
    const pdfY = transform[5];
    const yPt = viewport.height - pdfY;

    const fontName = (item.fontName || '').toLowerCase();
    const styleObj = textContent.styles ? textContent.styles[item.fontName] : null;
    const fontFamily = styleObj ? styleObj.fontFamily || 'Calibri' : 'Calibri';

    const isBold =
      fontName.includes('bold') ||
      fontName.includes('heavy') ||
      fontName.includes('black') ||
      fontName.includes('bd') ||
      (styleObj && styleObj.isBold);

    const isItalic =
      fontName.includes('italic') ||
      fontName.includes('oblique') ||
      fontName.includes('it') ||
      (styleObj && styleObj.isItalic);

    let color = '1F2937';

    items.push({
      str: item.str,
      xPt,
      yPt,
      widthPt: item.width || 0,
      heightPt: item.height || fontSizePt,
      fontSizePt,
      fontFamily,
      isBold,
      isItalic,
      color,
    });
  }

  return items;
}

/**
 * Clusters text items into horizontal line bands based on Y coordinates.
 */
function clusterTextItemsIntoLines(items) {
  if (items.length === 0) return [];

  items.sort((a, b) => a.yPt - b.yPt);

  const clusters = [];

  for (const item of items) {
    let matchedCluster = null;
    for (const cluster of clusters) {
      if (Math.abs(item.yPt - cluster.avgY) <= Math.max(6, item.fontSizePt * 0.65)) {
        matchedCluster = cluster;
        break;
      }
    }

    if (matchedCluster) {
      matchedCluster.items.push(item);
      matchedCluster.avgY =
        matchedCluster.items.reduce((sum, it) => sum + it.yPt, 0) / matchedCluster.items.length;
    } else {
      clusters.push({
        avgY: item.yPt,
        items: [item],
      });
    }
  }

  clusters.sort((a, b) => a.avgY - b.avgY);

  for (const cluster of clusters) {
    cluster.items.sort((a, b) => a.xPt - b.xPt);
  }

  return clusters;
}

/**
 * Merges adjacent text items on the same line into coherent text runs.
 */
function mergeLineItems(lineItems) {
  if (lineItems.length === 0) return [];

  const mergedRuns = [];
  let currentRun = null;

  for (const item of lineItems) {
    if (!currentRun) {
      currentRun = { ...item };
    } else {
      const prevRightX = currentRun.xPt + currentRun.widthPt;
      const gap = item.xPt - prevRightX;

      if (
        item.fontFamily === currentRun.fontFamily &&
        item.fontSizePt === currentRun.fontSizePt &&
        item.isBold === currentRun.isBold &&
        item.isItalic === currentRun.isItalic &&
        item.color === currentRun.color
      ) {
        if (gap > 2) {
          currentRun.str += ' ' + item.str;
        } else {
          currentRun.str += item.str;
        }
        currentRun.widthPt = Math.max(currentRun.widthPt, item.xPt + item.widthPt - currentRun.xPt);
      } else {
        mergedRuns.push(currentRun);
        currentRun = { ...item };
      }
    }
  }

  if (currentRun) {
    mergedRuns.push(currentRun);
  }

  return mergedRuns;
}

/**
 * Converts line clusters into Word Paragraphs.
 */
/**
 * Converts line clusters into Word Paragraphs with smart horizontal alignment and spacing.
 */
function buildParagraphsFromLineClusters(clusters, forcedAlignment = null, viewportWidth = 612) {
  const paragraphs = [];

  for (const cluster of clusters) {
    const lineRuns = mergeLineItems(cluster.items);
    if (lineRuns.length === 0) continue;

    const minX = Math.min(...cluster.items.map((it) => it.xPt));
    const maxX = Math.max(...cluster.items.map((it) => it.xPt + (it.widthPt || 0)));
    const lineWidth = maxX - minX;
    const lineMidX = (minX + maxX) / 2;
    const pageMidX = viewportWidth / 2;

    let alignment = forcedAlignment || AlignmentType.LEFT;
    if (!forcedAlignment) {
      if (Math.abs(lineMidX - pageMidX) < 45 && lineWidth < viewportWidth * 0.75) {
        alignment = AlignmentType.CENTER;
      }
    }

    const docxRuns = [];

    for (let i = 0; i < lineRuns.length; i++) {
      const run = lineRuns[i];
      if (i > 0) {
        const prevRun = lineRuns[i - 1];
        const gap = run.xPt - (prevRun.xPt + prevRun.widthPt);
        if (gap > 35) {
          docxRuns.push(new TextRun({ text: '\t\t\t\t' }));
        } else {
          docxRuns.push(new TextRun({ text: ' ' }));
        }
      }

      const halfPoints = Math.min(96, Math.max(16, run.fontSizePt * 2));

      docxRuns.push(
        new TextRun({
          text: run.str,
          font: run.fontFamily,
          size: halfPoints,
          bold: run.isBold,
          italic: run.isItalic,
          color: run.color,
        })
      );
    }

    if (docxRuns.length > 0) {
      const firstRun = lineRuns[0];
      const isTitle = firstRun.fontSizePt >= 18;

      paragraphs.push(
        new Paragraph({
          children: docxRuns,
          alignment: alignment,
          spacing: {
            before: isTitle ? 180 : 40,
            after: isTitle ? 120 : 60,
            line: 276,
          },
        })
      );
    }
  }

  return paragraphs;
}

/**
 * Detects if a group of line clusters represents a multi-column section.
 */
function isMultiColumnSection(clusters, pageMidX) {
  if (clusters.length < 2) return false;

  let leftCount = 0;
  let rightCount = 0;

  for (const cluster of clusters) {
    const minX = Math.min(...cluster.items.map((it) => it.xPt));
    const maxX = Math.max(...cluster.items.map((it) => it.xPt + it.widthPt));

    if (maxX <= pageMidX + 30) {
      leftCount++;
    } else if (minX >= pageMidX - 30) {
      rightCount++;
    }
  }

  return leftCount >= 2 && rightCount >= 2;
}

/**
 * Creates an exact visual replica Word (.docx) document from a PDF file.
 * Each page is rendered at high resolution to preserve 100% of original formatting, fonts, images, tables, logos, and layout ("Same to Same").
 */
export const createDocxExactReplica = async (pdfPath) => {
  const dataBuffer = new Uint8Array(await fs.promises.readFile(pdfPath));
  const canvasFactory = new NodeCanvasFactory();
  const loadingTask = pdfjsLib.getDocument({
    data: dataBuffer,
    canvasFactory,
    standardFontDataUrl: standardFontsPath,
    disableFontFace: false,
    nativeImageDecoderSupport: 'none',
  });

  const pdfDoc = await loadingTask.promise;
  const numPages = pdfDoc.numPages;
  const sections = [];

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const baseViewport = page.getViewport({ scale: 1.0 });
    const widthPt = baseViewport.width;
    const heightPt = baseViewport.height;

    // Render at 2.5 scale (180 DPI) for crisp, pixel-perfect text & image reproduction
    const renderScale = 2.5;
    const renderViewport = page.getViewport({ scale: renderScale });

    const canvasAndContext = canvasFactory.create(renderViewport.width, renderViewport.height);
    const renderContext = {
      canvasContext: canvasAndContext.context,
      viewport: renderViewport,
    };

    await page.render(renderContext).promise;
    const pngBuffer = canvasAndContext.canvas.toBuffer('image/png');
    canvasFactory.destroy(canvasAndContext);

    // Convert PDF points to dxa (1 pt = 20 dxa)
    const widthDxa = Math.round(widthPt * 20);
    const heightDxa = Math.round(heightPt * 20);

    sections.push({
      properties: {
        page: {
          size: {
            width: widthDxa,
            height: heightDxa,
          },
          margin: {
            top: 0,
            bottom: 0,
            left: 0,
            right: 0,
          },
        },
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new ImageRun({
              data: pngBuffer,
              transformation: {
                width: Math.round(widthPt * (96 / 72)),
                height: Math.round(heightPt * (96 / 72)),
              },
            }),
          ],
          spacing: { before: 0, after: 0, line: 240 },
        }),
      ],
    });
  }

  const doc = new Document({ sections });
  return await Packer.toBuffer(doc);
};

/**
 * Renders page content into Word sections with automatic column and section detection.
 */
function renderPageContentToDocx(lineClusters, pageImages, viewport) {
  const docxChildren = [];
  const pageMidX = viewport.width / 2;

  if (lineClusters.length === 0 && pageImages.length === 0) {
    return [
      new Paragraph({
        children: [
          new TextRun({
            text: ' ',
            font: 'Calibri',
            size: 24,
          }),
        ],
      }),
    ];
  }

  // 1. Separate Full-Width Header Clusters from Body Clusters
  const headerClusters = [];
  const bodyClusters = [];

  let isInsideBody = false;

  for (const cluster of lineClusters) {
    const minX = Math.min(...cluster.items.map((it) => it.xPt));
    const maxX = Math.max(...cluster.items.map((it) => it.xPt + (it.widthPt || 0)));
    const lineMidX = (minX + maxX) / 2;
    const lineWidth = maxX - minX;

    const firstStr = (cluster.items[0]?.str || '').trim().toLowerCase();
    const isBodyStartKeyword =
      firstStr.startsWith('abstract') ||
      firstStr.startsWith('1 ') ||
      firstStr.startsWith('1.') ||
      firstStr.startsWith('introduction');

    const isCenteredTitle = Math.abs(lineMidX - pageMidX) < 50 && lineWidth < viewport.width * 0.8;
    const isLargeTitleFont = cluster.items.some((it) => it.fontSizePt >= 16);

    if (isBodyStartKeyword) {
      isInsideBody = true;
    }

    if (!isInsideBody && (isCenteredTitle || isLargeTitleFont || cluster.avgY < 120)) {
      headerClusters.push(cluster);
    } else {
      isInsideBody = true;
      bodyClusters.push(cluster);
    }
  }

  // Render Header Section (Centered Titles & Side-by-side Authors)
  if (headerClusters.length > 0) {
    for (const cluster of headerClusters) {
      const minX = Math.min(...cluster.items.map((it) => it.xPt));
      const maxX = Math.max(...cluster.items.map((it) => it.xPt + (it.widthPt || 0)));
      const lineMidX = (minX + maxX) / 2;
      const lineWidth = maxX - minX;

      const isCentered = Math.abs(lineMidX - pageMidX) < 50 && lineWidth < viewport.width * 0.8;

      if (isCentered) {
        docxChildren.push(...buildParagraphsFromLineClusters([cluster], AlignmentType.CENTER, viewport.width));
      } else {
        docxChildren.push(...buildParagraphsFromLineClusters([cluster], null, viewport.width));
      }
    }
  }

  // 2. Classify Images into Banner Images (Full Width) vs Column Images
  const bannerImages = [];
  const columnImages = [];

  for (const img of pageImages) {
    if (img.widthPt >= viewport.width * 0.45) {
      bannerImages.push(img);
    } else {
      columnImages.push(img);
    }
  }

  // 3. Process Body Section
  let leftBodyCount = 0;
  let rightBodyCount = 0;

  for (const cluster of bodyClusters) {
    const minX = Math.min(...cluster.items.map((it) => it.xPt));
    const maxX = Math.max(...cluster.items.map((it) => it.xPt + (it.widthPt || 0)));

    if (maxX <= pageMidX + 30) {
      leftBodyCount++;
    } else if (minX >= pageMidX - 30) {
      rightBodyCount++;
    }
  }

  for (const img of columnImages) {
    if (img.xPt >= pageMidX - 30) {
      rightBodyCount++;
    } else {
      leftBodyCount++;
    }
  }

  const isTrueDualColumn = leftBodyCount >= 2 && rightBodyCount >= 1;

  if (isTrueDualColumn) {
    const leftColClusters = [];
    const rightColClusters = [];
    const leftColImages = [];
    const rightColImages = [];

    for (const cluster of bodyClusters) {
      const minX = Math.min(...cluster.items.map((it) => it.xPt));
      const maxX = Math.max(...cluster.items.map((it) => it.xPt + (it.widthPt || 0)));

      if (minX < pageMidX - 20 && maxX > pageMidX + 20) {
        const leftItems = cluster.items.filter((it) => it.xPt < pageMidX);
        const rightItems = cluster.items.filter((it) => it.xPt >= pageMidX);
        if (leftItems.length > 0) leftColClusters.push({ avgY: cluster.avgY, items: leftItems });
        if (rightItems.length > 0) rightColClusters.push({ avgY: cluster.avgY, items: rightItems });
      } else if (maxX <= pageMidX + 30) {
        leftColClusters.push(cluster);
      } else {
        rightColClusters.push(cluster);
      }
    }

    for (const img of columnImages) {
      if (img.xPt >= pageMidX - 30) {
        rightColImages.push(img);
      } else {
        leftColImages.push(img);
      }
    }

    const leftElements = [
      ...leftColClusters.map((c) => ({ type: 'cluster', yPt: c.avgY, cluster: c })),
      ...leftColImages.map((i) => ({ type: 'image', yPt: i.yPt, image: i })),
    ].sort((a, b) => a.yPt - b.yPt);

    const leftChildren = [];
    for (const elem of leftElements) {
      if (elem.type === 'cluster') {
        leftChildren.push(...buildParagraphsFromLineClusters([elem.cluster], AlignmentType.LEFT, viewport.width));
      } else if (elem.type === 'image') {
        leftChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: elem.image.buffer,
                transformation: { width: elem.image.widthPt, height: elem.image.heightPt },
              }),
            ],
            spacing: { before: 100, after: 100 },
          })
        );
      }
    }

    const rightElements = [
      ...rightColClusters.map((c) => ({ type: 'cluster', yPt: c.avgY, cluster: c })),
      ...rightColImages.map((i) => ({ type: 'image', yPt: i.yPt, image: i })),
    ].sort((a, b) => a.yPt - b.yPt);

    const rightChildren = [];
    for (const elem of rightElements) {
      if (elem.type === 'cluster') {
        rightChildren.push(...buildParagraphsFromLineClusters([elem.cluster], AlignmentType.LEFT, viewport.width));
      } else if (elem.type === 'image') {
        rightChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: elem.image.buffer,
                transformation: { width: elem.image.widthPt, height: elem.image.heightPt },
              }),
            ],
            spacing: { before: 100, after: 100 },
          })
        );
      }
    }

    const borderlessCell = {
      top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    };

    const table = new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [4450, 460, 4450],
      borders: borderlessCell,
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 4450, type: WidthType.DXA },
              borders: borderlessCell,
              children: leftChildren.length > 0 ? leftChildren : [new Paragraph({ children: [] })],
            }),
            new TableCell({
              width: { size: 460, type: WidthType.DXA },
              borders: borderlessCell,
              children: [new Paragraph({ children: [] })],
            }),
            new TableCell({
              width: { size: 4450, type: WidthType.DXA },
              borders: borderlessCell,
              children: rightChildren.length > 0 ? rightChildren : [new Paragraph({ children: [] })],
            }),
          ],
        }),
      ],
    });

    const pageStream = [
      { type: 'table', yPt: 100, table },
      ...bannerImages.map((img) => ({ type: 'banner_image', yPt: img.yPt, image: img })),
    ].sort((a, b) => a.yPt - b.yPt);

    for (const item of pageStream) {
      if (item.type === 'table') {
        docxChildren.push(item.table);
      } else if (item.type === 'banner_image') {
        docxChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: item.image.buffer,
                transformation: { width: item.image.widthPt, height: item.image.heightPt },
              }),
            ],
            spacing: { before: 120, after: 120 },
          })
        );
      }
    }
  } else {
    // Single Column Body Section Across 100% Page Width (No Squeezing!)
    const singleElements = [
      ...bodyClusters.map((c) => ({ type: 'cluster', yPt: c.avgY, cluster: c })),
      ...pageImages.map((i) => ({ type: 'image', yPt: i.yPt, image: i })),
    ].sort((a, b) => a.yPt - b.yPt);

    for (const elem of singleElements) {
      if (elem.type === 'cluster') {
        docxChildren.push(...buildParagraphsFromLineClusters([elem.cluster], null, viewport.width));
      } else if (elem.type === 'image') {
        docxChildren.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                data: elem.image.buffer,
                transformation: { width: elem.image.widthPt, height: elem.image.heightPt },
              }),
            ],
            spacing: { before: 120, after: 120 },
          })
        );
      }
    }
  }

  return docxChildren;
}

/**
 * Creates a Microsoft Word (.docx) file from PDF.
 * @param {string} pdfPath - Path to input PDF file.
 * @param {'exact'|'editable'} [mode='editable'] - Conversion mode ('exact' for 1:1 replica, 'editable' for flow text).
 */
export const createDocxFromPdf = async (pdfPath, mode = 'editable') => {
  if (mode === 'exact') {
    return await createDocxExactReplica(pdfPath);
  }

  const dataBuffer = new Uint8Array(await fs.promises.readFile(pdfPath));
  const loadingTask = pdfjsLib.getDocument({
    data: dataBuffer,
    canvasFactory: new NodeCanvasFactory(),
    standardFontDataUrl: standardFontsPath,
    disableFontFace: false,
    nativeImageDecoderSupport: 'none',
  });

  const pdfDoc = await loadingTask.promise;
  const numPages = pdfDoc.numPages;
  const sections = [];
  let totalExtractedChars = 0;

  for (let pageNum = 1; pageNum <= numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1.0 });

    const textContent = await page.getTextContent({ includeMarkedContent: true });
    const textItems = parseTextItems(textContent, viewport);
    totalExtractedChars += textItems.reduce((sum, item) => sum + (item.str ? item.str.replace(/[^a-zA-Z0-9]/g, '').length : 0), 0);
    const pageImages = await extractPageImages(page, viewport);

    const lineClusters = clusterTextItemsIntoLines(textItems);
    const docxChildren = renderPageContentToDocx(lineClusters, pageImages, viewport);

    sections.push({
      properties: {
        page: {
          margin: {
            top: 720,   // 0.5 inch margins
            bottom: 720,
            left: 720,
            right: 720,
          },
        },
      },
      children: docxChildren,
    });
  }

  if (totalExtractedChars < 15) {
    const err = new Error('PDF contains little or no extractable text layer.');
    err.isScannedPdf = true;
    throw err;
  }

  const doc = new Document({ sections });
  return await Packer.toBuffer(doc);
};

/**
 * Creates an editable Microsoft Word (.docx) file from raw text string.
 */
export const createDocxFromText = async (text) => {
  const lines = text.split(/\r?\n/);
  const docxChildren = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.length > 0) {
      docxChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: trimmed,
              font: 'Calibri',
              size: 24,
              color: '1F2937',
            }),
          ],
          spacing: {
            after: 140,
            line: 276,
          },
        })
      );
    }
  }

  if (docxChildren.length === 0) {
    docxChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: 'Empty document.',
            font: 'Calibri',
            size: 24,
          }),
        ],
      })
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: docxChildren,
      },
    ],
  });

  return await Packer.toBuffer(doc);
};
