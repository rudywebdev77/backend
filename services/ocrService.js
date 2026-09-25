import fs from 'fs';
import path from 'path';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';
import { PNG } from 'pngjs';
import { createWorker } from 'tesseract.js';

// Default maximum pages allowed for OCR processing per PDF
const DEFAULT_MAX_OCR_PAGES = 20;

/**
 * Perform Optical Character Recognition (OCR) on a scanned/image-based PDF.
 @param {string} filePath - Path to the uploaded PDF file.
  @param {Object} [options={}] - OCR configuration options.
  @param {number} [options.maxPages] - Maximum allowed pages for OCR.
  @returns {Promise<{ text: string, pageCount: number, processedPages: number }>}
 */
export const extractTextWithOcr = async (filePath, options = {}) => {
  const maxPages = options.maxPages || parseInt(process.env.MAX_OCR_PAGES, 10) || DEFAULT_MAX_OCR_PAGES;

  const dataBuffer = await fs.promises.readFile(filePath);
  const uint8Array = new Uint8Array(dataBuffer);

  // Load PDF document using pdfjs-dist in legacy mode without font/eval worker dependencies
  const loadingTask = pdfjsLib.getDocument({
    data: uint8Array,
    isEvalSupported: false,
    useSystemFonts: true,
    disableFontFace: true,
  });

  const pdfDocument = await loadingTask.promise;
  const numPages = pdfDocument.numPages;

  if (numPages === 0) {
    const err = new Error('The uploaded PDF contains no pages.');
    err.status = 400;
    throw err;
  }

  if (numPages > maxPages) {
    const limitErr = new Error(
      `This scanned document has ${numPages} pages, which exceeds the maximum allowed limit of ${maxPages} pages for OCR processing. Please split or upload a smaller PDF.`
    );
    limitErr.status = 400;
    throw limitErr;
  }

  let worker = null;
  const extractedPagesText = [];
  let successfulPagesCount = 0;

  try {
    // Initialize Tesseract worker with English language data and cache directory in node_modules/.cache
    const cachePath = path.join(process.cwd(), 'node_modules', '.cache', 'tesseract');
    worker = await createWorker('eng', 1, { cachePath });

    for (let pageNum = 1; pageNum <= numPages; pageNum++) {
      try {
        const page = await pdfDocument.getPage(pageNum);
        const ops = await page.getOperatorList();
        const pageImageTexts = [];

        for (let i = 0; i < ops.fnArray.length; i++) {
          const fn = ops.fnArray[i];
          if (
            fn === pdfjsLib.OPS.paintImageXObject ||
            fn === pdfjsLib.OPS.paintJpegXObject ||
            fn === pdfjsLib.OPS.paintInlineImageXObject
          ) {
            const imgName = ops.argsArray[i][0];
            const imgObj = await new Promise((resolve) => {
              page.objs.get(imgName, (obj) => resolve(obj));
            });

            if (imgObj && imgObj.width > 10 && imgObj.height > 10 && imgObj.data) {
              const png = new PNG({ width: imgObj.width, height: imgObj.height });
              const data = imgObj.data;

              if (data.length === imgObj.width * imgObj.height * 3) {
                let srcIdx = 0;
                let dstIdx = 0;
                for (let p = 0; p < imgObj.width * imgObj.height; p++) {
                  png.data[dstIdx] = data[srcIdx];
                  png.data[dstIdx + 1] = data[srcIdx + 1];
                  png.data[dstIdx + 2] = data[srcIdx + 2];
                  png.data[dstIdx + 3] = 255;
                  srcIdx += 3;
                  dstIdx += 4;
                }
              } else if (data.length === imgObj.width * imgObj.height * 4) {
                png.data.set(data);
              }

              const imageBuffer = PNG.sync.write(png);

              const recognitionResult = await worker.recognize(imageBuffer);
              const text = recognitionResult?.data?.text ? recognitionResult.data.text.trim() : '';
              if (text.length > 0) {
                pageImageTexts.push(text);
              }
            }
          }
        }

        if (pageImageTexts.length > 0) {
          extractedPagesText.push(`--- Page ${pageNum} ---\n${pageImageTexts.join('\n')}`);
          successfulPagesCount++;
        } else {
          extractedPagesText.push(`--- Page ${pageNum} ---\n[No text recognized on this page]`);
        }
      } catch (pageError) {
        console.warn(`OCR warning: Failed to process page ${pageNum}:`, pageError.message);
        extractedPagesText.push(`--- Page ${pageNum} ---\n[Page could not be processed]`);
      }
    }
  } finally {
    if (worker) {
      try {
        await worker.terminate();
      } catch (termErr) {
        console.warn('Tesseract worker termination warning:', termErr.message);
      }
    }
  }

  const combinedText = extractedPagesText.join('\n\n');
  const alphanumericCharCount = combinedText.replace(/[^a-zA-Z0-9]/g, '').length;

  if (alphanumericCharCount === 0 || successfulPagesCount === 0) {
    const err = new Error('OCR was unable to extract any readable text from this scanned PDF file.');
    err.status = 400;
    throw err;
  }

  return {
    text: combinedText,
    pageCount: numPages,
    processedPages: successfulPagesCount,
  };
};
