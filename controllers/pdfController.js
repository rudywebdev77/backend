import path from 'path';
import { isPdfFile } from '../utils/fileValidation.js';
import { extractTextFromPdf } from '../services/pdfService.js';
import { extractTextWithOcr } from '../services/ocrService.js';
import { createDocxFromPdf, createDocxFromText } from '../services/docxService.js';
import { removeFile } from '../utils/cleanup.js';

/**
 * Controller to handle PDF to Word (.docx) conversion with full layout preservation.
 * Preserves headers, multi-column newsletter/resume structures, icons, images, and visual formatting.
 */
export const convertPdfToWord = async (req, res, next) => {
  const file = req.file;

  if (!file) {
    return res.status(400).json({
      success: false,
      message: 'No PDF file was provided in the request.',
    });
  }

  try {
    // 1. Magic byte file type validation
    const isValidPdf = await isPdfFile(file.path);
    if (!isValidPdf) {
      await removeFile(file.path);
      return res.status(400).json({
        success: false,
        message: 'The uploaded file failed security validation. It is not a valid PDF.',
      });
    }

    const mode = req.query.mode || req.body.mode || 'exact';
    let docxBuffer = null;
    let conversionMethod = mode === 'exact' ? '1-to-1-exact-replica' : 'high-fidelity-layout';

    // 2. High-fidelity conversion (preserves images, headers, multi-columns, icons, and exact layout)
    try {
      docxBuffer = await createDocxFromPdf(file.path, mode);
    } catch (layoutError) {
      console.warn('High-fidelity PDF layout rendering failed, stack trace:', layoutError.stack || layoutError);

      // Fallback: Standard text extraction + OCR fallback
      let text = '';
      try {
        const extracted = await extractTextFromPdf(file.path);
        text = extracted.text;
        conversionMethod = 'text-fallback';
      } catch (parseError) {
        if (parseError.isScannedPdf) {
          console.log('PDF has little/no extractable text. Falling back to OCR processing...');
          const ocrResult = await extractTextWithOcr(file.path);
          text = ocrResult.text;
          conversionMethod = 'ocr-fallback';
        } else {
          throw parseError;
        }
      }

      docxBuffer = await createDocxFromText(text);
    }

    // 3. Formulate clean download filename
    const originalBasename = path.basename(file.originalname, path.extname(file.originalname));
    const safeBasename = originalBasename.replace(/[^a-zA-Z0-9_\- ]/g, '_') || 'converted';
    const downloadFilename = `${safeBasename}.docx`;

    // 4. Set response headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${downloadFilename}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition, X-Conversion-Method');
    res.setHeader('X-Conversion-Method', conversionMethod);

    // 5. Send generated document
    return res.send(docxBuffer);
  } catch (error) {
    next(error);
  } finally {
    // Always cleanup temporary uploaded file
    await removeFile(file.path);
  }
};
