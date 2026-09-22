import path from 'path';
import { isPdfFile } from '../utils/fileValidation.js';
import { extractTextFromPdf } from '../services/pdfService.js';
import { extractTextWithOcr } from '../services/ocrService.js';
import { createDocxFromText } from '../services/docxService.js';
import { removeFile } from '../utils/cleanup.js';

/**
 * Controller to handle PDF to Word (.docx) conversion.
 * Supports both normal text-based PDFs and scanned image-based PDFs (via automatic OCR fallback).
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

    let text = '';
    let conversionMethod = 'direct';

    // 2. Attempt standard text extraction first
    try {
      const extracted = await extractTextFromPdf(file.path);
      text = extracted.text;
    } catch (parseError) {
      // If detected as a scanned PDF, automatically fall back to OCR
      if (parseError.isScannedPdf) {
        console.log('PDF has little/no extractable text. Falling back to OCR processing...');
        const ocrResult = await extractTextWithOcr(file.path);
        text = ocrResult.text;
        conversionMethod = 'ocr';
      } else {
        throw parseError;
      }
    }

    // 3. Generate Word (.docx) file buffer from extracted text
    const docxBuffer = await createDocxFromText(text);

    // 4. Formulate clean download filename
    const originalBasename = path.basename(file.originalname, path.extname(file.originalname));
    const safeBasename = originalBasename.replace(/[^a-zA-Z0-9_\- ]/g, '_') || 'converted';
    const downloadFilename = `${safeBasename}.docx`;

    // 5. Set response headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${downloadFilename}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition, X-Conversion-Method');
    res.setHeader('X-Conversion-Method', conversionMethod);

    // 6. Send generated document
    return res.send(docxBuffer);
  } catch (error) {
    next(error);
  } finally {
    // Always cleanup temporary uploaded file
    await removeFile(file.path);
  }
};
