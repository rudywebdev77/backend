import fs from 'fs';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const pdfModule = require('pdf-parse');

/**
 * Extracts text content from a text-based PDF file.
 * @param {string} filePath - Absolute path to the PDF file.
 * @returns {Promise<{text: string, numPages: number}>}
 */
export const extractTextFromPdf = async (filePath) => {
  try {
    const dataBuffer = await fs.promises.readFile(filePath);
    let extractedText = '';
    let numPages = 1;

    if (pdfModule.PDFParse) {
      // pdf-parse v2.x API
      const parser = new pdfModule.PDFParse({ data: dataBuffer });
      const pdfData = await parser.getText();
      extractedText = pdfData.text ? pdfData.text.trim() : '';
      numPages = pdfData.total || (pdfData.pages ? pdfData.pages.length : 1);
      if (typeof parser.destroy === 'function') {
        await parser.destroy();
      }
    } else if (typeof pdfModule === 'function') {
      // pdf-parse v1.x API
      const pdfData = await pdfModule(dataBuffer);
      extractedText = pdfData.text ? pdfData.text.trim() : '';
      numPages = pdfData.numpages || 1;
    } else {
      throw new Error('Unsupported pdf-parse module format.');
    }

    // Clean up page marker lines like "-- 1 of 1 --" added by pdf-parse v2
    const cleanedText = extractedText
      .replace(/--\s*\d+\s*of\s*\d+\s*--/gi, '')
      .trim();

    // Check if the PDF has actual readable alphanumeric content
    const alphanumericCharCount = cleanedText.replace(/[^a-zA-Z0-9]/g, '').length;

    if (!cleanedText || alphanumericCharCount < 15) {
      const error = new Error(
        'This PDF appears to be scanned/image-based with little or no extractable text.'
      );
      error.isScannedPdf = true;
      error.status = 400;
      throw error;
    }

    return {
      text: cleanedText,
      numPages,
    };
  } catch (error) {
    if (error.isScannedPdf) {
      throw error;
    }

    console.error('PDF Parsing Error:', error.message);
    const parseErr = new Error('Could not extract text from the provided PDF file. It may be corrupted, encrypted, or password-protected.');
    parseErr.status = 400;
    throw parseErr;
  }
};
