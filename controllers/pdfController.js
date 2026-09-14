import path from 'path';
import { isPdfFile } from '../utils/fileValidation.js';
import { extractTextFromPdf } from '../services/pdfService.js';
import { createDocxFromText } from '../services/docxService.js';
import { removeFile } from '../utils/cleanup.js';

/**
 * Controller to handle PDF to Word (.docx) conversion.
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

    // 2. Extract text content from PDF
    const { text } = await extractTextFromPdf(file.path);

    // 3. Generate Word (.docx) file buffer
    const docxBuffer = await createDocxFromText(text);

    // 4. Formulate clean download filename
    const originalBasename = path.basename(file.originalname, path.extname(file.originalname));
    const safeBasename = originalBasename.replace(/[^a-zA-Z0-9_\- ]/g, '_') || 'converted';
    const downloadFilename = `${safeBasename}.docx`;

    // 5. Set response headers for file download
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${downloadFilename}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

    // 6. Send generated document
    return res.send(docxBuffer);
  } catch (error) {
    next(error);
  } finally {
    // Always cleanup temporary uploaded file
    await removeFile(file.path);
  }
};
