import path from 'path';
import { isPdfFile } from '../utils/fileValidation.js';
import {
  convertPdfToDocxWithAdobe,
  isAdobePdfConfigured,
} from '../services/adobePdfService.js';
import { removeFile } from '../utils/cleanup.js';

export const convertPdfToWord = async (req, res, next) => {
  const file = req.file;

  if (!file) {
    return res.status(400).json({
      success: false,
      message: 'No PDF file was provided in the request.',
    });
  }

  try {
    const isValidPdf = await isPdfFile(file.path);
    if (!isValidPdf) {
      return res.status(400).json({
        success: false,
        message: 'The uploaded file is not a valid PDF.',
      });
    }

    if (!isAdobePdfConfigured()) {
      console.error('[Converter] Adobe credentials missing in .env');
      return res.status(500).json({
        success: false,
        message: 'Conversion service is not configured.',
      });
    }

    let docxBuffer;
    try {
      const start = Date.now();
      docxBuffer = await convertPdfToDocxWithAdobe(file.path);
      console.log(`[Converter] ${file.originalname} converted in ${Date.now() - start}ms`);
    } catch (err) {
      console.error('[Converter] Adobe conversion failed:', err.message);
      return res.status(503).json({
        success: false,
        message: 'Conversion failed right now. Please try again in a moment.',
      });
    }

    const base = path.basename(file.originalname, path.extname(file.originalname));
    const safeName = base.replace(/[^a-zA-Z0-9_\- ]/g, '_') || 'converted';

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    );
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}.docx"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');

    return res.send(docxBuffer);
  } catch (error) {
    next(error);
  } finally {
    await removeFile(file.path);
  }
};