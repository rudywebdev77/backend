import express from 'express';
import rateLimit from 'express-rate-limit';
import { handleUpload } from '../middleware/uploadMiddleware.js';
import { convertPdfToWord } from '../controllers/pdfController.js';

const router = express.Router();

const convertLimiter = rateLimit({
  windowMs: 60 * 60 * 1000, // 1 ghanta
  max: 10, // ek IP se 10 conversions per ghanta
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    success: false,
    message: 'Conversion limit reached. Please try again later.',
  },
});

router.post('/pdf-to-word', convertLimiter, handleUpload, convertPdfToWord);

export default router;