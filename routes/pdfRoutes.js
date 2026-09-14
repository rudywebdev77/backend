import express from 'express';
import { handleUpload } from '../middleware/uploadMiddleware.js';
import { convertPdfToWord } from '../controllers/pdfController.js';

const router = express.Router();

router.post('/pdf-to-word', handleUpload, convertPdfToWord);

export default router;
