import multer from 'multer';
import path from 'path';
import fs from 'fs';
import dotenv from 'dotenv';

dotenv.config();

const uploadDir = path.resolve('uploads');

// Ensure upload directory exists
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const maxFileSize = parseInt(process.env.MAX_FILE_SIZE || '10485760', 10); // Default 10MB

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, `pdf-${uniqueSuffix}.pdf`);
  },
});

const fileFilter = (req, file, cb) => {
  const fileExt = path.extname(file.originalname).toLowerCase();
  const mimeType = file.mimetype;

  if (fileExt === '.pdf' || mimeType === 'application/pdf') {
    cb(null, true);
  } else {
    const error = new Error('Only PDF files (.pdf) are allowed.');
    error.status = 400;
    cb(error, false);
  }
};

export const uploadSinglePdf = multer({
  storage,
  limits: {
    fileSize: maxFileSize,
  },
  fileFilter,
}).single('pdf');

/**
 * Express wrapper middleware to handle Multer errors cleanly.
 */
export const handleUpload = (req, res, next) => {
  uploadSinglePdf(req, res, (err) => {
    if (err) {
      if (err instanceof multer.MulterError) {
        if (err.code === 'LIMIT_FILE_SIZE') {
          const sizeMB = (maxFileSize / (1024 * 1024)).toFixed(0);
          return res.status(400).json({
            success: false,
            message: `File size exceeds the limit of ${sizeMB}MB.`,
          });
        }
        return res.status(400).json({
          success: false,
          message: `File upload error: ${err.message}`,
        });
      }
      return res.status(err.status || 400).json({
        success: false,
        message: err.message || 'Invalid file uploaded.',
      });
    }
    next();
  });
};
