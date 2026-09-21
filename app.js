import express from 'express';
import cors from 'cors';
import helmet from 'helmet';
import rateLimit from 'express-rate-limit';
import dotenv from 'dotenv';

import healthRoutes from './routes/healthRoutes.js';
import pdfRoutes from './routes/pdfRoutes.js';
import { errorHandler } from './middleware/errorHandler.js';

dotenv.config();

const app = express();
const CLIENT_URL = process.env.CLIENT_URL || 'http://localhost:5173';

// Security headers
app.use(
  helmet({
    crossOriginResourcePolicy: { policy: 'cross-origin' },
  })
);

// CORS configuration supporting local dev, Vercel deployments, and custom CLIENT_URL
const corsOptions = {
  origin: (origin, callback) => {
    // Allow non-browser requests or missing origin (same-origin, serverless, curl)
    if (!origin) return callback(null, true);

    // In non-production mode, allow all origins
    if (process.env.NODE_ENV !== 'production') return callback(null, true);

    // Allow explicitly configured CLIENT_URL
    if (CLIENT_URL && origin === CLIENT_URL) return callback(null, true);

    // Allow any Vercel domain (*.vercel.app)
    try {
      const url = new URL(origin);
      if (url.hostname.endsWith('.vercel.app')) {
        return callback(null, true);
      }
    } catch (e) {}

    // Allow local network IP
    if (/^https?:\/\/(192\.168\.\d+\.\d+|10\.\d+\.\d+\.\d+|172\.(1[6-9]|2\d|3[0-1])\.\d+\.\d+)(:\d+)?$/.test(origin)) {
      return callback(null, true);
    }

    callback(new Error(`CORS blocked request from origin: ${origin}`));
  },
  credentials: true,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  exposedHeaders: ['Content-Disposition'],
};

app.use(cors(corsOptions));

// Rate limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 100, // Limit each IP to 100 requests per window
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    success: false,
    message: 'Too many requests from this IP, please try again after 15 minutes.',
  },
});

app.use(limiter);

// Middleware
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// API Routes
app.use('/api', healthRoutes);
app.use('/api', pdfRoutes);

// 404 Route handler
app.use((req, res) => {
  res.status(404).json({
    success: false,
    message: `API endpoint ${req.originalUrl} not found.`,
  });
});

// Centralized error handler
app.use(errorHandler);

export default app;
