import { fileTypeFromFile } from 'file-type';
import fs from 'fs';

/**
 * Validates whether the given file is genuinely a PDF file using magic bytes.
 * @param {string} filePath 
 * @returns {Promise<boolean>}
 */
export const isPdfFile = async (filePath) => {
  try {
    const fileType = await fileTypeFromFile(filePath);
    if (fileType && fileType.mime === 'application/pdf') {
      return true;
    }

    // Fallback magic byte check for standard PDF header %PDF-
    const fd = await fs.promises.open(filePath, 'r');
    const buffer = Buffer.alloc(5);
    await fd.read(buffer, 0, 5, 0);
    await fd.close();

    const header = buffer.toString('utf-8');
    return header.startsWith('%PDF-');
  } catch (error) {
    console.error('PDF header validation error:', error);
    return false;
  }
};
