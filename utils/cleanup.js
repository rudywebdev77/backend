import fs from 'fs';
import path from 'path';

/**
 * Safely removes a file from the filesystem if it exists.
 * @param {string} filePath - Absolute or relative path to the file.
 */
export const removeFile = async (filePath) => {
  if (!filePath) return;
  try {
    const resolvedPath = path.resolve(filePath);
    if (fs.existsSync(resolvedPath)) {
      await fs.promises.unlink(resolvedPath);
    }
  } catch (error) {
    console.error(`Failed to delete temporary file at ${filePath}:`, error.message);
  }
};
