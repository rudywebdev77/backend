import fs from 'fs';

/**
 * Checks that the file starts with the PDF magic bytes (%PDF-).
 * @param {string} filePath
 * @returns {Promise<boolean>}
 */
export const isPdfFile = async (filePath) => {
  let fd;
  try {
    fd = await fs.promises.open(filePath, 'r');
    const buffer = Buffer.alloc(1024);
    const { bytesRead } = await fd.read(buffer, 0, 1024, 0);
    // The PDF header may be preceded by a few junk bytes, so search the first 1024 bytes
    return buffer.subarray(0, bytesRead).includes('%PDF-');
  } catch (error) {
    console.error('PDF header validation error:', error.message);
    return false;
  } finally {
    if (fd) await fd.close();
  }
};