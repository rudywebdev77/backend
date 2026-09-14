import { Document, Paragraph, TextRun, Packer } from 'docx';

/**
 * Creates an editable Microsoft Word (.docx) file buffer from extracted text.
 * @param {string} text - Extracted text content.
 * @returns {Promise<Buffer>} - Generated DOCX file buffer.
 */
export const createDocxFromText = async (text) => {
  const lines = text.split(/\r?\n/);
  const docxChildren = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (trimmed.length > 0) {
      docxChildren.push(
        new Paragraph({
          children: [
            new TextRun({
              text: trimmed,
              font: 'Calibri',
              size: 24, // 12pt font
              color: '1F2937', // Dark gray body text
            }),
          ],
          spacing: {
            after: 140, // 7pt spacing after paragraph
            line: 276,  // 1.15 line spacing
          },
        })
      );
    }
  }

  if (docxChildren.length === 0) {
    docxChildren.push(
      new Paragraph({
        children: [
          new TextRun({
            text: 'Empty document.',
            font: 'Calibri',
            size: 24,
          }),
        ],
      })
    );
  }

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: docxChildren,
      },
    ],
  });

  return await Packer.toBuffer(doc);
};
