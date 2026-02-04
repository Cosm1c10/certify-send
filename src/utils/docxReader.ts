import mammoth from 'mammoth';

/**
 * Extract text content from a Word document (.docx)
 */
export const extractTextFromDocx = async (file: File): Promise<string> => {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value;
};
