import * as pdfjsLib from 'pdfjs-dist';

// Configure the worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

/**
 * Converts the first page of a PDF file to a Base64 JPEG image
 * @param file - The PDF file to convert
 * @param scale - Render scale (default: 2 for good quality)
 * @param quality - JPEG quality (0-1, default: 0.85)
 * @returns Base64 encoded JPEG string (without data URL prefix)
 */
export async function pdfToImage(
  file: File,
  scale: number = 2,
  quality: number = 0.85
): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const page = await pdf.getPage(1);

  const viewport = page.getViewport({ scale });

  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');

  if (!context) {
    throw new Error('Failed to get canvas 2D context');
  }

  canvas.height = viewport.height;
  canvas.width = viewport.width;

  await page.render({
    canvasContext: context,
    viewport: viewport,
  }).promise;

  // Convert to JPEG and return Base64 (without the data URL prefix)
  const dataUrl = canvas.toDataURL('image/jpeg', quality);
  const base64 = dataUrl.replace(/^data:image\/jpeg;base64,/, '');

  return base64;
}

/**
 * Converts the first page of a PDF file to a Base64 data URL
 * @param file - The PDF file to convert
 * @param scale - Render scale (default: 2 for good quality)
 * @param quality - JPEG quality (0-1, default: 0.85)
 * @returns Full data URL string (e.g., "data:image/jpeg;base64,...")
 */
export async function pdfToImageDataUrl(
  file: File,
  scale: number = 2,
  quality: number = 0.85
): Promise<string> {
  const base64 = await pdfToImage(file, scale, quality);
  return `data:image/jpeg;base64,${base64}`;
}
