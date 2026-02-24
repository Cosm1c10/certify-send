import * as pdfjsLib from 'pdfjs-dist';

// Configure the worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

/**
 * Converts the first 2 pages of a PDF file to a single stitched Base64 JPEG image.
 * Pages are stacked vertically so the AI can read content from both pages.
 *
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

  const maxPages = Math.min(pdf.numPages, 2);
  const pageCanvases: { canvas: HTMLCanvasElement; width: number; height: number }[] = [];

  // Render each page to its own temporary canvas
  for (let i = 1; i <= maxPages; i++) {
    const page = await pdf.getPage(i);
    const viewport = page.getViewport({ scale });

    const tempCanvas = document.createElement('canvas');
    const tempContext = tempCanvas.getContext('2d');

    if (!tempContext) {
      throw new Error(`Failed to get canvas 2D context for page ${i}`);
    }

    tempCanvas.width = viewport.width;
    tempCanvas.height = viewport.height;

    await page.render({
      canvasContext: tempContext,
      viewport: viewport,
      canvas: tempCanvas,
    } as any).promise;

    pageCanvases.push({
      canvas: tempCanvas,
      width: viewport.width,
      height: viewport.height,
    });
  }

  // Create the combined canvas
  const combinedWidth = Math.max(...pageCanvases.map(p => p.width));
  const combinedHeight = pageCanvases.reduce((sum, p) => sum + p.height, 0);

  const mainCanvas = document.createElement('canvas');
  const mainContext = mainCanvas.getContext('2d');

  if (!mainContext) {
    throw new Error('Failed to get canvas 2D context for combined image');
  }

  mainCanvas.width = combinedWidth;
  mainCanvas.height = combinedHeight;

  // Fill with white background
  mainContext.fillStyle = '#FFFFFF';
  mainContext.fillRect(0, 0, combinedWidth, combinedHeight);

  // Draw each page canvas stacked vertically
  let yOffset = 0;
  for (const { canvas, height } of pageCanvases) {
    mainContext.drawImage(canvas, 0, yOffset);
    yOffset += height;
  }

  // Convert to JPEG and return Base64 (without the data URL prefix)
  const dataUrl = mainCanvas.toDataURL('image/jpeg', quality);
  const base64 = dataUrl.replace(/^data:image\/jpeg;base64,/, '');

  return base64;
}

/**
 * Converts the first 2 pages of a PDF file to a Base64 data URL
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
