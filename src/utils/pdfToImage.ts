import * as pdfjsLib from 'pdfjs-dist';

// Configure the worker
pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

// =============================================================================
// SMART SCALING STRATEGY
// Keeps total pixel count within AI vision limits while maintaining readability
// =============================================================================
const MAX_PAGES = 5;
const SCALE_HIGH = 2.0;   // 1-2 pages: full quality
const SCALE_LOW = 1.5;    // 3-5 pages: reduced to keep total size manageable
const JPEG_QUALITY = 0.8; // Optimized for API payload size

/**
 * Converts up to the first 5 pages of a PDF into a single stitched JPEG image.
 * Uses Smart Scaling: 2.0x for short docs, 1.5x for longer ones.
 * Pages are stacked vertically so the AI can read the full document.
 *
 * @param file - The PDF file to convert
 * @returns Base64 encoded JPEG string (without data URL prefix)
 */
export async function pdfToImage(file: File): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  const pageCount = Math.min(pdf.numPages, MAX_PAGES);
  const scale = pageCount <= 2 ? SCALE_HIGH : SCALE_LOW;

  console.log(`PDF "${file.name}": ${pdf.numPages} total pages, rendering ${pageCount} at scale ${scale}`);

  const pageCanvases: { canvas: HTMLCanvasElement; width: number; height: number }[] = [];

  // Render each page to its own temporary canvas
  for (let i = 1; i <= pageCount; i++) {
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
  const dataUrl = mainCanvas.toDataURL('image/jpeg', JPEG_QUALITY);
  const base64 = dataUrl.replace(/^data:image\/jpeg;base64,/, '');

  console.log(`PDF "${file.name}": stitched ${pageCount} pages -> ${(base64.length / 1024).toFixed(0)}KB base64`);

  return base64;
}

/**
 * Converts up to the first 5 pages of a PDF file to a Base64 data URL
 * @param file - The PDF file to convert
 * @returns Full data URL string (e.g., "data:image/jpeg;base64,...")
 */
export async function pdfToImageDataUrl(file: File): Promise<string> {
  const base64 = await pdfToImage(file);
  return `data:image/jpeg;base64,${base64}`;
}
