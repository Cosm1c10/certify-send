import { useCallback, useRef, useState } from 'react';
import { Upload, Loader2, FileUp, FolderOpen } from 'lucide-react';
import { cn } from '@/lib/utils';
import { extractTextFromDocx } from '@/utils/docxReader';

export interface FileWithBase64 {
  file: File;
  base64Image: string;
  textContent?: string; // For DOCX files
  isDocx?: boolean;
}

interface DropZoneProps {
  onFilesProcess: (files: FileWithBase64[]) => void;
  isProcessing: boolean;
  processingProgress?: { current: number; total: number };
}

const SUPPORTED_IMAGE_TYPES = ['image/jpeg', 'image/jpg', 'image/png'];
const SUPPORTED_PDF_TYPE = 'application/pdf';
const SUPPORTED_DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

// Accepted MIME types for folder-scan filtering
const ACCEPTED_MIME_TYPES = new Set([
  SUPPORTED_PDF_TYPE,
  SUPPORTED_DOCX_TYPE,
  ...SUPPORTED_IMAGE_TYPES,
]);

const DropZone = ({ onFilesProcess, isProcessing, processingProgress }: DropZoneProps) => {
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [conversionProgress, setConversionProgress] = useState({ current: 0, total: 0 });

  // Fallback ref for browsers without showDirectoryPicker (Firefox, older Safari)
  const folderInputRef = useRef<HTMLInputElement>(null);

  // Convert image file directly to base64
  const convertImageToBase64 = useCallback(async (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        resolve(result);
      };
      reader.onerror = () => reject(new Error('Failed to read image file'));
      reader.readAsDataURL(file);
    });
  }, []);

  // Convert PDF to base64 image - extracts ONLY PAGE 1
  const convertPdfToBase64 = useCallback(async (file: File): Promise<string> => {
    const pdfjsLib = await import('pdfjs-dist');
    pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
      'pdfjs-dist/build/pdf.worker.min.mjs',
      import.meta.url
    ).toString();

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    // Extract ONLY page 1
    const page = await pdf.getPage(1);

    const scale = 2;
    const viewport = page.getViewport({ scale });

    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d')!;
    canvas.height = viewport.height;
    canvas.width = viewport.width;

    await page.render({
      canvasContext: context,
      viewport: viewport,
      canvas: canvas,
    } as Parameters<typeof page.render>[0]).promise;

    // JPEG at 0.85 quality is ~70% smaller than PNG with no meaningful OCR loss.
    // This prevents "Failed to send a request to the Edge Function" errors on
    // large/complex PDFs that produce oversized PNG payloads.
    return canvas.toDataURL('image/jpeg', 0.85);
  }, []);

  const processFiles = useCallback(async (files: File[]) => {
    // Guard: Prevent processing if already busy
    if (isProcessing || isConverting) {
      console.warn('Already processing - ignoring new upload');
      return;
    }

    // Filter for supported file types (PDFs, images, and DOCX)
    const supportedFiles = files.filter(file =>
      file.type === SUPPORTED_PDF_TYPE ||
      file.type === SUPPORTED_DOCX_TYPE ||
      SUPPORTED_IMAGE_TYPES.includes(file.type)
    );

    if (supportedFiles.length === 0) {
      alert('Please upload PDF, Word (.docx), or image files (JPG, PNG)');
      return;
    }

    const skippedCount = files.length - supportedFiles.length;
    if (skippedCount > 0) {
      alert(`${skippedCount} unsupported file(s) were skipped`);
    }

    setIsConverting(true);
    setConversionProgress({ current: 0, total: supportedFiles.length });

    const convertedFiles: FileWithBase64[] = [];

    for (let i = 0; i < supportedFiles.length; i++) {
      const file = supportedFiles[i];
      setConversionProgress({ current: i + 1, total: supportedFiles.length });

      try {
        if (file.type === SUPPORTED_DOCX_TYPE) {
          // DOCX file - extract text content
          const textContent = await extractTextFromDocx(file);
          convertedFiles.push({ file, base64Image: '', textContent, isDocx: true });
        } else if (file.type === SUPPORTED_PDF_TYPE) {
          // PDF file - extract page 1 only
          const base64Image = await convertPdfToBase64(file);
          convertedFiles.push({ file, base64Image });
        } else {
          // Image file - convert directly to base64
          const base64Image = await convertImageToBase64(file);
          convertedFiles.push({ file, base64Image });
        }
      } catch (error) {
        console.error(`Error converting file ${file.name}:`, error);
      }
    }

    setIsConverting(false);
    setConversionProgress({ current: 0, total: 0 });

    if (convertedFiles.length > 0) {
      onFilesProcess(convertedFiles);
    } else {
      alert('Failed to convert any files');
    }
  }, [convertPdfToBase64, convertImageToBase64, onFilesProcess]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      processFiles(files);
    }
  }, [processFiles]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      processFiles(Array.from(files));
    }
    e.target.value = '';
  }, [processFiles]);

  // Fallback handler for webkitdirectory input (Firefox / older browsers)
  const handleFolderInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;
    const files = Array.from(fileList).filter(
      f => !f.name.startsWith('.') && ACCEPTED_MIME_TYPES.has(f.type)
    );
    if (files.length > 0) processFiles(files);
    else alert('No supported files found in the selected folder (PDF, DOCX, JPG, PNG)');
    e.target.value = '';
  }, [processFiles]);

  /**
   * Open a folder picker using the modern File System Access API when available
   * (Chrome 86+, Edge 86+). Falls back to a hidden <input webkitdirectory> for
   * Firefox and older browsers.
   *
   * Why showDirectoryPicker instead of webkitdirectory?
   * With webkitdirectory the OS file picker opens in "file" mode — clicking a
   * folder navigates INTO it instead of selecting it, which is confusing.
   * showDirectoryPicker() opens a true "Select Folder" OS dialog where a single
   * click on the folder + the "Select Folder" button is all the user needs.
   */
  const handleFolderButtonClick = useCallback(async (e: React.MouseEvent) => {
    e.stopPropagation();
    if (isProcessing || isConverting) return;

    // Modern path: File System Access API
    if ('showDirectoryPicker' in window) {
      try {
        // @ts-expect-error showDirectoryPicker is not yet in lib.dom.d.ts
        const dirHandle: FileSystemDirectoryHandle = await window.showDirectoryPicker({ mode: 'read' });

        const files: File[] = [];
        // @ts-expect-error FileSystemDirectoryHandle.values() not in all TS targets
        for await (const entry of dirHandle.values()) {
          if (entry.kind === 'file') {
            const file: File = await entry.getFile();
            if (!file.name.startsWith('.') && ACCEPTED_MIME_TYPES.has(file.type)) {
              files.push(file);
            }
          }
        }

        if (files.length > 0) {
          processFiles(files);
        } else {
          alert('No supported files found in the selected folder (PDF, DOCX, JPG, PNG)');
        }
      } catch (err) {
        // AbortError = user pressed Cancel — silently ignore
        if ((err as DOMException).name !== 'AbortError') {
          console.error('[DropZone] showDirectoryPicker error:', err);
        }
      }
      return;
    }

    // Fallback: webkitdirectory hidden input
    folderInputRef.current?.click();
  }, [isProcessing, isConverting, processFiles]);

  const getProgressText = () => {
    if (isConverting) {
      return `Converting files... ${conversionProgress.current} of ${conversionProgress.total}`;
    }
    if (isProcessing && processingProgress) {
      return `Analyzing certificates... ${processingProgress.current} of ${processingProgress.total}`;
    }
    return 'Processing...';
  };

  const getProgressPercentage = () => {
    if (isConverting && conversionProgress.total > 0) {
      return (conversionProgress.current / conversionProgress.total) * 100;
    }
    if (isProcessing && processingProgress && processingProgress.total > 0) {
      return (processingProgress.current / processingProgress.total) * 100;
    }
    return 0;
  };

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={cn(
        "relative border-2 border-dashed rounded-xl transition-all duration-300 cursor-pointer",
        "flex flex-col items-center justify-center gap-3 sm:gap-4",
        "min-h-[140px] sm:min-h-[180px] py-6 sm:py-10 px-4",
        isDragging
          ? "border-yellow-500 bg-yellow-50/50"
          : "border-gray-200 hover:border-gray-300 bg-gray-50/50 hover:bg-gray-100/50",
        (isProcessing || isConverting) && "pointer-events-none"
      )}
    >
      {/* Standard file picker — covers the entire dropzone for click-to-browse */}
      <input
        type="file"
        accept=".pdf,.jpg,.jpeg,.png,.docx"
        multiple
        onChange={handleFileSelect}
        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
        disabled={isProcessing || isConverting}
      />

      {/* Fallback folder picker for browsers without showDirectoryPicker */}
      {/* @ts-expect-error webkitdirectory / directory not in React InputHTMLAttributes */}
      <input
        ref={folderInputRef}
        type="file"
        style={{ display: 'none' }}
        multiple
        webkitdirectory=""
        directory=""
        onChange={handleFolderInputChange}
        disabled={isProcessing || isConverting}
      />

      {(isProcessing || isConverting) ? (
        <div className="flex flex-col items-center gap-3 sm:gap-4 px-4 sm:px-8 w-full max-w-md">
          <div className="relative">
            <div className="w-10 h-10 sm:w-14 sm:h-14 rounded-full bg-yellow-100 flex items-center justify-center">
              <Loader2 className="w-5 h-5 sm:w-7 sm:h-7 text-yellow-600 animate-spin" />
            </div>
          </div>
          <div className="text-center">
            <p className="text-sm sm:text-base font-medium text-gray-900 mb-1">
              {getProgressText()}
            </p>
            <p className="text-xs sm:text-sm text-gray-500">
              Please wait while we process your documents
            </p>
          </div>
          {/* Progress Bar */}
          <div className="w-full h-1.5 sm:h-2 bg-gray-200 rounded-full overflow-hidden">
            <div
              className="h-full bg-yellow-500 rounded-full transition-all duration-300 ease-out"
              style={{ width: `${getProgressPercentage()}%` }}
            />
          </div>
        </div>
      ) : (
        <>
          <div className={cn(
            "w-10 h-10 sm:w-14 sm:h-14 rounded-xl flex items-center justify-center transition-colors",
            isDragging ? "bg-yellow-100" : "bg-gray-100"
          )}>
            {isDragging ? (
              <FileUp className="w-5 h-5 sm:w-7 sm:h-7 text-yellow-600" />
            ) : (
              <Upload className="w-5 h-5 sm:w-7 sm:h-7 text-gray-400" />
            )}
          </div>
          <div className="text-center">
            <p className="text-sm sm:text-base font-medium text-gray-900">
              {isDragging ? 'Drop your files here' : 'Drop certificates here'}
            </p>
            <div className="flex items-center justify-center gap-2 mt-1">
              <p className="text-xs sm:text-sm text-gray-500">
                or <span className="text-yellow-600 font-medium">tap to browse</span>
              </p>
              <span className="text-gray-300 select-none">·</span>
              {/*
                relative z-10 places this button above the invisible overlay input.
                e.stopPropagation() prevents the click from bubbling to that overlay.
              */}
              <button
                type="button"
                onClick={handleFolderButtonClick}
                className="relative z-10 flex items-center gap-1 text-xs sm:text-sm text-yellow-600 font-medium hover:text-yellow-700 transition-colors"
              >
                <FolderOpen className="w-3.5 h-3.5" />
                Select Folder
              </button>
            </div>
          </div>
          <div className="flex items-center gap-2 text-[10px] sm:text-xs text-gray-400">
            <span className="px-1.5 sm:px-2 py-0.5 bg-gray-100 rounded">PDF</span>
            <span className="px-1.5 sm:px-2 py-0.5 bg-gray-100 rounded">DOCX</span>
            <span className="px-1.5 sm:px-2 py-0.5 bg-gray-100 rounded">JPG</span>
            <span className="px-1.5 sm:px-2 py-0.5 bg-gray-100 rounded">PNG</span>
          </div>
        </>
      )}
    </div>
  );
};

export default DropZone;
