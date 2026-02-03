import { useCallback, useState } from 'react';
import { Upload, Loader2, FileUp } from 'lucide-react';
import { cn } from '@/lib/utils';

export interface FileWithBase64 {
  file: File;
  base64Image: string;
}

interface DropZoneProps {
  onFilesProcess: (files: FileWithBase64[]) => void;
  isProcessing: boolean;
  processingProgress?: { current: number; total: number };
}

const SUPPORTED_IMAGE_TYPES = ['image/jpeg', 'image/jpg', 'image/png'];
const SUPPORTED_PDF_TYPE = 'application/pdf';

const DropZone = ({ onFilesProcess, isProcessing, processingProgress }: DropZoneProps) => {
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [conversionProgress, setConversionProgress] = useState({ current: 0, total: 0 });

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

    return canvas.toDataURL('image/png');
  }, []);

  const processFiles = useCallback(async (files: File[]) => {
    // Guard: Prevent processing if already busy
    if (isProcessing || isConverting) {
      console.warn('Already processing - ignoring new upload');
      return;
    }

    // Filter for supported file types (PDFs and images)
    const supportedFiles = files.filter(file =>
      file.type === SUPPORTED_PDF_TYPE || SUPPORTED_IMAGE_TYPES.includes(file.type)
    );

    if (supportedFiles.length === 0) {
      alert('Please upload PDF or image files (JPG, PNG)');
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
        if (file.type === SUPPORTED_PDF_TYPE) {
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
      <input
        type="file"
        accept=".pdf,.jpg,.jpeg,.png"
        multiple
        onChange={handleFileSelect}
        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
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
            <p className="text-xs sm:text-sm text-gray-500 mt-1">
              or <span className="text-yellow-600 font-medium">tap to browse</span> your files
            </p>
          </div>
          <div className="flex items-center gap-2 text-[10px] sm:text-xs text-gray-400">
            <span className="px-1.5 sm:px-2 py-0.5 bg-gray-100 rounded">PDF</span>
            <span className="px-1.5 sm:px-2 py-0.5 bg-gray-100 rounded">JPG</span>
            <span className="px-1.5 sm:px-2 py-0.5 bg-gray-100 rounded">PNG</span>
            <span>Unlimited files</span>
          </div>
        </>
      )}
    </div>
  );
};

export default DropZone;
