import { useCallback, useState } from 'react';
import { Upload, Loader2 } from 'lucide-react';
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

const DropZone = ({ onFilesProcess, isProcessing, processingProgress }: DropZoneProps) => {
  const [isDragging, setIsDragging] = useState(false);
  const [isConverting, setIsConverting] = useState(false);
  const [conversionProgress, setConversionProgress] = useState({ current: 0, total: 0 });

  const convertPdfToBase64 = useCallback(async (file: File): Promise<string> => {
    const pdfjsLib = await import('pdfjs-dist');
    pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
      'pdfjs-dist/build/pdf.worker.min.mjs',
      import.meta.url
    ).toString();

    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
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
    }).promise;

    return canvas.toDataURL('image/png');
  }, []);

  const processFiles = useCallback(async (files: File[]) => {
    const pdfFiles = files.filter(file => file.type === 'application/pdf');

    if (pdfFiles.length === 0) {
      alert('Please upload PDF files');
      return;
    }

    if (pdfFiles.length !== files.length) {
      alert(`${files.length - pdfFiles.length} non-PDF files were skipped`);
    }

    setIsConverting(true);
    setConversionProgress({ current: 0, total: pdfFiles.length });
    console.log(`Starting PDF conversion for ${pdfFiles.length} files`);

    const convertedFiles: FileWithBase64[] = [];

    for (let i = 0; i < pdfFiles.length; i++) {
      const file = pdfFiles[i];
      setConversionProgress({ current: i + 1, total: pdfFiles.length });

      try {
        console.log(`Converting ${i + 1}/${pdfFiles.length}: ${file.name}`);
        const base64Image = await convertPdfToBase64(file);
        convertedFiles.push({ file, base64Image });
      } catch (error) {
        console.error(`Error converting PDF ${file.name}:`, error);
      }
    }

    setIsConverting(false);
    setConversionProgress({ current: 0, total: 0 });

    if (convertedFiles.length > 0) {
      onFilesProcess(convertedFiles);
    } else {
      alert('Failed to convert any PDF files');
    }
  }, [convertPdfToBase64, onFilesProcess]);

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
    // Reset input so the same files can be selected again
    e.target.value = '';
  }, [processFiles]);

  const getProgressText = () => {
    if (isConverting) {
      return `Converting PDFs... ${conversionProgress.current} of ${conversionProgress.total}`;
    }
    if (isProcessing && processingProgress) {
      return `Analyzing certificates... ${processingProgress.current} of ${processingProgress.total}`;
    }
    return 'Processing...';
  };

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={cn(
        "relative border-2 border-dashed rounded-xl p-12 transition-all duration-300 cursor-pointer",
        "flex flex-col items-center justify-center gap-4",
        "min-h-[280px]",
        isDragging 
          ? "border-primary bg-primary/5 scale-[1.02]" 
          : "border-border hover:border-primary/50 hover:bg-muted/50",
        (isProcessing || isConverting) && "pointer-events-none opacity-70"
      )}
    >
      <input
        type="file"
        accept=".pdf"
        multiple
        onChange={handleFileSelect}
        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
        disabled={isProcessing || isConverting}
      />

      {(isProcessing || isConverting) ? (
        <>
          <Loader2 className="w-16 h-16 text-primary animate-spin" />
          <p className="text-lg font-medium text-foreground">
            {getProgressText()}
          </p>
        </>
      ) : (
        <>
          <Upload className="w-14 h-14 text-muted-foreground mb-2" />
          <div className="text-center">
            <p className="text-lg font-medium text-foreground">
              Drop your PDF certificates here
            </p>
            <p className="text-sm text-muted-foreground mt-1">
              or click to browse files (multiple supported)
            </p>
          </div>
        </>
      )}
    </div>
  );
};

export default DropZone;
