import { useCallback, useState } from 'react';
import { Upload, FileText, Loader2 } from 'lucide-react';
import { cn } from '@/lib/utils';

interface DropZoneProps {
  onFileProcess: (file: File, base64Image: string) => void;
  isProcessing: boolean;
}

const DropZone = ({ onFileProcess, isProcessing }: DropZoneProps) => {
  const [isDragging, setIsDragging] = useState(false);

  const processFile = useCallback(async (file: File) => {
    if (file.type !== 'application/pdf') {
      alert('Please upload a PDF file');
      return;
    }

    const pdfjsLib = await import('pdfjs-dist');
    pdfjsLib.GlobalWorkerOptions.workerSrc = `//cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.js`;

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
      canvas: canvas,
    }).promise;

    const base64Image = canvas.toDataURL('image/png');
    onFileProcess(file, base64Image);
  }, [onFileProcess]);

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
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      processFile(files[0]);
    }
  }, [processFile]);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      processFile(files[0]);
    }
  }, [processFile]);

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
        isProcessing && "pointer-events-none opacity-70"
      )}
    >
      <input
        type="file"
        accept=".pdf"
        onChange={handleFileSelect}
        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
        disabled={isProcessing}
      />
      
      {isProcessing ? (
        <>
          <Loader2 className="w-16 h-16 text-primary animate-spin" />
          <p className="text-lg font-medium text-foreground">Processing certificate...</p>
        </>
      ) : (
        <>
          <div className="flex items-center justify-center gap-3 mb-2">
            <Upload className="w-14 h-14 text-muted-foreground" />
            <FileText className="w-10 h-10 text-primary" />
          </div>
          <div className="text-center">
            <p className="text-lg font-medium text-foreground">
              Drop your PDF certificate here
            </p>
            <p className="text-sm text-muted-foreground mt-1">
              or click to browse files
            </p>
          </div>
        </>
      )}
    </div>
  );
};

export default DropZone;
