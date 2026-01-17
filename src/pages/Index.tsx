import { Download, ShieldCheck, RotateCcw, AlertCircle } from 'lucide-react';
import { Button } from '@/components/ui/button';
import DropZone from '@/components/DropZone';
import ReviewTable from '@/components/ReviewTable';
import { useCertificates } from '@/hooks/useCertificates';
import { exportToExcel } from '@/utils/exportExcel';

const Index = () => {
  const {
    certificates,
    isProcessing,
    processingProgress,
    processingErrors,
    analyzeCertificates,
    clearCertificates
  } = useCertificates();

  const handleExport = async () => {
    if (certificates.length === 0) {
      alert('No certificates to export');
      return;
    }
    try {
      await exportToExcel(certificates);
    } catch (error) {
      console.error('Export failed:', error);
      alert('Failed to export Excel file');
    }
  };

  return (
    <div className="min-h-screen bg-background">
      <div className="max-w-6xl mx-auto px-6 py-12">
        {/* Header */}
        <div className="text-center mb-12">
          <div className="flex items-center justify-center gap-3 mb-4">
            <ShieldCheck className="w-10 h-10 text-primary" />
            <h1 className="text-3xl font-bold text-foreground">
              Certificate Analyzer
            </h1>
          </div>
          <p className="text-muted-foreground max-w-xl mx-auto">
            Upload supplier certificates to automatically extract and review key information.
            Supports PDF documents with instant data extraction.
          </p>
        </div>

        {/* Drop Zone */}
        <div className="mb-12">
          <DropZone
            onFilesProcess={analyzeCertificates}
            isProcessing={isProcessing}
            processingProgress={processingProgress}
          />
        </div>

        {/* Processing Errors */}
        {processingErrors.length > 0 && (
          <div className="mb-8 p-4 bg-destructive/10 border border-destructive/20 rounded-xl">
            <div className="flex items-center gap-2 mb-2">
              <AlertCircle className="w-5 h-5 text-destructive" />
              <h3 className="font-semibold text-destructive">
                {processingErrors.length} file(s) failed to process
              </h3>
            </div>
            <ul className="text-sm text-destructive/80 list-disc list-inside">
              {processingErrors.map((error, index) => (
                <li key={index}>{error}</li>
              ))}
            </ul>
          </div>
        )}

        {/* Review Table */}
        <div className="mb-8">
          <h2 className="text-xl font-semibold text-foreground mb-4">
            Review Table
          </h2>
          <ReviewTable certificates={certificates} />
        </div>

        {/* Footer */}
        <div className="flex justify-center gap-4 pt-8 border-t border-border">
          {certificates.length > 0 && (
            <Button
              onClick={clearCertificates}
              variant="outline"
              size="lg"
              className="gap-2"
              disabled={isProcessing}
            >
              <RotateCcw className="w-5 h-5" />
              Start Again
            </Button>
          )}
          <Button
            onClick={handleExport}
            disabled={certificates.length === 0 || isProcessing}
            size="lg"
            className="gap-2"
          >
            <Download className="w-5 h-5" />
            Download Excel
          </Button>
        </div>
      </div>
    </div>
  );
};

export default Index;
