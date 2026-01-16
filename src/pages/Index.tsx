import { Download, ShieldCheck } from 'lucide-react';
import { Button } from '@/components/ui/button';
import DropZone from '@/components/DropZone';
import ReviewTable from '@/components/ReviewTable';
import { useCertificates } from '@/hooks/useCertificates';
import { exportToExcel } from '@/utils/exportExcel';

const Index = () => {
  const { certificates, isProcessing, analyzeCertificate } = useCertificates();

  const handleExport = () => {
    if (certificates.length === 0) {
      alert('No certificates to export');
      return;
    }
    exportToExcel(certificates);
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
            onFileProcess={analyzeCertificate} 
            isProcessing={isProcessing} 
          />
        </div>

        {/* Review Table */}
        <div className="mb-8">
          <h2 className="text-xl font-semibold text-foreground mb-4">
            Review Table
          </h2>
          <ReviewTable certificates={certificates} />
        </div>

        {/* Footer */}
        <div className="flex justify-center pt-8 border-t border-border">
          <Button
            onClick={handleExport}
            disabled={certificates.length === 0}
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
