import { Download, ShieldCheck, RotateCcw, AlertCircle, FileSearch, Zap, Globe } from 'lucide-react';
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
    <div className="min-h-screen bg-gray-50">
      {/* Catering Disposables Header */}
      <header className="bg-slate-700">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-3 sm:py-4">
          <div className="flex items-center justify-between gap-3">
            {/* CD Logo in white container */}
            <div className="bg-white rounded-lg px-2 sm:px-3 py-1.5 sm:py-2 flex-shrink-0">
              <img
                src="/cd-logo.gif"
                alt="Catering Disposables"
                className="h-8 sm:h-12 w-auto"
              />
            </div>

            {/* Badge */}
            <div className="cd-badge text-xs sm:text-sm px-3 sm:px-4 py-1.5 sm:py-2">
              <span className="hidden sm:inline">AI Compliance Intelligence Engine</span>
              <span className="sm:hidden">AI Compliance Engine</span>
            </div>
          </div>
        </div>
      </header>

      {/* Hero Section */}
      <section className="bg-white border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-8 sm:py-12">
          <div className="max-w-3xl mx-auto text-center">
            <div className="flex items-center justify-center gap-2 mb-3 sm:mb-4">
              <ShieldCheck className="w-4 h-4 sm:w-5 sm:h-5 text-yellow-500" />
              <span className="text-xs sm:text-sm font-medium text-yellow-600 uppercase tracking-wide">
                Certificate Analysis
              </span>
            </div>
            <h2 className="text-2xl sm:text-3xl font-bold text-gray-900 mb-3 sm:mb-4 px-2">
              Automated Supplier{' '}
              <span className="relative inline-block">
                <span className="relative z-10">Certificate Extraction</span>
                <span className="absolute left-0 bottom-0.5 sm:bottom-1 w-full h-2 sm:h-3 bg-yellow-400/30 -z-0"></span>
              </span>
            </h2>
            <p className="text-gray-600 text-base sm:text-lg leading-relaxed mb-5 sm:mb-6 px-2">
              Upload supplier compliance certificates and let our AI-powered engine automatically
              extract, validate, and organize critical certification data. Streamline your
              compliance workflow with instant document processing.
            </p>

            {/* Feature Pills */}
            <div className="flex flex-col sm:flex-row flex-wrap justify-center gap-2 sm:gap-3 px-4 sm:px-0">
              <div className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-gray-100 rounded-full text-xs sm:text-sm text-gray-700">
                <Zap className="w-3.5 h-3.5 sm:w-4 sm:h-4 text-yellow-500" />
                Instant Processing
              </div>
              <div className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-gray-100 rounded-full text-xs sm:text-sm text-gray-700">
                <FileSearch className="w-3.5 h-3.5 sm:w-4 sm:h-4 text-yellow-500" />
                AI-Powered Extraction
              </div>
              <div className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-gray-100 rounded-full text-xs sm:text-sm text-gray-700">
                <Globe className="w-3.5 h-3.5 sm:w-4 sm:h-4 text-yellow-500" />
                Browser-Based Tool
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 py-6 sm:py-10">
        {/* Upload Card */}
        <div className="enterprise-card-elevated p-4 sm:p-8 mb-6 sm:mb-8">
          <div className="mb-4 sm:mb-6">
            <h3 className="text-base sm:text-lg font-semibold text-gray-900 mb-1">
              Upload Certificates
            </h3>
            <p className="text-xs sm:text-sm text-gray-500">
              Drag and drop PDF certificates or click to browse. Batch uploads supported.
            </p>
          </div>

          <DropZone
            onFilesProcess={analyzeCertificates}
            isProcessing={isProcessing}
            processingProgress={processingProgress}
          />
        </div>

        {/* Processing Errors */}
        {processingErrors.length > 0 && (
          <div className="mb-6 sm:mb-8 p-4 sm:p-5 bg-red-50 border border-red-200 rounded-xl">
            <div className="flex items-center gap-2 mb-2 sm:mb-3">
              <AlertCircle className="w-4 h-4 sm:w-5 sm:h-5 text-red-600 flex-shrink-0" />
              <h3 className="font-semibold text-red-800 text-sm sm:text-base">
                {processingErrors.length} file(s) failed to process
              </h3>
            </div>
            <ul className="text-xs sm:text-sm text-red-700 list-disc list-inside space-y-1">
              {processingErrors.map((error, index) => (
                <li key={index}>{error}</li>
              ))}
            </ul>
          </div>
        )}

        {/* Results Card */}
        <div className="enterprise-card-elevated p-4 sm:p-8">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-4 sm:mb-6">
            <div>
              <h3 className="text-base sm:text-lg font-semibold text-gray-900 mb-1">
                Extraction Results
              </h3>
              <p className="text-xs sm:text-sm text-gray-500">
                {certificates.length === 0
                  ? 'Processed certificates will appear here'
                  : `${certificates.length} certificate${certificates.length !== 1 ? 's' : ''} processed`
                }
              </p>
            </div>

            {/* Action Buttons */}
            <div className="flex items-center gap-2 sm:gap-3">
              {certificates.length > 0 && (
                <Button
                  onClick={clearCertificates}
                  variant="outline"
                  size="sm"
                  className="gap-1.5 sm:gap-2 text-gray-600 border-gray-300 hover:bg-gray-50 text-xs sm:text-sm"
                  disabled={isProcessing}
                >
                  <RotateCcw className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
                  <span className="hidden xs:inline">Clear All</span>
                  <span className="xs:hidden">Clear</span>
                </Button>
              )}
              <Button
                onClick={handleExport}
                disabled={certificates.length === 0 || isProcessing}
                size="sm"
                className="gap-1.5 sm:gap-2 bg-yellow-500 hover:bg-yellow-600 text-gray-900 font-semibold shadow-sm text-xs sm:text-sm"
              >
                <Download className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
                <span className="hidden xs:inline">Export to Excel</span>
                <span className="xs:hidden">Export</span>
              </Button>
            </div>
          </div>

          {/* Horizontal scroll wrapper for table on mobile */}
          <div className="-mx-4 sm:mx-0">
            <div className="overflow-x-auto px-4 sm:px-0">
              <ReviewTable certificates={certificates} />
            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="mt-6 sm:mt-10 bg-slate-700">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-6 sm:py-8">
          <div className="flex flex-col items-center gap-4 sm:gap-6 md:flex-row md:justify-between">
            {/* Logo in white container */}
            <div className="bg-white rounded-lg px-2 sm:px-3 py-1.5 sm:py-2">
              <img
                src="/cd-logo.gif"
                alt="Catering Disposables"
                className="h-8 sm:h-10 w-auto"
              />
            </div>

            {/* Center info + Agency Branding */}
            <div className="text-center">
              <p className="text-white text-xs sm:text-sm font-medium">AI Compliance Intelligence Engine</p>
              <p className="text-slate-300 text-[10px] sm:text-xs mt-1">Automated Certificate Extraction & Validation</p>
              <p className="text-slate-400 text-[10px] sm:text-xs mt-2 sm:mt-3">© 2026 AgenticFloww. All rights reserved.</p>
            </div>

            {/* Version */}
            <p className="text-slate-400 text-[10px] sm:text-xs">v1.1</p>
          </div>
        </div>
      </footer>
    </div>
  );
};

export default Index;
