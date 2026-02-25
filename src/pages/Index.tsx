import { useState } from 'react';
import { Download, ShieldCheck, RotateCcw, AlertCircle, FileSearch, Zap, Globe, FileSpreadsheet, ChevronsUpDown, Check, X, UserCheck, Plus } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Popover, PopoverContent, PopoverTrigger } from '@/components/ui/popover';
import { Command, CommandEmpty, CommandGroup, CommandInput, CommandItem, CommandList, CommandSeparator } from '@/components/ui/command';
import { cn } from '@/lib/utils';
import DropZone from '@/components/DropZone';
import ReviewTable from '@/components/ReviewTable';
import MasterFileSync from '@/components/MasterFileSync';
import NewSuppliersDialog from '@/components/NewSuppliersDialog';
import { useCertificates } from '@/hooks/useCertificates';
import { useMasterFile } from '@/hooks/useMasterFile';
import { exportToExcel, prepareExportData, NewSupplierInfo } from '@/utils/exportExcel';
import { exportFeederExcel } from '@/utils/exportFeederExcel';
import { appendToMasterExcel } from '@/utils/appendMasterExcel';

const Index = () => {
  // Master File must be initialized FIRST so we can pass supplierMap to useCertificates
  const {
    masterFile,
    isLoading: isMasterFileLoading,
    error: masterFileError,
    loadMasterFile,
    clearMasterFile,
  } = useMasterFile();

  const {
    certificates,
    isProcessing,
    processingProgress,
    processingErrors,
    analyzeCertificates,
    clearCertificates
  } = useCertificates(masterFile.supplierMap);

  const [feederStats, setFeederStats] = useState<{ matched: number; newSuppliers: number } | null>(null);

  // Supplier override state (broker/manufacturer workflow)
  const [selectedSupplierOverride, setSelectedSupplierOverride] = useState('');
  const [comboboxOpen, setComboboxOpen] = useState(false);
  // Tracks live search text inside the combobox — used to offer "Add new supplier" option
  const [supplierSearchValue, setSupplierSearchValue] = useState('');
  // "Add new supplier" inline-input mode (constant button below the dropdown)
  const [addingNewSupplier, setAddingNewSupplier] = useState(false);
  const [newSupplierAccount, setNewSupplierAccount] = useState('');   // short code, e.g. "SOWPAK"
  const [newSupplierName, setNewSupplierName]       = useState('');   // full name, e.g. "Sowinpak Ltd"
  // Confirmed account code for the currently selected supplier override
  const [selectedSupplierAccount, setSelectedSupplierAccount] = useState('');

  // New Suppliers Dialog state
  const [showNewSuppliersDialog, setShowNewSuppliersDialog] = useState(false);
  const [detectedNewSuppliers, setDetectedNewSuppliers] = useState<NewSupplierInfo[]>([]);
  const [pendingExportType, setPendingExportType] = useState<'excel' | 'feeder' | null>(null);

  // Proceed with export after user confirms
  const proceedWithExport = async () => {
    setShowNewSuppliersDialog(false);

    if (pendingExportType === 'excel') {
      try {
        if (masterFile.rawBuffer) {
          await appendToMasterExcel(masterFile.rawBuffer, certificates, masterFile.supplierMap);
        } else {
          await exportToExcel(certificates, masterFile.supplierMap);
        }
      } catch (error) {
        console.error('Export failed:', error);
        const msg = error instanceof Error ? error.message : String(error);
        alert(`Failed to export Excel file:\n\n${msg}`);
      }
    } else if (pendingExportType === 'feeder') {
      try {
        const stats = await exportFeederExcel(certificates, masterFile.supplierMap);
        if (masterFile.isLoaded) {
          alert(`Feeder file exported!\n\n${stats.matched} suppliers matched from Master File\n${stats.newSuppliers} NEW suppliers (flagged in purple)\n${stats.duplicatesRemoved} duplicates removed`);
        } else {
          alert(`Feeder file exported!\n\n${stats.total} certificates\n${stats.duplicatesRemoved} duplicates removed`);
        }
      } catch (error) {
        console.error('Feeder export failed:', error);
        alert('Failed to export feeder file');
      }
    }

    setPendingExportType(null);
    setDetectedNewSuppliers([]);
  };

  const handleExport = async () => {
    if (certificates.length === 0) {
      alert('No certificates to export');
      return;
    }

    // Check for new suppliers before exporting
    const { newSuppliers } = prepareExportData(certificates, masterFile.supplierMap);

    if (newSuppliers.length > 0 && masterFile.isLoaded) {
      // Show warning dialog
      setDetectedNewSuppliers(newSuppliers);
      setPendingExportType('excel');
      setShowNewSuppliersDialog(true);
      return;
    }

    // No new suppliers or no master file - proceed directly
    try {
      if (masterFile.rawBuffer) {
        await appendToMasterExcel(masterFile.rawBuffer, certificates, masterFile.supplierMap);
      } else {
        await exportToExcel(certificates, masterFile.supplierMap);
      }
    } catch (error) {
      console.error('Export failed:', error);
      const msg = error instanceof Error ? error.message : String(error);
      alert(`Failed to export Excel file:\n\n${msg}`);
    }
  };

  const handleExportFeeder = async () => {
    if (certificates.length === 0) {
      alert('No certificates to export');
      return;
    }

    // Check for new suppliers before exporting
    const { newSuppliers } = prepareExportData(certificates, masterFile.supplierMap);

    if (newSuppliers.length > 0 && masterFile.isLoaded) {
      // Show warning dialog
      setDetectedNewSuppliers(newSuppliers);
      setPendingExportType('feeder');
      setShowNewSuppliersDialog(true);
      return;
    }

    // No new suppliers or no master file - proceed directly
    try {
      const stats = await exportFeederExcel(certificates, masterFile.supplierMap);
      setFeederStats({ matched: stats.matched, newSuppliers: stats.newSuppliers });

      if (masterFile.isLoaded) {
        alert(`Feeder file exported!\n\n${stats.matched} suppliers matched from Master File\n${stats.newSuppliers} NEW suppliers (flagged in purple)\n${stats.duplicatesRemoved} duplicates removed`);
      } else {
        alert(`Feeder file exported!\n\n${stats.total} certificates\n${stats.duplicatesRemoved} duplicates removed\n\nTip: Sync with Master File to auto-correct supplier names!`);
      }
    } catch (error) {
      console.error('Feeder export failed:', error);
      alert('Failed to export feeder file');
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

          {/* Supplier Override Input — assign all uploaded PDFs to a specific supplier */}
          <div className="mb-4">
            <label className="block text-xs font-medium text-gray-600 mb-1.5">
              <UserCheck className="inline w-3.5 h-3.5 mr-1 text-gray-400" />
              Assign to Supplier Account
              <span className="text-gray-400 font-normal ml-1">(Optional — overrides AI extraction)</span>
            </label>

            {masterFile.isLoaded ? (
              /* Searchable combobox when Master File is loaded */
              <Popover open={comboboxOpen} onOpenChange={(open) => { setComboboxOpen(open); if (!open) setSupplierSearchValue(''); }}>
                <PopoverTrigger asChild>
                  <button
                    type="button"
                    className={cn(
                      "w-full flex items-center justify-between rounded-lg border px-3 py-2 text-sm",
                      "bg-white hover:bg-gray-50 transition-colors text-left",
                      selectedSupplierOverride
                        ? "border-yellow-400 bg-yellow-50/60 text-gray-900"
                        : "border-gray-200 text-gray-400"
                    )}
                  >
                    <span className={selectedSupplierOverride ? "text-gray-900 font-medium" : ""}>
                      {selectedSupplierOverride || "Search or add a supplier..."}
                    </span>
                    <div className="flex items-center gap-1 flex-shrink-0 ml-2">
                      {selectedSupplierOverride && (
                        <span
                          role="button"
                          tabIndex={0}
                          onClick={(e) => { e.stopPropagation(); setSelectedSupplierOverride(''); setSelectedSupplierAccount(''); }}
                          onKeyDown={(e) => { if (e.key === 'Enter') { e.stopPropagation(); setSelectedSupplierOverride(''); setSelectedSupplierAccount(''); } }}
                          className="p-0.5 rounded hover:bg-yellow-200 text-yellow-700"
                        >
                          <X className="w-3.5 h-3.5" />
                        </span>
                      )}
                      <ChevronsUpDown className="w-3.5 h-3.5 text-gray-400" />
                    </div>
                  </button>
                </PopoverTrigger>
                <PopoverContent className="w-[--radix-popover-trigger-width] p-0" align="start">
                  <Command>
                    <CommandInput
                      placeholder="Search or type a new supplier name..."
                      className="h-9"
                      onValueChange={setSupplierSearchValue}
                    />
                    <CommandList>
                      <CommandEmpty>
                        {supplierSearchValue.trim()
                          ? 'No match — use "Add new" below.'
                          : 'No suppliers found.'}
                      </CommandEmpty>
                      <CommandGroup>
                        {Object.values(masterFile.supplierMap)
                          .map(e => e.officialName)
                          .filter((name, i, arr) => arr.indexOf(name) === i) // dedupe
                          .sort()
                          .map(name => (
                            <CommandItem
                              key={name}
                              value={name}
                              onSelect={(val) => {
                                setSelectedSupplierOverride(val === selectedSupplierOverride ? '' : val);
                                setComboboxOpen(false);
                              }}
                            >
                              <Check className={cn("mr-2 h-4 w-4", selectedSupplierOverride === name ? "opacity-100" : "opacity-0")} />
                              {name}
                            </CommandItem>
                          ))
                        }
                      </CommandGroup>

                      {/* "Add new supplier" option — visible whenever the user has typed
                          text that isn't an exact match for an existing supplier */}
                      {supplierSearchValue.trim() &&
                        !Object.values(masterFile.supplierMap).some(
                          e => e.officialName.toLowerCase() === supplierSearchValue.toLowerCase().trim()
                        ) && (
                        <>
                          <CommandSeparator />
                          <CommandGroup>
                            <CommandItem
                              value={`add new supplier ${supplierSearchValue}`}
                              onSelect={() => {
                                setSelectedSupplierOverride(supplierSearchValue.trim());
                                setComboboxOpen(false);
                                setSupplierSearchValue('');
                              }}
                              className="text-green-700"
                            >
                              <Plus className="mr-2 h-4 w-4 text-green-600 flex-shrink-0" />
                              Add new: <span className="ml-1 font-semibold">{supplierSearchValue.trim()}</span>
                            </CommandItem>
                          </CommandGroup>
                        </>
                      )}
                    </CommandList>
                  </Command>
                </PopoverContent>
              </Popover>
            ) : (
              /* Plain text input when no Master File */
              <div className="relative">
                <Input
                  type="text"
                  placeholder="Type supplier name to override AI extraction..."
                  value={selectedSupplierOverride}
                  onChange={(e) => setSelectedSupplierOverride(e.target.value)}
                  className={cn(
                    "text-sm pr-8",
                    selectedSupplierOverride && "border-yellow-400 bg-yellow-50/60"
                  )}
                />
                {selectedSupplierOverride && (
                  <button
                    type="button"
                    onClick={() => setSelectedSupplierOverride('')}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>
            )}

            {/* ── Add New Supplier ─────────────────────────────────────────────
                Always-visible shortcut so the user doesn't have to type inside the
                combobox. Only shown when a Master File is loaded (without one the
                plain text input above already serves this purpose).             */}
            {masterFile.isLoaded && (
              addingNewSupplier ? (
                /* Two-field inline form: Account code + Full name */
                <div className="mt-2 space-y-1.5">
                  <div className="flex items-center gap-1.5">
                    <Input
                      autoFocus
                      type="text"
                      placeholder="Account code (e.g. SOWPAK)"
                      value={newSupplierAccount}
                      onChange={(e) => setNewSupplierAccount(e.target.value)}
                      onKeyDown={(e) => { if (e.key === 'Escape') { setNewSupplierAccount(''); setNewSupplierName(''); setAddingNewSupplier(false); } }}
                      className="h-8 text-sm flex-1 border-green-300 focus-visible:ring-green-400"
                    />
                  </div>
                  <div className="flex items-center gap-1.5">
                    <Input
                      type="text"
                      placeholder="Full company name (e.g. Sowinpak Ltd)"
                      value={newSupplierName}
                      onChange={(e) => setNewSupplierName(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && newSupplierAccount.trim() && newSupplierName.trim()) {
                          setSelectedSupplierOverride(newSupplierName.trim());
                          setSelectedSupplierAccount(newSupplierAccount.trim());
                          setNewSupplierAccount('');
                          setNewSupplierName('');
                          setAddingNewSupplier(false);
                        }
                        if (e.key === 'Escape') {
                          setNewSupplierAccount('');
                          setNewSupplierName('');
                          setAddingNewSupplier(false);
                        }
                      }}
                      className="h-8 text-sm flex-1 border-green-300 focus-visible:ring-green-400"
                    />
                    <button
                      type="button"
                      disabled={!newSupplierAccount.trim() || !newSupplierName.trim()}
                      onClick={() => {
                        if (newSupplierAccount.trim() && newSupplierName.trim()) {
                          setSelectedSupplierOverride(newSupplierName.trim());
                          setSelectedSupplierAccount(newSupplierAccount.trim());
                          setNewSupplierAccount('');
                          setNewSupplierName('');
                          setAddingNewSupplier(false);
                        }
                      }}
                      className="h-8 w-8 flex items-center justify-center rounded border border-green-300 text-green-700 hover:bg-green-50 disabled:opacity-40 disabled:cursor-not-allowed flex-shrink-0"
                    >
                      <Check className="w-3.5 h-3.5" />
                    </button>
                    <button
                      type="button"
                      onClick={() => { setNewSupplierAccount(''); setNewSupplierName(''); setAddingNewSupplier(false); }}
                      className="h-8 w-8 flex items-center justify-center rounded border border-gray-200 text-gray-400 hover:bg-gray-50 flex-shrink-0"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>
              ) : (
                /* Constant "Add new supplier" button */
                <button
                  type="button"
                  onClick={() => setAddingNewSupplier(true)}
                  className="mt-2 flex items-center gap-1 text-xs font-medium text-green-700 hover:text-green-800 hover:underline"
                >
                  <Plus className="w-3 h-3" />
                  Add new supplier
                </button>
              )
            )}

            {selectedSupplierOverride && (
              <p className="text-xs text-yellow-700 mt-1 flex items-center gap-1 flex-wrap">
                <Check className="w-3 h-3 flex-shrink-0" />
                All uploaded certificates will be assigned to{' '}
                <strong>{selectedSupplierOverride}</strong>
                {selectedSupplierAccount && (
                  <span className="text-gray-500 font-normal">
                    (account: <strong className="text-gray-700">{selectedSupplierAccount}</strong>)
                  </span>
                )}
              </p>
            )}
          </div>

          <DropZone
            onFilesProcess={(files) => analyzeCertificates(
              files,
              selectedSupplierOverride || undefined,
              selectedSupplierAccount  || undefined
            )}
            isProcessing={isProcessing}
            processingProgress={processingProgress}
          />

          {/* Master File Sync Section */}
          <div className="mt-4 pt-4 border-t border-gray-200">
            <MasterFileSync
              isLoaded={masterFile.isLoaded}
              isLoading={isMasterFileLoading}
              error={masterFileError}
              fileName={masterFile.fileName}
              totalSuppliers={masterFile.totalSuppliers}
              onFileSelect={loadMasterFile}
              onClear={clearMasterFile}
            />
          </div>
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
                onClick={handleExportFeeder}
                disabled={certificates.length === 0 || isProcessing}
                size="sm"
                variant="outline"
                className="gap-1.5 sm:gap-2 border-purple-300 text-purple-700 hover:bg-purple-50 font-semibold text-xs sm:text-sm"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
                <span className="hidden xs:inline">Export Feeder</span>
                <span className="xs:hidden">Feeder</span>
              </Button>
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
            <p className="text-slate-400 text-[10px] sm:text-xs">v1.2</p>
          </div>
        </div>
      </footer>

      {/* New Suppliers Warning Dialog */}
      <NewSuppliersDialog
        isOpen={showNewSuppliersDialog}
        newSuppliers={detectedNewSuppliers}
        onClose={() => {
          setShowNewSuppliersDialog(false);
          setPendingExportType(null);
          setDetectedNewSuppliers([]);
        }}
        onProceed={proceedWithExport}
        hasMasterFile={masterFile.isLoaded}
      />
    </div>
  );
};

export default Index;
