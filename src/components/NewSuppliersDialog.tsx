import { AlertTriangle, X, Download, FileWarning } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { NewSupplierInfo } from '@/utils/exportExcel';

// =============================================================================
// NEW SUPPLIERS DIALOG
// Warning modal shown when unknown suppliers are detected during export
// =============================================================================

interface NewSuppliersDialogProps {
  isOpen: boolean;
  newSuppliers: NewSupplierInfo[];
  onClose: () => void;
  onProceed: () => void;
  hasMasterFile: boolean;
}

const NewSuppliersDialog = ({
  isOpen,
  newSuppliers,
  onClose,
  onProceed,
  hasMasterFile,
}: NewSuppliersDialogProps) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center">
      {/* Backdrop */}
      <div
        className="absolute inset-0 bg-black/50 backdrop-blur-sm"
        onClick={onClose}
      />

      {/* Dialog */}
      <div className="relative bg-white rounded-xl shadow-2xl max-w-lg w-full mx-4 max-h-[80vh] flex flex-col">
        {/* Header */}
        <div className="flex items-center justify-between p-4 border-b border-gray-200">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center">
              <AlertTriangle className="w-5 h-5 text-amber-600" />
            </div>
            <div>
              <h2 className="text-lg font-semibold text-gray-900">
                New Suppliers Detected
              </h2>
              <p className="text-sm text-gray-500">
                {newSuppliers.length} supplier{newSuppliers.length !== 1 ? 's' : ''} not found in Master File
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
          >
            <X className="w-5 h-5 text-gray-500" />
          </button>
        </div>

        {/* Content */}
        <div className="flex-1 overflow-y-auto p-4">
          {!hasMasterFile && (
            <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
              <p className="text-sm text-blue-800">
                <strong>Tip:</strong> Upload your Master File to auto-match supplier names and reduce false positives.
              </p>
            </div>
          )}

          <p className="text-sm text-gray-600 mb-3">
            The following suppliers were extracted from certificates but don't match any existing supplier in your Master File.
            This could indicate:
          </p>
          <ul className="text-sm text-gray-600 mb-4 list-disc list-inside space-y-1">
            <li>A genuinely <strong>new supplier</strong> to add to your records</li>
            <li>A <strong>spelling variation</strong> or typo in the certificate</li>
            <li>A <strong>subsidiary or alias</strong> of an existing supplier</li>
          </ul>

          {/* Supplier List */}
          <div className="space-y-2">
            {newSuppliers.map((supplier, index) => (
              <div
                key={`${supplier.supplierName}-${index}`}
                className="p-3 bg-amber-50 border border-amber-200 rounded-lg"
              >
                <div className="flex items-start gap-3">
                  <FileWarning className="w-4 h-4 text-amber-600 mt-0.5 flex-shrink-0" />
                  <div className="flex-1 min-w-0">
                    <p className="font-medium text-gray-900 truncate">
                      {supplier.supplierName}
                    </p>
                    <p className="text-xs text-gray-500 mt-0.5">
                      From: {supplier.fileName} | Cert: {supplier.certification}
                    </p>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Footer */}
        <div className="p-4 border-t border-gray-200 bg-gray-50 rounded-b-xl">
          <div className="flex items-center justify-end gap-3">
            <Button
              variant="outline"
              onClick={onClose}
              className="text-gray-600"
            >
              Cancel
            </Button>
            <Button
              onClick={onProceed}
              className="bg-amber-500 hover:bg-amber-600 text-white gap-2"
            >
              <Download className="w-4 h-4" />
              Export Anyway
            </Button>
          </div>
          <p className="text-xs text-gray-500 mt-2 text-center">
            New suppliers will be flagged in the exported file for your review.
          </p>
        </div>
      </div>
    </div>
  );
};

export default NewSuppliersDialog;
