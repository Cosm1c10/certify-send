import { AlertTriangle, X } from 'lucide-react';
import { Button } from '@/components/ui/button';

interface QuotaModalProps {
  isOpen: boolean;
  onClose: () => void;
}

const QuotaModal = ({ isOpen, onClose }: QuotaModalProps) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center">
      {/* Backdrop */}
      <div
        className="absolute inset-0 bg-black/50 backdrop-blur-sm"
        onClick={onClose}
      />

      {/* Modal */}
      <div className="relative bg-white rounded-2xl shadow-2xl max-w-md w-full mx-4 overflow-hidden animate-in fade-in zoom-in-95 duration-200">
        {/* Header with warning color */}
        <div className="bg-amber-50 border-b border-amber-200 px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-amber-100 flex items-center justify-center">
                <AlertTriangle className="w-5 h-5 text-amber-600" />
              </div>
              <h3 className="text-lg font-semibold text-amber-800">
                Demo Quota Reached
              </h3>
            </div>
            <button
              onClick={onClose}
              className="text-amber-600 hover:text-amber-800 transition-colors p-1 rounded-lg hover:bg-amber-100"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Content */}
        <div className="px-6 py-5">
          <p className="text-gray-700 leading-relaxed">
            This staging environment is capped at <span className="font-semibold text-amber-700">5 files</span> for testing purposes.
          </p>
          <p className="text-gray-600 mt-3 text-sm leading-relaxed">
            To process your full backlog, please contact your administrator to deploy the Production Environment.
          </p>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 bg-gray-50 border-t border-gray-100">
          <Button
            onClick={onClose}
            className="w-full bg-amber-500 hover:bg-amber-600 text-white font-medium"
          >
            Understood
          </Button>
        </div>
      </div>
    </div>
  );
};

export default QuotaModal;
