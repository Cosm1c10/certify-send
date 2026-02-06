import { useCallback, useRef } from 'react';
import { Database, Loader2, CheckCircle, XCircle, Upload } from 'lucide-react';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';

// =============================================================================
// MASTER FILE SYNC COMPONENT
// Dropzone for uploading the client's Master Excel file to build DYNAMIC_SUPPLIER_MAP
// =============================================================================

interface MasterFileSyncProps {
  isLoaded: boolean;
  isLoading: boolean;
  error: string | null;
  fileName: string | null;
  totalSuppliers: number;
  onFileSelect: (file: File) => void;
  onClear: () => void;
}

const MasterFileSync = ({
  isLoaded,
  isLoading,
  error,
  fileName,
  totalSuppliers,
  onFileSelect,
  onClear,
}: MasterFileSyncProps) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onFileSelect(file);
    }
    e.target.value = '';
  }, [onFileSelect]);

  const handleClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  // Render different states
  if (isLoading) {
    return (
      <div className="flex items-center gap-3 p-4 bg-blue-50 border border-blue-200 rounded-lg">
        <Loader2 className="w-5 h-5 text-blue-600 animate-spin flex-shrink-0" />
        <div>
          <p className="text-sm font-medium text-blue-900">Loading Master File...</p>
          <p className="text-xs text-blue-700">Building supplier map</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-between gap-3 p-4 bg-red-50 border border-red-200 rounded-lg">
        <div className="flex items-center gap-3">
          <XCircle className="w-5 h-5 text-red-600 flex-shrink-0" />
          <div>
            <p className="text-sm font-medium text-red-900">Failed to load Master File</p>
            <p className="text-xs text-red-700">{error}</p>
          </div>
        </div>
        <Button
          variant="outline"
          size="sm"
          onClick={handleClick}
          className="text-red-700 border-red-300 hover:bg-red-100"
        >
          Retry
        </Button>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          className="hidden"
        />
      </div>
    );
  }

  if (isLoaded && fileName) {
    return (
      <div className="flex items-center justify-between gap-3 p-4 bg-green-50 border border-green-200 rounded-lg">
        <div className="flex items-center gap-3">
          <CheckCircle className="w-5 h-5 text-green-600 flex-shrink-0" />
          <div>
            <p className="text-sm font-medium text-green-900">Master File Synced</p>
            <p className="text-xs text-green-700">
              {fileName} - {totalSuppliers} suppliers loaded
            </p>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <Button
            variant="outline"
            size="sm"
            onClick={handleClick}
            className="text-green-700 border-green-300 hover:bg-green-100"
          >
            Replace
          </Button>
          <Button
            variant="outline"
            size="sm"
            onClick={onClear}
            className="text-gray-600 border-gray-300 hover:bg-gray-100"
          >
            Clear
          </Button>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileChange}
          className="hidden"
        />
      </div>
    );
  }

  // Default: Not loaded
  return (
    <div
      onClick={handleClick}
      className={cn(
        "flex items-center gap-3 p-4 bg-gray-50 border border-dashed border-gray-300 rounded-lg",
        "cursor-pointer hover:bg-gray-100 hover:border-gray-400 transition-colors"
      )}
    >
      <div className="w-10 h-10 rounded-lg bg-gray-200 flex items-center justify-center flex-shrink-0">
        <Database className="w-5 h-5 text-gray-500" />
      </div>
      <div className="flex-1">
        <p className="text-sm font-medium text-gray-900">Sync with Master File</p>
        <p className="text-xs text-gray-500">
          Upload your Master Excel to auto-correct supplier names
        </p>
      </div>
      <Upload className="w-4 h-4 text-gray-400" />
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileChange}
        className="hidden"
      />
    </div>
  );
};

export default MasterFileSync;
