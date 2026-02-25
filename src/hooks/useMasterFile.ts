import { useState, useCallback } from 'react';
import { parseMasterFile, DynamicSupplierMap, MasterFileData, matchSupplier } from '@/utils/masterFileParser';
import { CertificateData } from '@/types/certificate';

// =============================================================================
// MASTER FILE SYNC HOOK
// Manages the DYNAMIC_SUPPLIER_MAP for the Feeder System
// =============================================================================

interface MasterFileState {
  isLoaded: boolean;
  fileName: string | null;
  totalSuppliers: number;
  supplierMap: DynamicSupplierMap;
  rawBuffer: ArrayBuffer | null;  // Original file binary — used for "Append to Master" export
}

export const useMasterFile = () => {
  const [masterFile, setMasterFile] = useState<MasterFileState>({
    isLoaded: false,
    fileName: null,
    totalSuppliers: 0,
    supplierMap: {},
    rawBuffer: null,
  });
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  /**
   * Load and parse the Master Excel file
   */
  const loadMasterFile = useCallback(async (file: File) => {
    setIsLoading(true);
    setError(null);

    try {
      // Validate file type
      const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel', // .xls
      ];

      if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
        throw new Error('Please upload an Excel file (.xlsx or .xls)');
      }

      console.log(`Loading Master File: ${file.name}`);
      // Read buffer FIRST — needed for "Append to Master" export later
      const rawBuffer = await file.arrayBuffer();
      const data: MasterFileData = await parseMasterFile(file);  // parseMasterFile reads buffer internally too — both calls are fine

      setMasterFile({
        isLoaded: true,
        fileName: data.fileName,
        totalSuppliers: data.totalSuppliers,
        supplierMap: data.supplierMap,
        rawBuffer,
      });

      console.log(`Master File loaded: ${data.totalSuppliers} suppliers mapped`);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Failed to parse Master File';
      console.error('Master File error:', message);
      setError(message);
      setMasterFile({
        isLoaded: false,
        fileName: null,
        totalSuppliers: 0,
        supplierMap: {},
        rawBuffer: null,
      });
    } finally {
      setIsLoading(false);
    }
  }, []);

  /**
   * Clear the loaded Master File
   */
  const clearMasterFile = useCallback(() => {
    setMasterFile({
      isLoaded: false,
      fileName: null,
      totalSuppliers: 0,
      supplierMap: {},
      rawBuffer: null,
    });
    setError(null);
  }, []);

  /**
   * Apply supplier name corrections to a list of certificates
   * Returns corrected certificates with matched supplier names
   */
  const applyCorrectedSupplierNames = useCallback((
    certificates: CertificateData[]
  ): { corrected: CertificateData[]; stats: { matched: number; newSuppliers: number } } => {
    if (!masterFile.isLoaded || Object.keys(masterFile.supplierMap).length === 0) {
      return {
        corrected: certificates,
        stats: { matched: 0, newSuppliers: certificates.length },
      };
    }

    let matchedCount = 0;
    let newSuppliersCount = 0;

    const corrected = certificates.map(cert => {
      const result = matchSupplier(cert.supplierName, masterFile.supplierMap, 0.75);

      if (result.wasMatched) {
        matchedCount++;
        return {
          ...cert,
          supplierName: result.matchedName,
          // Store original name for reference
          _originalSupplierName: cert.supplierName,
          _matchConfidence: result.confidence,
        } as CertificateData;
      } else {
        newSuppliersCount++;
        return cert;
      }
    });

    return {
      corrected,
      stats: { matched: matchedCount, newSuppliers: newSuppliersCount },
    };
  }, [masterFile]);

  /**
   * Check if a supplier name matches any in the Master File
   */
  const checkSupplierMatch = useCallback((supplierName: string) => {
    return matchSupplier(supplierName, masterFile.supplierMap, 0.75);
  }, [masterFile.supplierMap]);

  return {
    masterFile,
    isLoading,
    error,
    loadMasterFile,
    clearMasterFile,
    applyCorrectedSupplierNames,
    checkSupplierMatch,
  };
};
