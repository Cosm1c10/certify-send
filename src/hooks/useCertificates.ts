import { useState, useCallback, useRef } from 'react';
import { CertificateData } from '@/types/certificate';
import { supabase } from '@/integrations/supabase/client';
import { processWithOpenAI } from '@/utils/processWithOpenAI';
import { FileWithBase64 } from '@/components/DropZone';
import { DynamicSupplierMap, matchSupplier } from '@/utils/masterFileParser';

interface EdgeFunctionResponse {
  supplier_name: string;
  certificate_number: string;
  country?: string;
  scope?: string;           // Product description
  measure?: string;         // Regulation reference
  certification: string;
  product_category?: string;
  date_issued: string;
  date_expired?: string | null;
  // Legacy field aliases
  ec_regulation?: string;   // Alias for measure
  region?: string;          // Alias for country
  date_expiry?: string;     // Alias for date_expired
}

interface ProcessingProgress {
  current: number;
  total: number;
}

// Call OpenAI directly whenever VITE_OPENAI_API_KEY is present — both in local
// development AND in production (Vercel env vars). This bypasses the Supabase
// Edge Function entirely, eliminating payload-size limits and aggressive timeouts
// that cause "Failed to send a request to the Edge Function" on large PDFs.
// If the key is not set, we fall back to the Supabase Edge Function.
const USE_LOCAL_OPENAI = !!import.meta.env.VITE_OPENAI_API_KEY;

console.log('OpenAI Mode:', USE_LOCAL_OPENAI ? 'DIRECT (browser → OpenAI)' : 'EDGE (Supabase function)');

function determineStatus(expiryDate: string): CertificateData['status'] {
  if (!expiryDate || expiryDate === 'Not Found') {
    return 'unknown';
  }

  const expiry = new Date(expiryDate);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (isNaN(expiry.getTime())) {
    return 'unknown';
  }

  return expiry >= today ? 'valid' : 'expired';
}

export const useCertificates = (supplierMap?: DynamicSupplierMap) => {
  // Keep a ref so the latest supplierMap is always available inside callbacks
  const supplierMapRef = useRef<DynamicSupplierMap | undefined>(supplierMap);
  supplierMapRef.current = supplierMap;
  const [certificates, setCertificates] = useState<CertificateData[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingProgress, setProcessingProgress] = useState<ProcessingProgress>({ current: 0, total: 0 });
  const [processingErrors, setProcessingErrors] = useState<string[]>([]);

  // Processing lock to prevent double-triggering
  const isProcessingRef = useRef(false);

  const processSingleCertificate = async (file: File, base64Image: string, textContent?: string): Promise<CertificateData> => {
    let data: EdgeFunctionResponse;
    const isDocx = !!textContent;

    if (USE_LOCAL_OPENAI) {
      console.log('Calling OpenAI API directly for:', file.name, isDocx ? '(DOCX)' : '');
      data = await processWithOpenAI(base64Image, file.name, textContent);
    } else {
      if (isDocx) {
        // For DOCX files, send text content to a text-processing endpoint
        // Note: Edge function would need to be updated to handle text input
        const { data: responseData, error } = await supabase.functions.invoke<EdgeFunctionResponse>(
          'process-certificate',
          {
            body: { text: textContent, filename: file.name },
          }
        );

        if (error) {
          throw new Error(error.message || 'Failed to process certificate');
        }

        if (!responseData) {
          throw new Error('No data returned from Edge Function');
        }

        data = responseData;
      } else {
        const base64Data = base64Image.replace(/^data:image\/\w+;base64,/, '');
        const { data: responseData, error } = await supabase.functions.invoke<EdgeFunctionResponse>(
          'process-certificate',
          {
            body: { image: base64Data, filename: file.name },
          }
        );

        if (error) {
          throw new Error(error.message || 'Failed to process certificate');
        }

        if (!responseData) {
          throw new Error('No data returned from Edge Function');
        }

        data = responseData;
      }
    }

    // Handle field aliases for backwards compatibility
    const expiryDate = data.date_expired || data.date_expiry || '';
    const country = data.country || data.region || '';
    const measure = data.measure || data.ec_regulation || '';

    return {
      id: crypto.randomUUID(),
      fileName: file.name,
      supplierName: data.supplier_name || '',
      certificateNumber: data.certificate_number || '',
      country: country,
      scope: data.scope || '',                    // Product description
      measure: measure,                            // Regulation reference
      certification: data.certification || '',
      productCategory: data.product_category || '',
      issueDate: data.date_issued || '',
      expiryDate: expiryDate,
      status: determineStatus(expiryDate),
      // Legacy fields for backwards compatibility
      product: data.scope || data.product_category || '',
      ecRegulation: measure,
      certType: data.certification || '',
    };
  };

  const analyzeCertificates = useCallback(async (
    files: FileWithBase64[],
    supplierOverride?: string,
    supplierAccountOverride?: string   // explicit account code for brand-new suppliers
  ) => {
    // Guard: Prevent double-triggering
    if (files.length === 0) return;
    if (isProcessingRef.current) {
      console.warn('Already processing - ignoring duplicate call');
      return;
    }

    // Lock processing
    isProcessingRef.current = true;
    const overrideName    = supplierOverride?.trim()        || '';
    const overrideAccount = supplierAccountOverride?.trim() || '';
    console.log(
      `Starting batch processing for ${files.length} items` +
      (overrideName    ? ` [Override: "${overrideName}"]`       : '') +
      (overrideAccount ? ` [Account: "${overrideAccount}"]`     : '')
    );

    setIsProcessing(true);
    setProcessingProgress({ current: 0, total: files.length });
    setProcessingErrors([]);

    const errors: string[] = [];
    const newCertificates: CertificateData[] = [];

    // Process ALL files first, collect results into local array
    for (let i = 0; i < files.length; i++) {
      const { file, base64Image, textContent } = files[i];
      setProcessingProgress({ current: i + 1, total: files.length });

      try {
        console.log(`Processing ${i + 1}/${files.length}: ${file.name}`);
        const newCertificate = await processSingleCertificate(file, base64Image, textContent);

        // SUPPLIER OVERRIDE: If the user explicitly assigned a supplier, use that name
        // regardless of what the AI extracted (broker/manufacturer workflow)
        if (overrideName) {
          newCertificate.supplierName = overrideName;

          if (overrideAccount) {
            // User typed an explicit account code (new supplier flow) — use it directly
            // without fuzzy-matching so the exact code lands in col A of the Excel.
            (newCertificate as any)._matchedAccount = overrideAccount;
          } else {
            // Try to find the account code from the Master File for existing suppliers
            const currentMap = supplierMapRef.current;
            if (currentMap) {
              const result = matchSupplier(overrideName, currentMap, 0.75);
              if (result.wasMatched) {
                (newCertificate as any)._matchedAccount = result.matchedAccount ?? '';
              }
            }
          }
        }

        newCertificates.push(newCertificate);
      } catch (error) {
        console.error(`Error processing ${file.name}:`, error);
        const message = error instanceof Error ? error.message : 'Unknown error';
        errors.push(`${file.name}: ${message}`);
      }
    }

    // SUPPLIER NAME CORRECTION: Apply Master File matching before saving
    // Skip if supplierOverride is set — user already chose the correct name
    const currentMap = supplierMapRef.current;
    const mapSize = currentMap ? Object.keys(currentMap).length : 0;
    const correctedCertificates = (!overrideName && currentMap && mapSize > 0)
      ? newCertificates.map(cert => {
          const result = matchSupplier(cert.supplierName, currentMap, 0.75);
          if (result.wasMatched && result.matchedName !== cert.supplierName) {
            console.log(`Supplier corrected: "${cert.supplierName}" → "${result.matchedName}"`);
            return { ...cert, supplierName: result.matchedName };
          }
          return cert;
        })
      : newCertificates;

    // ATOMIC UPDATE: Update state ONCE with all corrected certificates
    if (correctedCertificates.length > 0) {
      setCertificates((prev) => [...prev, ...correctedCertificates]);
    }

    // Cleanup
    setIsProcessing(false);
    setProcessingProgress({ current: 0, total: 0 });
    isProcessingRef.current = false; // Unlock

    if (errors.length > 0) {
      setProcessingErrors(errors);
    }
  }, []);

  const clearCertificates = useCallback(() => {
    setCertificates([]);
    setProcessingErrors([]);
  }, []);

  // Allow external updates to certificates (e.g., supplier name correction)
  const updateCertificates = useCallback((updater: (prev: CertificateData[]) => CertificateData[]) => {
    setCertificates(updater);
  }, []);

  return {
    certificates,
    isProcessing,
    processingProgress,
    processingErrors,
    analyzeCertificates,
    clearCertificates,
    updateCertificates,
  };
};
