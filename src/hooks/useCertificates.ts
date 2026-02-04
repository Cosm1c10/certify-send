import { useState, useCallback, useRef } from 'react';
import { CertificateData } from '@/types/certificate';
import { supabase } from '@/integrations/supabase/client';
import { processWithOpenAI } from '@/utils/processWithOpenAI';
import { FileWithBase64 } from '@/components/DropZone';

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

// Check if we should use local OpenAI (for development) or Supabase Edge Function (for production)
const USE_LOCAL_OPENAI = !!import.meta.env.VITE_OPENAI_API_KEY && import.meta.env.DEV;

console.log('OpenAI Mode:', USE_LOCAL_OPENAI ? 'LOCAL (direct API)' : 'PRODUCTION (Edge Function)');

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

export const useCertificates = () => {
  const [certificates, setCertificates] = useState<CertificateData[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingProgress, setProcessingProgress] = useState<ProcessingProgress>({ current: 0, total: 0 });
  const [processingErrors, setProcessingErrors] = useState<string[]>([]);

  // Processing lock to prevent double-triggering
  const isProcessingRef = useRef(false);

  const processSingleCertificate = async (file: File, base64Image: string): Promise<CertificateData> => {
    let data: EdgeFunctionResponse;

    if (USE_LOCAL_OPENAI) {
      console.log('Calling OpenAI API directly for:', file.name);
      data = await processWithOpenAI(base64Image, file.name);
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

  const analyzeCertificates = useCallback(async (files: FileWithBase64[]) => {
    // Guard: Prevent double-triggering
    if (files.length === 0) return;
    if (isProcessingRef.current) {
      console.warn('Already processing - ignoring duplicate call');
      return;
    }

    // Lock processing
    isProcessingRef.current = true;
    console.log(`Starting batch processing for ${files.length} items`);

    setIsProcessing(true);
    setProcessingProgress({ current: 0, total: files.length });
    setProcessingErrors([]);

    const errors: string[] = [];
    const newCertificates: CertificateData[] = [];

    // Process ALL files first, collect results into local array
    for (let i = 0; i < files.length; i++) {
      const { file, base64Image } = files[i];
      setProcessingProgress({ current: i + 1, total: files.length });

      try {
        console.log(`Processing ${i + 1}/${files.length}: ${file.name}`);
        const newCertificate = await processSingleCertificate(file, base64Image);
        newCertificates.push(newCertificate);
      } catch (error) {
        console.error(`Error processing ${file.name}:`, error);
        const message = error instanceof Error ? error.message : 'Unknown error';
        errors.push(`${file.name}: ${message}`);
      }
    }

    // ATOMIC UPDATE: Update state ONCE with all new certificates
    if (newCertificates.length > 0) {
      setCertificates((prev) => [...prev, ...newCertificates]);
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

  return {
    certificates,
    isProcessing,
    processingProgress,
    processingErrors,
    analyzeCertificates,
    clearCertificates,
  };
};
