import { useState, useCallback } from 'react';
import { CertificateData } from '@/types/certificate';
import { supabase } from '@/integrations/supabase/client';
import { processWithOpenAI } from '@/utils/processWithOpenAI';
import { FileWithBase64 } from '@/components/DropZone';

interface EdgeFunctionResponse {
  supplier_name: string;
  certificate_number: string;
  country: string;
  product_category: string;
  ec_regulation: string;
  certification: string;
  date_issued: string;
  date_expired: string;
}

interface ProcessingProgress {
  current: number;
  total: number;
}

// Check if we should use local OpenAI (for development) or Supabase Edge Function (for production)
const USE_LOCAL_OPENAI = !!import.meta.env.VITE_OPENAI_API_KEY && import.meta.env.DEV;

console.log('OpenAI Mode:', USE_LOCAL_OPENAI ? 'LOCAL (direct API)' : 'PRODUCTION (Edge Function)');
console.log('VITE_OPENAI_API_KEY present:', !!import.meta.env.VITE_OPENAI_API_KEY);
console.log('DEV mode:', import.meta.env.DEV);

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

    return {
      id: crypto.randomUUID(),
      fileName: file.name,
      supplierName: data.supplier_name || '',
      certificateNumber: data.certificate_number || '',
      product: data.product_category || '',
      country: data.country || '',
      ecRegulation: data.ec_regulation || '',
      certification: data.certification || '',
      certType: data.certification || '',
      issueDate: data.date_issued || '',
      expiryDate: data.date_expired || '',
      status: determineStatus(data.date_expired),
    };
  };

  const analyzeCertificates = useCallback(async (files: FileWithBase64[]) => {
    if (files.length === 0) return;

    console.log(`Starting batch processing for ${files.length} files`);
    setIsProcessing(true);
    setProcessingProgress({ current: 0, total: files.length });
    setProcessingErrors([]);

    const errors: string[] = [];

    for (let i = 0; i < files.length; i++) {
      const { file, base64Image } = files[i];
      setProcessingProgress({ current: i + 1, total: files.length });

      try {
        console.log(`Processing ${i + 1}/${files.length}: ${file.name}`);
        const newCertificate = await processSingleCertificate(file, base64Image);
        setCertificates((prev) => [...prev, newCertificate]);
      } catch (error) {
        console.error(`Error processing ${file.name}:`, error);
        const message = error instanceof Error ? error.message : 'Unknown error';
        errors.push(`${file.name}: ${message}`);
      }
    }

    setIsProcessing(false);
    setProcessingProgress({ current: 0, total: 0 });

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
