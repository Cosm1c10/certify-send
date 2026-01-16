import { useState, useCallback } from 'react';
import { CertificateData } from '@/types/certificate';
import { supabase } from '@/integrations/supabase/client';
import { processWithOpenAI } from '@/utils/processWithOpenAI';

interface EdgeFunctionResponse {
  supplier_name: string;
  country: string;
  product_category: string;
  ec_regulation: string;
  certification: string;
  date_issued: string;
  date_expired: string;
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

  const analyzeCertificate = useCallback(async (file: File, base64Image: string) => {
    console.log('analyzeCertificate called with file:', file.name);
    console.log('base64Image length:', base64Image.length);
    setIsProcessing(true);

    try {
      let data: EdgeFunctionResponse;

      if (USE_LOCAL_OPENAI) {
        // Development: Call OpenAI directly from the browser
        console.log('Calling OpenAI API directly...');
        data = await processWithOpenAI(base64Image);
        console.log('OpenAI response:', data);
      } else {
        // Production: Use Supabase Edge Function
        const base64Data = base64Image.replace(/^data:image\/\w+;base64,/, '');
        const { data: responseData, error } = await supabase.functions.invoke<EdgeFunctionResponse>(
          'process-certificate',
          {
            body: { image: base64Data },
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

      const newCertificate: CertificateData = {
        id: crypto.randomUUID(),
        fileName: file.name,
        supplierName: data.supplier_name || '',
        product: data.product_category || '',
        country: data.country || '',
        ecRegulation: data.ec_regulation || '',
        certification: data.certification || '',
        certType: data.certification || '', // Kept for backwards compatibility
        issueDate: data.date_issued || '',
        expiryDate: data.date_expired || '',
        status: determineStatus(data.date_expired),
      };

      setCertificates((prev) => [...prev, newCertificate]);
    } catch (error) {
      console.error('Error analyzing certificate:', error);
      const message = error instanceof Error ? error.message : 'Unknown error';
      alert(`Failed to analyze certificate: ${message}`);
    } finally {
      setIsProcessing(false);
    }
  }, []);

  return {
    certificates,
    isProcessing,
    analyzeCertificate,
  };
};
