import { useState, useCallback } from 'react';
import { CertificateData } from '@/types/certificate';

// Mock function to simulate API response - replace with actual Edge Function call
const mockAnalyzeCertificate = async (fileName: string): Promise<Omit<CertificateData, 'id' | 'fileName'>> => {
  // Simulate API delay
  await new Promise(resolve => setTimeout(resolve, 1500));
  
  // Mock data - in production, this would come from your Edge Function
  const mockResponses = [
    {
      supplierName: 'Global Organics Ltd',
      product: 'Organic Wheat Flour',
      country: 'Germany',
      certType: 'ISO 22000',
      issueDate: '2024-01-15',
      expiryDate: '2025-01-14',
      status: 'valid' as const,
    },
    {
      supplierName: 'Pacific Foods Inc',
      product: 'Premium Rice',
      country: 'Japan',
      certType: 'FSSC 22000',
      issueDate: '2023-06-01',
      expiryDate: '2024-05-31',
      status: 'expired' as const,
    },
    {
      supplierName: 'Nordic Harvest AB',
      product: 'Oat Bran',
      country: 'Sweden',
      certType: 'BRC',
      issueDate: '2024-03-10',
      expiryDate: '2025-03-09',
      status: 'valid' as const,
    },
  ];
  
  return mockResponses[Math.floor(Math.random() * mockResponses.length)];
};

export const useCertificates = () => {
  const [certificates, setCertificates] = useState<CertificateData[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);

  const analyzeCertificate = useCallback(async (file: File, base64Image: string) => {
    setIsProcessing(true);
    
    try {
      // TODO: Replace with actual Edge Function call
      // const response = await fetch('/api/analyze-cert', {
      //   method: 'POST',
      //   headers: { 'Content-Type': 'application/json' },
      //   body: JSON.stringify({ image: base64Image }),
      // });
      // const data = await response.json();
      
      const data = await mockAnalyzeCertificate(file.name);
      
      const newCertificate: CertificateData = {
        id: crypto.randomUUID(),
        fileName: file.name,
        ...data,
      };
      
      setCertificates(prev => [...prev, newCertificate]);
    } catch (error) {
      console.error('Error analyzing certificate:', error);
      alert('Failed to analyze certificate. Please try again.');
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
