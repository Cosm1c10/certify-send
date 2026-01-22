export interface CertificateData {
  id: string;
  supplierName: string;
  certificateNumber: string;
  product: string;
  country: string;
  ecRegulation: string;
  certification: string;
  certType: string; // Kept for backwards compatibility
  issueDate: string;
  expiryDate: string;
  status: 'valid' | 'expired' | 'pending' | 'unknown';
  fileName: string;
}
