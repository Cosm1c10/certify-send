export interface CertificateData {
  id: string;
  supplierName: string;
  product: string;
  country: string;
  certType: string;
  issueDate: string;
  expiryDate: string;
  status: 'valid' | 'expired' | 'pending' | 'unknown';
  fileName: string;
}
