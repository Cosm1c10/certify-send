export interface CertificateData {
  id: string;
  supplierName: string;
  certificateNumber: string;
  country: string;
  scope: string;           // Product description (e.g., "Paper Cup", "PET Bottles")
  measure: string;         // Regulation reference (e.g., "Commission Regulation (EU) No 10/2011")
  certification: string;   // Document type (e.g., "BRCGS", "ISO 9001", "DoC")
  productCategory: string; // Material type (e.g., "Paper", "Rigid Plastics")
  issueDate: string;
  expiryDate: string;
  status: 'valid' | 'expired' | 'pending' | 'unknown';
  fileName: string;
  // Legacy fields for backwards compatibility
  product?: string;
  ecRegulation?: string;
  certType?: string;
}

// API response from the Edge Function
export interface ExtractedCertificateData {
  supplier_name: string;
  certificate_number: string;
  country: string;
  scope: string;
  measure: string;
  certification: string;
  product_category: string;
  date_issued: string;
  date_expired: string | null;
}
