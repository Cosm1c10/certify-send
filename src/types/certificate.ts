export interface CertificateData {
  id: string;
  supplierName: string;
  certificateNumber: string;
  country: string;
  scope: string;           // Symbol: "!" (Factory/System cert) or "+" (Product-specific cert)
  measure: string;         // Mapped regulation (e.g., "(EC) No 2023/2006", "EN 13432 (Compostable)")
  certification: string;   // Document type (e.g., "BRCGS", "ISO 9001", "DIN CERTCO")
  productCategory: string; // Product description (e.g., "Aqueous Coated Paper Cup", "PET Bottles")
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
