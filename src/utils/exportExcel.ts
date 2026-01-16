import * as XLSX from 'xlsx';
import { CertificateData } from '@/types/certificate';

export const exportToExcel = (certificates: CertificateData[]) => {
  const exportData = certificates.map(cert => ({
    'Supplier Name': cert.supplierName,
    'Product': cert.product,
    'Country': cert.country,
    'Certificate Type': cert.certType,
    'Issue Date': cert.issueDate,
    'Expiry Date': cert.expiryDate,
    'Status': cert.status.charAt(0).toUpperCase() + cert.status.slice(1),
  }));

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  
  // Set column widths
  worksheet['!cols'] = [
    { wch: 20 }, // Supplier Name
    { wch: 20 }, // Product
    { wch: 15 }, // Country
    { wch: 15 }, // Certificate Type
    { wch: 12 }, // Issue Date
    { wch: 12 }, // Expiry Date
    { wch: 10 }, // Status
  ];

  XLSX.utils.book_append_sheet(workbook, worksheet, 'Certificates');
  
  const date = new Date().toISOString().split('T')[0];
  XLSX.writeFile(workbook, `certificates-${date}.xlsx`);
};
