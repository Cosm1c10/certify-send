import * as XLSX from 'xlsx';
import { CertificateData } from '@/types/certificate';

function calculateDaysToExpiry(expiryDate: string): number | null {
  if (!expiryDate || expiryDate === 'Not Found' || expiryDate === '') {
    return null;
  }

  const expiry = new Date(expiryDate);
  if (isNaN(expiry.getTime())) {
    return null;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  expiry.setHours(0, 0, 0, 0);

  const diffTime = expiry.getTime() - today.getTime();
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  return diffDays;
}

// Client terminology for status
function getClientStatus(daysToExpiry: number | null): string {
  if (daysToExpiry === null) {
    return 'Unknown';
  }

  if (daysToExpiry < 0) {
    return 'Expired'; // Matches Client Sheet
  } else if (daysToExpiry >= 0 && daysToExpiry < 30) {
    return 'Expiring Soon'; // Helpful warning
  } else {
    return 'Up to date'; // Matches Client Sheet EXACTLY
  }
}

export const exportToExcel = (certificates: CertificateData[]) => {
  const exportData = certificates.map(cert => {
    const daysToExpiry = calculateDaysToExpiry(cert.expiryDate);
    const status = getClientStatus(daysToExpiry);

    return {
      'Supplier Name': cert.supplierName || '',
      'Country': cert.country || '',
      'Product Category': cert.product || '',
      'Measure': cert.ecRegulation || '', // Renamed to match Master Sheet
      'Certification': cert.certification || '',
      'Issued': cert.issueDate || '', // Renamed to match Master Sheet
      'Date of Expiry': cert.expiryDate || '', // Renamed to match Master Sheet
      'Status': status, // Uses client terminology
      'Days to Expire': daysToExpiry !== null ? daysToExpiry : '', // Renamed to match Master Sheet
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(exportData, {
    header: [
      'Supplier Name',
      'Country',
      'Product Category',
      'Measure',
      'Certification',
      'Issued',
      'Date of Expiry',
      'Status',
      'Days to Expire',
    ],
  });

  const workbook = XLSX.utils.book_new();

  // Set column widths
  worksheet['!cols'] = [
    { wch: 30 }, // Supplier Name
    { wch: 15 }, // Country
    { wch: 20 }, // Product Category
    { wch: 35 }, // Measure (longer for full regulation text)
    { wch: 18 }, // Certification
    { wch: 12 }, // Issued
    { wch: 14 }, // Date of Expiry
    { wch: 14 }, // Status
    { wch: 14 }, // Days to Expire
  ];

  XLSX.utils.book_append_sheet(workbook, worksheet, 'Certificates');

  const date = new Date().toISOString().split('T')[0];
  XLSX.writeFile(workbook, `certificates-${date}.xlsx`);
};
