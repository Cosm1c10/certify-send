import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';

/**
 * Apply the 3-Year Rule for missing expiry dates
 * If date_expired is null/empty but date_issued exists, set expiry = issued + 3 years
 */
function applyThreeYearRule(issueDate: string, expiryDate: string | null | undefined): string {
  // If expiry date exists and is valid, use it
  if (expiryDate && expiryDate !== 'Not Found' && expiryDate !== '') {
    return expiryDate;
  }

  // If no issue date, we can't calculate
  if (!issueDate || issueDate === 'Not Found' || issueDate === '') {
    return '';
  }

  // Apply 3-year rule
  const issued = new Date(issueDate);
  if (isNaN(issued.getTime())) {
    return '';
  }

  const calculatedExpiry = new Date(issued);
  calculatedExpiry.setFullYear(calculatedExpiry.getFullYear() + 3);
  return calculatedExpiry.toISOString().split('T')[0];
}

/**
 * Calculate days until expiry
 */
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
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

/**
 * Determine status based on expiry date
 * If expiry < today -> "Expired", else "Up to date"
 */
function getStatus(daysToExpiry: number | null): string {
  if (daysToExpiry === null) {
    return 'Unknown';
  }
  return daysToExpiry < 0 ? 'Expired' : 'Up to date';
}

/**
 * Generate deduplication key
 * ALWAYS includes: Certificate Number + Certification + Measure
 * This ensures different standards with same cert number are NOT treated as duplicates
 */
function getDeduplicationKey(cert: CertificateData): string {
  const certNumber = (cert.certificateNumber || '').trim().toLowerCase();
  const certification = (cert.certification || '').trim().toLowerCase();
  const measure = (cert.measure || cert.ecRegulation || '').trim().toLowerCase();
  const supplier = (cert.supplierName || '').trim().toLowerCase();

  // ALWAYS include certification and measure in the key to prevent false dedup
  // e.g., same cert number but different standards = different rows
  if (certNumber && certNumber !== 'not found' && certNumber !== '-') {
    return `cert:${certNumber}|${certification}|${measure}`;
  }

  // Fallback: Supplier Name + Measure + Certification
  return `combo:${supplier}|${measure}|${certification}`;
}

export const exportToExcel = async (certificates: CertificateData[]) => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Certificate Analyzer';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Certificates');

  // STEP 1: Deduplicate by Supplier Name + Measure + Scope (keep first occurrence)
  const seen = new Set<string>();
  const uniqueCertificates = certificates.filter((cert) => {
    const key = getDeduplicationKey(cert);

    if (seen.has(key)) {
      console.log(`Duplicate skipped: ${key}`);
      return false;
    }
    seen.add(key);
    return true;
  });

  // STEP 2: Sort by supplier_name (A-Z), then by measure (A-Z)
  const sortedCertificates = [...uniqueCertificates].sort((a, b) => {
    const nameA = (a.supplierName || '').toLowerCase();
    const nameB = (b.supplierName || '').toLowerCase();
    const nameCompare = nameA.localeCompare(nameB);

    if (nameCompare !== 0) return nameCompare;

    const measureA = (a.measure || a.ecRegulation || '').toLowerCase();
    const measureB = (b.measure || b.ecRegulation || '').toLowerCase();
    return measureA.localeCompare(measureB);
  });

  // Define the 15 columns matching client's Master File structure
  worksheet.columns = [
    { header: 'Supplier Account', key: 'supplierAccount', width: 18 },
    { header: 'Supplier Name', key: 'supplierName', width: 30 },
    { header: 'Country', key: 'country', width: 15 },
    { header: 'Scope', key: 'scope', width: 35 },
    { header: 'Measure', key: 'measure', width: 40 },
    { header: 'Certification', key: 'certification', width: 22 },
    { header: 'Product Category', key: 'productCategory', width: 18 },
    { header: 'Status', key: 'status', width: 14 },
    { header: 'Issued', key: 'issued', width: 12 },
    { header: 'Date of Expiry', key: 'dateOfExpiry', width: 14 },
    { header: 'Days to Expire', key: 'daysToExpire', width: 14 },
    { header: 'Contact person & Email', key: 'contactEmail', width: 28 },
    { header: 'Date Request sent', key: 'dateRequestSent', width: 18 },
    { header: 'Date Received', key: 'dateReceived', width: 14 },
    { header: 'Comments', key: 'comments', width: 25 },
  ];

  // Style the header row
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' }, // Blue header
  };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
  headerRow.height = 25;

  // Today's date for "Date Received" column
  const todayStr = new Date().toISOString().split('T')[0];

  // Add data rows
  sortedCertificates.forEach((cert) => {
    // Apply 3-year rule to get effective expiry date
    const effectiveExpiryDate = applyThreeYearRule(
      cert.issueDate,
      cert.expiryDate
    );

    const daysToExpiry = calculateDaysToExpiry(effectiveExpiryDate);
    const status = getStatus(daysToExpiry);

    worksheet.addRow({
      supplierAccount: '',                                           // Column 1: Empty
      supplierName: cert.supplierName || '',                         // Column 2: Supplier Name
      country: cert.country || '',                                   // Column 3: Country
      scope: cert.scope || cert.product || '',                       // Column 4: Scope (fallback to product)
      measure: cert.measure || cert.ecRegulation || '',              // Column 5: Measure (fallback to ecRegulation)
      certification: cert.certification || '',                        // Column 6: Certification
      productCategory: cert.productCategory || '',                   // Column 7: Product Category
      status: status,                                                // Column 8: Status
      issued: cert.issueDate || '',                                  // Column 9: Issued
      dateOfExpiry: effectiveExpiryDate,                            // Column 10: Date of Expiry (with 3-year rule)
      daysToExpire: daysToExpiry !== null ? daysToExpiry : '',      // Column 11: Days to Expire
      contactEmail: '',                                              // Column 12: Empty
      dateRequestSent: '',                                           // Column 13: Empty
      dateReceived: todayStr,                                        // Column 14: Today's date
      comments: '',                                                  // Column 15: Empty
    });
  });

  // Apply conditional formatting to data rows
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header

    const statusCell = row.getCell('status');
    const daysCell = row.getCell('daysToExpire');
    const status = statusCell.value as string;

    // Apply borders to all cells
    row.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
        right: { style: 'thin', color: { argb: 'FFD9D9D9' } },
      };
    });

    // Status-specific styling
    if (status === 'Expired') {
      // Red background for expired
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' }, // Light red
      };
      statusCell.font = { bold: true, color: { argb: 'FF9C0006' } }; // Dark red text

      daysCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' },
      };
      daysCell.font = { bold: true, color: { argb: 'FF9C0006' } };
    } else if (status === 'Up to date') {
      // Green background for valid
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' }, // Light green
      };
      statusCell.font = { bold: true, color: { argb: 'FF006100' } }; // Dark green text

      daysCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' },
      };
      daysCell.font = { bold: true, color: { argb: 'FF006100' } };
    } else {
      // Gray for unknown
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE0E0E0' },
      };
      statusCell.font = { color: { argb: 'FF666666' } };
    }
  });

  // Add alternating row colors for readability (skip status/days columns 8 and 11)
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (rowNumber % 2 === 0) {
      row.eachCell((cell, colNumber) => {
        // Don't override status (col 8) and days to expire (col 11) cells
        if (colNumber !== 8 && colNumber !== 11) {
          if (!cell.fill || (cell.fill as ExcelJS.FillPattern).pattern !== 'solid') {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFF5F5F5' }, // Light gray for even rows
            };
          }
        }
      });
    }
  });

  // Freeze the header row
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  // Generate buffer and save
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const date = new Date().toISOString().split('T')[0];
  saveAs(blob, `certificates-${date}.xlsx`);
};
