import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';
import { DynamicSupplierMap, matchSupplier } from './masterFileParser';

// =============================================================================
// FEEDER EXCEL EXPORTER
// Creates a separate "feeder" file for copy-paste import to Master File
// Applies supplier name auto-correction and flags new suppliers
// =============================================================================

/**
 * Apply the 3-Year Rule for missing expiry dates
 */
function applyThreeYearRule(issueDate: string, expiryDate: string | null | undefined): string {
  if (expiryDate && expiryDate !== 'Not Found' && expiryDate !== '' && expiryDate !== '-') {
    return expiryDate;
  }

  if (!issueDate || issueDate === 'Not Found' || issueDate === '' || issueDate === '-') {
    return '';
  }

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
  if (!expiryDate || expiryDate === 'Not Found' || expiryDate === '' || expiryDate === '-') {
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
 * NOTE: Client's Master File requires "Up to date" (not "Valid") for conditional formatting
 */
function getStatus(daysToExpiry: number | null): string {
  if (daysToExpiry === null) {
    return 'Unknown';
  }
  return daysToExpiry < 0 ? 'Expired' : 'Up to date';
}

/**
 * Check if a certificate has a valid expiry date
 */
function hasValidExpiry(cert: CertificateData): boolean {
  const expiry = cert.expiryDate || '';
  const issue = cert.issueDate || '';

  if (expiry && expiry !== 'Not Found' && expiry !== '' && expiry !== '-') {
    return true;
  }

  if (issue && issue !== 'Not Found' && issue !== '' && issue !== '-') {
    return true;
  }

  return false;
}

/**
 * THE "SAUREBH FILTER" - Deduplication Logic
 * Group by: Supplier Name + Certification + Measure
 * Keep the one with a valid expiry date (latest expiry wins)
 */
function applySaurabhFilter(certificates: CertificateData[]): CertificateData[] {
  const groups = new Map<string, CertificateData[]>();

  for (const cert of certificates) {
    const supplierName = (cert.supplierName || '').trim().toLowerCase();
    const certification = (cert.certification || '').trim().toLowerCase();
    const measure = (cert.measure || cert.ecRegulation || '').trim().toLowerCase();
    const key = `${supplierName}|${certification}|${measure}`;

    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key)!.push(cert);
  }

  const deduplicated: CertificateData[] = [];

  for (const [key, certs] of groups) {
    if (certs.length === 1) {
      deduplicated.push(certs[0]);
    } else {
      console.log(`Saurebh Filter: Found ${certs.length} duplicates for "${key}"`);

      const sorted = [...certs].sort((a, b) => {
        const aHasExpiry = hasValidExpiry(a);
        const bHasExpiry = hasValidExpiry(b);

        if (aHasExpiry && !bHasExpiry) return -1;
        if (!aHasExpiry && bHasExpiry) return 1;

        if (aHasExpiry && bHasExpiry) {
          const aExpiry = applyThreeYearRule(a.issueDate || '', a.expiryDate || '');
          const bExpiry = applyThreeYearRule(b.issueDate || '', b.expiryDate || '');
          return bExpiry.localeCompare(aExpiry);
        }

        return 0;
      });

      deduplicated.push(sorted[0]);
      console.log(`Saurebh Filter: Kept 1, removed ${certs.length - 1} duplicate(s)`);
    }
  }

  return deduplicated;
}

interface FeederExportStats {
  total: number;
  matched: number;
  newSuppliers: number;
  duplicatesRemoved: number;
}

/**
 * Export Feeder Excel file with supplier name auto-correction
 *
 * This creates a SEPARATE file for copy-paste import to the Master File.
 * The Master File is READ-ONLY (to preserve conditional formatting).
 */
export const exportFeederExcel = async (
  certificates: CertificateData[],
  supplierMap: DynamicSupplierMap = {}
): Promise<FeederExportStats> => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Certificate Analyzer - Feeder System';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('New Certificates');

  // ==========================================================================
  // STEP 1: Apply "Saurebh Filter" Deduplication
  // ==========================================================================
  const deduplicated = applySaurabhFilter(certificates);
  const duplicatesRemoved = certificates.length - deduplicated.length;
  console.log(`Saurebh Filter: ${certificates.length} -> ${deduplicated.length} certificates`);

  // ==========================================================================
  // STEP 2: Apply Supplier Name Auto-Correction
  // ==========================================================================
  let matchedCount = 0;
  let newSuppliersCount = 0;
  const hasSupplierMap = Object.keys(supplierMap).length > 0;

  const corrected = deduplicated.map(cert => {
    if (!hasSupplierMap) {
      newSuppliersCount++;
      return { ...cert, _isNewSupplier: true, _matchedAccount: '' };
    }

    const result = matchSupplier(cert.supplierName, supplierMap, 0.75);

    if (result.wasMatched) {
      matchedCount++;
      return {
        ...cert,
        supplierName: result.matchedName,
        _originalSupplierName: cert.supplierName,
        _matchConfidence: result.confidence,
        _isNewSupplier: false,
        _matchedAccount: result.matchedAccount || '',
      };
    } else {
      newSuppliersCount++;
      return { ...cert, _isNewSupplier: true, _matchedAccount: '' };
    }
  });

  // ==========================================================================
  // STEP 3: Sort Alphabetically by Supplier Name
  // ==========================================================================
  const sortedCertificates = [...corrected].sort((a, b) => {
    const nameA = (a.supplierName || '').toLowerCase();
    const nameB = (b.supplierName || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  // ==========================================================================
  // STEP 4: Define V7 Master Columns (Feeder format)
  // ==========================================================================
  worksheet.columns = [
    { header: 'Supplier Account', key: 'supplierAccount', width: 18 },
    { header: 'Supplier Name', key: 'supplierName', width: 30 },
    { header: 'Country', key: 'country', width: 15 },
    { header: 'Scope', key: 'scope', width: 10 },
    { header: 'Measure', key: 'measure', width: 40 },
    { header: 'Certification', key: 'certification', width: 25 },
    { header: 'Product Category', key: 'productCategory', width: 18 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Issued', key: 'issued', width: 12 },
    { header: 'Date of Expiry', key: 'dateOfExpiry', width: 14 },
    { header: 'Days to Expire', key: 'daysToExpire', width: 14 },
    { header: 'Contact person & Email', key: 'contactEmail', width: 28 },
    { header: 'Date Request sent', key: 'dateRequestSent', width: 18 },
    { header: 'Date Received', key: 'dateReceived', width: 14 },
    { header: 'Comments', key: 'comments', width: 25 },
    // Extra columns for feeder system
    { header: 'NEW SUPPLIER?', key: 'newSupplier', width: 14 },
    { header: 'Original Name', key: 'originalName', width: 30 },
  ];

  // ==========================================================================
  // STEP 5: Style the Header Row
  // ==========================================================================
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' },
  };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
  headerRow.height = 25;

  // Special styling for "NEW SUPPLIER?" column
  headerRow.getCell('newSupplier').fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF9C27B0' }, // Purple for attention
  };

  // ==========================================================================
  // STEP 6: Add Data Rows
  // ==========================================================================
  sortedCertificates.forEach((cert: any) => {
    const issueDate = cert.issueDate || '';
    const rawExpiryDate = cert.expiryDate || '';
    const effectiveExpiryDate = applyThreeYearRule(issueDate, rawExpiryDate);
    const daysToExpiry = calculateDaysToExpiry(effectiveExpiryDate);
    const status = getStatus(daysToExpiry);

    worksheet.addRow({
      supplierAccount: cert._matchedAccount || '',
      supplierName: cert.supplierName || '',
      country: cert.country || '',
      scope: cert.scope || '',
      measure: cert.measure || cert.ecRegulation || '',
      certification: cert.certification || '',
      productCategory: cert.productCategory || cert.product || '',
      status: status,
      issued: issueDate,
      dateOfExpiry: effectiveExpiryDate,
      daysToExpire: daysToExpiry !== null ? daysToExpiry : '',
      contactEmail: '',
      dateRequestSent: '',
      dateReceived: '',
      comments: '',
      newSupplier: cert._isNewSupplier ? 'YES - NEW' : '',
      originalName: cert._originalSupplierName || '',
    });
  });

  // ==========================================================================
  // STEP 7: Apply Conditional Formatting
  // ==========================================================================
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const statusCell = row.getCell('status');
    const daysCell = row.getCell('daysToExpire');
    const newSupplierCell = row.getCell('newSupplier');
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
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' },
      };
      statusCell.font = { bold: true, color: { argb: 'FF9C0006' } };

      daysCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCCCC' },
      };
      daysCell.font = { bold: true, color: { argb: 'FF9C0006' } };
    } else if (status === 'Up to date') {
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' },
      };
      statusCell.font = { bold: true, color: { argb: 'FF006100' } };

      daysCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' },
      };
      daysCell.font = { bold: true, color: { argb: 'FF006100' } };
    } else {
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFEB9C' },
      };
      statusCell.font = { bold: true, color: { argb: 'FF9C5700' } };
    }

    // Highlight NEW SUPPLIER cells
    if (newSupplierCell.value) {
      newSupplierCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE1BEE7' }, // Light purple
      };
      newSupplierCell.font = { bold: true, color: { argb: 'FF6A1B9A' } };
    }
  });

  // Alternating row colors for readability
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (rowNumber % 2 === 0) {
      row.eachCell((cell, colNumber) => {
        // Don't override special columns
        if (colNumber !== 8 && colNumber !== 11 && colNumber !== 16) {
          const currentFill = cell.fill as ExcelJS.FillPattern;
          if (!currentFill || currentFill.pattern !== 'solid') {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: 'FFF5F5F5' },
            };
          }
        }
      });
    }
  });

  // Freeze header row
  worksheet.views = [{ state: 'frozen', ySplit: 1 }];

  // ==========================================================================
  // STEP 8: Generate and Download
  // ==========================================================================
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const date = new Date().toISOString().split('T')[0];
  saveAs(blob, `New_Certificates_Import_${date}.xlsx`);

  const stats: FeederExportStats = {
    total: sortedCertificates.length,
    matched: matchedCount,
    newSuppliers: newSuppliersCount,
    duplicatesRemoved,
  };

  console.log('Feeder Export Stats:', stats);
  return stats;
};
