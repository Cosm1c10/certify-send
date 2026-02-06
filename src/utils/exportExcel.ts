import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';
import { DynamicSupplierMap, matchSupplier } from './masterFileParser';

// =============================================================================
// V7 MASTER FILE EXCEL EXPORTER
// Implements exact client column structure + "Saurebh Filter" deduplication
// + NEW SUPPLIER DETECTION for Feeder System integration
// =============================================================================

/**
 * Represents a new/unknown supplier not found in the Master File
 */
export interface NewSupplierInfo {
  supplierName: string;
  fileName: string;
  certification: string;
}

/**
 * Result of preparing export data with supplier matching
 */
export interface PrepareExportResult {
  processedCertificates: CertificateData[];
  newSuppliers: NewSupplierInfo[];
  matchedCount: number;
  duplicatesRemoved: number;
}

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

  // Has direct expiry
  if (expiry && expiry !== 'Not Found' && expiry !== '' && expiry !== '-') {
    return true;
  }

  // Has issue date (can calculate expiry via 3-year rule)
  if (issue && issue !== 'Not Found' && issue !== '' && issue !== '-') {
    return true;
  }

  return false;
}

/**
 * THE "SAUREBH FILTER" - Deduplication Logic
 *
 * Rules:
 * 1. Group by: Supplier Name + Certification
 * 2. If duplicates exist, keep ONLY ONE
 * 3. Priority: Prefer the one with a valid Date of Expiry
 */
function applySaurabhFilter(certificates: CertificateData[]): CertificateData[] {
  // Group certificates by Supplier Name + Certification
  const groups = new Map<string, CertificateData[]>();

  for (const cert of certificates) {
    const supplierName = (cert.supplierName || '').trim().toLowerCase();
    const certification = (cert.certification || '').trim().toLowerCase();
    const key = `${supplierName}|${certification}`;

    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key)!.push(cert);
  }

  // For each group, select the best certificate
  const deduplicated: CertificateData[] = [];

  for (const [key, certs] of groups) {
    if (certs.length === 1) {
      // No duplicates, keep as-is
      deduplicated.push(certs[0]);
    } else {
      // Multiple entries - apply priority rules
      console.log(`Saurebh Filter: Found ${certs.length} duplicates for "${key}"`);

      // Sort by priority: valid expiry first, then by expiry date (latest first)
      const sorted = [...certs].sort((a, b) => {
        const aHasExpiry = hasValidExpiry(a);
        const bHasExpiry = hasValidExpiry(b);

        // Prefer one with valid expiry
        if (aHasExpiry && !bHasExpiry) return -1;
        if (!aHasExpiry && bHasExpiry) return 1;

        // If both have expiry, prefer the one with later expiry date
        if (aHasExpiry && bHasExpiry) {
          const aExpiry = applyThreeYearRule(a.issueDate || '', a.expiryDate || '');
          const bExpiry = applyThreeYearRule(b.issueDate || '', b.expiryDate || '');
          return bExpiry.localeCompare(aExpiry); // Later date first
        }

        return 0;
      });

      // Keep only the best one
      deduplicated.push(sorted[0]);
      console.log(`Saurebh Filter: Kept 1, removed ${certs.length - 1} duplicate(s)`);
    }
  }

  return deduplicated;
}

/**
 * PREPARE EXPORT DATA - Separates data processing from file generation
 *
 * This function:
 * 1. Applies Saurebh Filter deduplication
 * 2. Matches suppliers against the Master File (if provided)
 * 3. Tracks NEW/UNKNOWN suppliers not in Master File
 * 4. Returns processed data for export + new supplier warnings
 */
export function prepareExportData(
  certificates: CertificateData[],
  supplierMap: DynamicSupplierMap = {}
): PrepareExportResult {
  // Step 1: Apply deduplication
  const deduplicated = applySaurabhFilter(certificates);
  const duplicatesRemoved = certificates.length - deduplicated.length;
  console.log(`Saurebh Filter: ${certificates.length} -> ${deduplicated.length} certificates`);

  // Step 2: Match against Master File and track new suppliers
  const hasSupplierMap = Object.keys(supplierMap).length > 0;
  const newSuppliers: NewSupplierInfo[] = [];
  const seenNewSuppliers = new Set<string>(); // Avoid duplicate warnings
  let matchedCount = 0;

  const processedCertificates = deduplicated.map(cert => {
    if (!hasSupplierMap) {
      // No Master File loaded - all suppliers are "new"
      const key = cert.supplierName.toLowerCase().trim();
      if (!seenNewSuppliers.has(key) && cert.supplierName) {
        seenNewSuppliers.add(key);
        newSuppliers.push({
          supplierName: cert.supplierName,
          fileName: cert.fileName,
          certification: cert.certification || 'Unknown',
        });
      }
      return cert;
    }

    // Match against Master File
    const result = matchSupplier(cert.supplierName, supplierMap, 0.75);

    if (result.wasMatched) {
      matchedCount++;
      // Update supplier name to official Master File name
      return {
        ...cert,
        supplierName: result.matchedName,
      };
    } else {
      // NEW SUPPLIER - not in Master File
      const key = cert.supplierName.toLowerCase().trim();
      if (!seenNewSuppliers.has(key) && cert.supplierName) {
        seenNewSuppliers.add(key);
        newSuppliers.push({
          supplierName: cert.supplierName,
          fileName: cert.fileName,
          certification: cert.certification || 'Unknown',
        });
      }
      return cert;
    }
  });

  console.log(`Supplier matching: ${matchedCount} matched, ${newSuppliers.length} new suppliers detected`);

  return {
    processedCertificates,
    newSuppliers,
    matchedCount,
    duplicatesRemoved,
  };
}

export const exportToExcel = async (
  certificates: CertificateData[],
  supplierMap: DynamicSupplierMap = {}
) => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Certificate Analyzer';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Certificates');

  // ==========================================================================
  // STEP 1: Apply "Saurebh Filter" Deduplication + Supplier Matching
  // ==========================================================================
  const { processedCertificates: deduplicated } = prepareExportData(certificates, supplierMap);

  // ==========================================================================
  // STEP 2: Sort Alphabetically by Supplier Name
  // ==========================================================================
  const sortedCertificates = [...deduplicated].sort((a, b) => {
    const nameA = (a.supplierName || '').toLowerCase();
    const nameB = (b.supplierName || '').toLowerCase();
    return nameA.localeCompare(nameB);
  });

  // ==========================================================================
  // STEP 3: Define V7 Master Columns (Strict Order)
  // ==========================================================================
  worksheet.columns = [
    { header: 'Supplier Account', key: 'supplierAccount', width: 18 },        // Col 1
    { header: 'Supplier Name', key: 'supplierName', width: 30 },              // Col 2
    { header: 'Country', key: 'country', width: 15 },                         // Col 3
    { header: 'Scope', key: 'scope', width: 10 },                             // Col 4
    { header: 'Measure', key: 'measure', width: 40 },                         // Col 5
    { header: 'Certification', key: 'certification', width: 25 },             // Col 6
    { header: 'Product Category', key: 'productCategory', width: 18 },        // Col 7
    { header: 'Status', key: 'status', width: 12 },                           // Col 8
    { header: 'Issued', key: 'issued', width: 12 },                           // Col 9
    { header: 'Date of Expiry', key: 'dateOfExpiry', width: 14 },             // Col 10
    { header: 'Days to Expire', key: 'daysToExpire', width: 14 },             // Col 11
    { header: 'Contact person & Email', key: 'contactEmail', width: 28 },     // Col 12
    { header: 'Date Request sent', key: 'dateRequestSent', width: 18 },       // Col 13
    { header: 'Date Received', key: 'dateReceived', width: 14 },              // Col 14
    { header: 'Comments', key: 'comments', width: 25 },                       // Col 15
  ];

  // ==========================================================================
  // STEP 4: Style the Header Row
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

  // ==========================================================================
  // STEP 5: Add Data Rows
  // ==========================================================================
  sortedCertificates.forEach((cert) => {
    // Get issue and expiry dates
    const issueDate = cert.issueDate || '';
    const rawExpiryDate = cert.expiryDate || '';

    // Apply 3-year rule to get effective expiry date
    const effectiveExpiryDate = applyThreeYearRule(issueDate, rawExpiryDate);

    // Calculate days to expiry and status
    const daysToExpiry = calculateDaysToExpiry(effectiveExpiryDate);
    const status = getStatus(daysToExpiry);

    worksheet.addRow({
      supplierAccount: '',                                                    // Col 1: Empty
      supplierName: cert.supplierName || '',                                  // Col 2
      country: cert.country || '',                                            // Col 3
      scope: cert.scope || '',                                                // Col 4: "!" or "+"
      measure: cert.measure || cert.ecRegulation || '',                       // Col 5
      certification: cert.certification || '',                                // Col 6
      productCategory: cert.productCategory || cert.product || '',            // Col 7
      status: status,                                                         // Col 8
      issued: issueDate,                                                      // Col 9
      dateOfExpiry: effectiveExpiryDate,                                      // Col 10
      daysToExpire: daysToExpiry !== null ? daysToExpiry : '',                // Col 11
      contactEmail: '',                                                       // Col 12: Empty
      dateRequestSent: '',                                                    // Col 13: Empty
      dateReceived: '',                                                       // Col 14: Empty
      comments: '',                                                           // Col 15: Empty
    });
  });

  // ==========================================================================
  // STEP 6: Apply Conditional Formatting
  // ==========================================================================
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

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
      // Red for expired
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
      // Green for valid
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
      // Orange/Yellow for unknown (needs attention)
      statusCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFEB9C' },
      };
      statusCell.font = { bold: true, color: { argb: 'FF9C5700' } };
    }
  });

  // Alternating row colors for readability
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (rowNumber % 2 === 0) {
      row.eachCell((cell, colNumber) => {
        // Don't override status (col 8) and days (col 11)
        if (colNumber !== 8 && colNumber !== 11) {
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
  // STEP 7: Generate and Download
  // ==========================================================================
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });

  const date = new Date().toISOString().split('T')[0];
  saveAs(blob, `certificates-${date}.xlsx`);
};
