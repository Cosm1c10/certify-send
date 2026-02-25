import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';
import { DynamicSupplierMap } from '@/utils/masterFileParser';
import { prepareExportData } from '@/utils/exportExcel';

// =============================================================================
// APPEND TO MASTER EXCEL
// Reads the original Master File using SheetJS (robust, no 'anchors' bug),
// appends new certificate rows to the data sheet, and downloads the modified
// workbook. cellStyles is intentionally disabled to preserve named ranges.
// =============================================================================

/**
 * Normalize a header cell value for column map lookup.
 */
function normalizeHeader(value: unknown): string {
  return String(value ?? '').toLowerCase().trim().replace(/\s+/g, ' ');
}

/**
 * Apply the 3-year rule: if expiryDate is missing but issueDate exists,
 * expiry is assumed to be 3 years from the issue date.
 */
function resolveExpiryDate(cert: CertificateData): string {
  if (cert.expiryDate && cert.expiryDate !== 'Not Found') return cert.expiryDate;
  if (cert.issueDate && cert.issueDate !== 'Not Found') {
    const issued = new Date(cert.issueDate);
    if (!isNaN(issued.getTime())) {
      issued.setFullYear(issued.getFullYear() + 3);
      return issued.toISOString().slice(0, 10);
    }
  }
  return '';
}

function calcDaysToExpiry(expiryDate: string): number | null {
  if (!expiryDate) return null;
  const expiry = new Date(expiryDate);
  if (isNaN(expiry.getTime())) return null;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.floor((expiry.getTime() - today.getTime()) / 86_400_000);
}

function getStatus(days: number | null): string {
  if (days === null) return 'Unknown';
  return days < 0 ? 'Expired' : 'Up to date';
}

/**
 * Find the sheet name that contains supplier data by scanning for
 * "Supplier Name" or "Supplier Account" in the first 10 rows.
 */
function findDataSheetName(workbook: XLSX.WorkBook): string {
  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    if (!ws) continue;
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    for (let r = 0; r < Math.min(10, rows.length); r++) {
      for (const cell of (rows[r] || [])) {
        const val = normalizeHeader(cell);
        if (val.includes('supplier name') || val.includes('supplier account')) {
          return name;
        }
      }
    }
  }
  // Fallback: first sheet that isn't "Instructions"
  return workbook.SheetNames.find(n => n.toLowerCase() !== 'instructions')
    ?? workbook.SheetNames[0];
}

/**
 * Scan the first 10 rows of a sheet to find the header row.
 * Returns headerRowIndex (0-based) and columnMap (normalizedHeader → 0-based col index).
 */
function findHeaderInfo(rows: any[][]): {
  headerRowIndex: number;
  columnMap: Record<string, number>;
} {
  for (let r = 0; r < Math.min(10, rows.length); r++) {
    const row = rows[r] || [];
    const columnMap: Record<string, number> = {};
    let hasSupplierName = false;
    for (let c = 0; c < row.length; c++) {
      const normalized = normalizeHeader(row[c]);
      columnMap[normalized] = c;
      if (normalized.includes('supplier name')) hasSupplierName = true;
    }
    if (hasSupplierName) return { headerRowIndex: r, columnMap };
  }
  // V7 defaults (0-based)
  return {
    headerRowIndex: 2,
    columnMap: {
      'supplier account': 0,
      'supplier name': 1,
      'country': 2,
      'scope': 3,
      'measure': 4,
      'certification': 5,
      'product category': 6,
      'status': 7,
      'issued': 8,
      'date of expiry': 9,
      'days to expire': 10,
    },
  };
}

/**
 * Main function: append extracted certificates to the original Master File
 * and download the modified workbook.
 */
export async function appendToMasterExcel(
  originalBuffer: ArrayBuffer,
  certificates: CertificateData[],
  supplierMap?: DynamicSupplierMap
): Promise<void> {
  if (certificates.length === 0) {
    alert('No certificates to append');
    return;
  }

  // Step 1: Apply Saurebh Filter + supplier name matching
  const { processedCertificates } = prepareExportData(certificates, supplierMap);
  if (processedCertificates.length === 0) {
    alert('No certificates remaining after deduplication');
    return;
  }

  // Step 2: Load the workbook with SheetJS (handles complex Excel files without 'anchors' bug)
  // NOTE: cellStyles: true is intentionally omitted — it corrupts named ranges in the workbook
  const workbook = XLSX.read(originalBuffer, {
    type: 'array',
    bookVBA: false,
  });

  // Step 3: Find the data sheet
  const sheetName = findDataSheetName(workbook);
  const ws = workbook.Sheets[sheetName];
  if (!ws) {
    throw new Error(`Sheet "${sheetName}" not found in Master File`);
  }

  // Step 4: Find header row + column positions
  const allRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const { headerRowIndex, columnMap } = findHeaderInfo(allRows);

  // Helper: resolve column alias → 0-based index
  const colIdx = (aliases: string[]): number | undefined => {
    for (const alias of aliases) {
      if (columnMap[alias] !== undefined) return columnMap[alias];
    }
    return undefined;
  };

  const cols = {
    supplierAccount: colIdx(['supplier account', 'account']),
    supplierName:    colIdx(['supplier name', 'supplier']),
    country:         colIdx(['country']),
    scope:           colIdx(['scope']),
    measure:         colIdx(['measure']),
    certification:   colIdx(['certification', 'certification ']),
    productCategory: colIdx(['product category', 'product category ']),
    status:          colIdx(['status']),
    issued:          colIdx(['issued', 'date issued', 'issue date']),
    dateOfExpiry:    colIdx(['date of expiry', 'expiry date', 'expiry']),
    daysToExpire:    colIdx(['days to expire', 'days to expiry']),
  };

  // Step 5: Find the current data range
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

  // Step 6: Append one row per certificate
  for (let i = 0; i < processedCertificates.length; i++) {
    const cert = processedCertificates[i];
    const expiryDate = resolveExpiryDate(cert);
    const days = calcDaysToExpiry(expiryDate);
    const status = getStatus(days);

    // New row index (0-based), one after the current last row
    const newRowIdx = range.e.r + 1;

    // Set cell values by column position
    const setCell = (colIndex: number | undefined, value: unknown) => {
      if (colIndex === undefined || value === undefined || value === null || value === '') return;
      const addr = XLSX.utils.encode_cell({ r: newRowIdx, c: colIndex });
      if (!ws[addr]) ws[addr] = {};
      ws[addr].t = typeof value === 'number' ? 'n' : 's';
      ws[addr].v = value;
    };

    setCell(cols.supplierAccount, (cert as any)._matchedAccount ?? '');
    setCell(cols.supplierName,    cert.supplierName);
    setCell(cols.country,         cert.country);
    setCell(cols.scope,           cert.scope);
    setCell(cols.measure,         cert.measure);
    setCell(cols.certification,   cert.certification);
    setCell(cols.productCategory, cert.productCategory);
    setCell(cols.status,          status);
    setCell(cols.issued,          cert.issueDate || '');
    setCell(cols.dateOfExpiry,    expiryDate || '');
    setCell(cols.daysToExpire,    days !== null ? days : '');

    // Expand the sheet range to include the new row
    range.e.r = newRowIdx;
  }

  // Update the sheet's !ref to include new rows
  ws['!ref'] = XLSX.utils.encode_range(range);

  console.log(`Appended ${processedCertificates.length} rows to sheet "${sheetName}"`);

  // Step 7: Write and download
  // NOTE: cellStyles: true omitted to preserve named ranges + workbook metadata integrity
  const output = XLSX.write(workbook, {
    type: 'array',
    bookType: 'xlsx',
  });

  const blob = new Blob([output], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const today = new Date().toISOString().slice(0, 10);
  saveAs(blob, `Updated_Master_File_${today}.xlsx`);
}
