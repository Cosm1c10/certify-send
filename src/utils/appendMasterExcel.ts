import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';
import { DynamicSupplierMap } from '@/utils/masterFileParser';
import { prepareExportData } from '@/utils/exportExcel';

// =============================================================================
// APPEND TO MASTER EXCEL
// Reads the original Master File using ExcelJS, finds the data sheet,
// appends new certificate rows at the bottom with safe style inheritance,
// and downloads the modified workbook — zero repair dialogs.
// =============================================================================

/**
 * Strip drawing XML from an XLSX buffer before ExcelJS loads it.
 *
 * The client's Master File contains small icon shapes (!, +) anchored to cells
 * in the Scope column. ExcelJS crashes on these with:
 *   "Cannot read properties of undefined (reading 'Anchors')"
 *
 * An XLSX is a ZIP archive. We use JSZip to:
 *   1. Remove xl/drawings/*.xml  (the shape/icon definitions)
 *   2. Remove xl/drawings/_rels/*.rels  (their relationship files)
 *   3. Patch xl/worksheets/_rels/sheet*.xml.rels to remove drawing references
 *   4. Patch [Content_Types].xml to remove drawing content-type overrides
 *
 * This lets ExcelJS load cleanly. The icons disappear from the output,
 * but all data, named ranges, tables, and conditional formats are preserved.
 */
async function stripDrawings(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);

  // Collect drawing paths before iterating (avoid mutation during iteration)
  const drawingPaths = Object.keys(zip.files).filter(
    p => p.startsWith('xl/drawings/') && !p.endsWith('/')
  );

  if (drawingPaths.length === 0) {
    // No drawings — return original buffer unchanged
    return buffer;
  }

  // 1 & 2: Remove drawing XML and their _rels
  for (const path of drawingPaths) {
    zip.remove(path);
  }

  // 3: Patch worksheet _rels files — remove any <Relationship> that points to a drawing
  const wsRelPaths = Object.keys(zip.files).filter(
    p => p.startsWith('xl/worksheets/_rels/') && p.endsWith('.rels')
  );
  for (const relPath of wsRelPaths) {
    let xml = await zip.file(relPath)!.async('string');
    // Remove complete self-closing Relationship elements whose Type ends in /drawing.
    // Previous regex stopped at the first "/" inside the element and left a dangling
    // "drawings/drawing1.xml"/>" fragment — causing the "reading 'Target'" crash.
    xml = xml.replace(/<Relationship[^>]*Type="[^"]*\/drawing"[^>]*\/>/g, '');
    zip.file(relPath, xml);
  }

  // 4: Patch [Content_Types].xml — remove Override entries for drawing files
  const ctFile = zip.file('[Content_Types].xml');
  if (ctFile) {
    let ctXml = await ctFile.async('string');
    ctXml = ctXml.replace(/<Override[^>]+drawing[^>]*\/>/g, '');
    zip.file('[Content_Types].xml', ctXml);
  }

  const result = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  console.log(`[stripDrawings] Removed ${drawingPaths.length} drawing file(s) from XLSX`);
  return result;
}

/**
 * Deep-clone cell style and data validation without copying ExcelJS internal
 * proxy IDs or XML references. Using `{ ...cell.style }` copies those refs
 * and produces a corrupted workbook on write.
 */
function safeCopyStyleAndValidation(
  sourceCell: ExcelJS.Cell,
  targetCell: ExcelJS.Cell
): void {
  if (sourceCell.style) {
    targetCell.style = {
      font:      sourceCell.style.font      ? JSON.parse(JSON.stringify(sourceCell.style.font))      : undefined,
      border:    sourceCell.style.border    ? JSON.parse(JSON.stringify(sourceCell.style.border))    : undefined,
      fill:      sourceCell.style.fill      ? JSON.parse(JSON.stringify(sourceCell.style.fill))      : undefined,
      alignment: sourceCell.style.alignment ? JSON.parse(JSON.stringify(sourceCell.style.alignment)) : undefined,
      numFmt:    sourceCell.style.numFmt,
    };
  }
  if (sourceCell.dataValidation) {
    targetCell.dataValidation = JSON.parse(JSON.stringify(sourceCell.dataValidation));
  }
}

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
 * Parse a date string safely.
 * Returns a Date object if valid (ExcelJS writes it as a proper Excel date serial).
 * Returns the original string if it can't be parsed (e.g. "No Date").
 * Returns null for empty / "Not Found".
 */
function parseDateValue(dateStr: string): Date | string | null {
  if (!dateStr || dateStr === 'Not Found' || dateStr === 'No Date') return null;
  const d = new Date(dateStr);
  if (!isNaN(d.getTime())) return d;
  return dateStr;
}

/**
 * Find the worksheet that contains supplier data.
 * Scans all sheets for "Supplier Name" or "Supplier Account" in the first 10 rows.
 */
function findDataSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | undefined {
  for (const ws of workbook.worksheets) {
    for (let r = 1; r <= Math.min(10, ws.rowCount); r++) {
      let found = false;
      ws.getRow(r).eachCell({ includeEmpty: false }, (cell) => {
        const val = normalizeHeader(cell.value);
        if (val.includes('supplier name') || val.includes('supplier account')) {
          found = true;
        }
      });
      if (found) return ws;
    }
  }
  // Fallback: first sheet that isn't "Instructions"
  return (
    workbook.worksheets.find(ws => ws.name.toLowerCase() !== 'instructions') ??
    workbook.worksheets[0]
  );
}

/**
 * Scan the first 10 rows of a sheet to find the header row.
 * Returns headerRowNum (1-based) and columnMap (normalizedHeader → 1-based col number).
 */
function findHeaderInfo(ws: ExcelJS.Worksheet): {
  headerRowNum: number;
  columnMap: Record<string, number>;
} {
  for (let r = 1; r <= Math.min(10, ws.rowCount); r++) {
    const columnMap: Record<string, number> = {};
    let hasSupplierName = false;
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const normalized = normalizeHeader(cell.value);
      columnMap[normalized] = colNumber;
      if (normalized.includes('supplier name')) hasSupplierName = true;
    });
    if (hasSupplierName) return { headerRowNum: r, columnMap };
  }
  // V7 defaults (1-based column numbers)
  return {
    headerRowNum: 3,
    columnMap: {
      'supplier account': 1,
      'supplier name':    2,
      'country':          3,
      'scope':            4,
      'measure':          5,
      'certification':    6,
      'product category': 7,
      'status':           8,
      'issued':           9,
      'date of expiry':   10,
      'days to expire':   11,
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

  // Step 2: Strip drawing XML (anchored icons) so ExcelJS doesn't crash,
  //         then load the cleaned buffer
  const cleanBuffer = await stripDrawings(originalBuffer);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(cleanBuffer);

  // Step 3: Find the data sheet
  const ws = findDataSheet(workbook);
  if (!ws) {
    throw new Error('No data sheet found in Master File');
  }

  // Step 4: Find header row + build column map
  const { headerRowNum, columnMap } = findHeaderInfo(ws);

  // Helper: resolve column aliases → 1-based column number
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

  // Step 5: Find the last populated row — its style is used as the template
  const lastRowNum = ws.lastRow?.number ?? headerRowNum;
  const templateRow = ws.getRow(lastRowNum);

  // Step 6: Append one row per certificate
  for (let i = 0; i < processedCertificates.length; i++) {
    const cert = processedCertificates[i];
    const expiryDate = resolveExpiryDate(cert);
    const days = calcDaysToExpiry(expiryDate);
    const status = getStatus(days);

    const targetRowNum = lastRowNum + 1 + i;
    const newRow = ws.getRow(targetRowNum);

    // ACTION 1: Deep-safe style + validation copy (no proxy ID bleed)
    templateRow.eachCell({ includeEmpty: true }, (templateCell, colNumber) => {
      safeCopyStyleAndValidation(templateCell, newRow.getCell(colNumber));
    });

    // ACTION 2: Set values cell-by-cell (never addRow, never undefined)
    const setCell = (colIndex: number | undefined, value: ExcelJS.CellValue | null) => {
      if (colIndex === undefined) return;
      newRow.getCell(colIndex).value = value ?? null;
    };

    setCell(cols.supplierAccount, (cert as any)._matchedAccount || null);
    setCell(cols.supplierName,    cert.supplierName    || null);
    setCell(cols.country,         cert.country         || null);
    setCell(cols.scope,           cert.scope           || null);
    setCell(cols.measure,         cert.measure         || null);
    setCell(cols.certification,   cert.certification   || null);
    setCell(cols.productCategory, cert.productCategory || null);
    setCell(cols.status,          status);

    // ACTION 3: Strict date types — Date object if valid, string fallback, null if empty
    setCell(cols.issued,       parseDateValue(cert.issueDate || ''));
    setCell(cols.dateOfExpiry, parseDateValue(expiryDate));
    setCell(cols.daysToExpire, days !== null ? days : null);

    newRow.commit();
  }

  console.log(`Appended ${processedCertificates.length} rows to sheet "${ws.name}"`);

  // Step 7: Write and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer as ArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const today = new Date().toISOString().slice(0, 10);
  saveAs(blob, `Updated_Master_File_${today}.xlsx`);
}
