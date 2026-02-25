import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';
import { DynamicSupplierMap } from '@/utils/masterFileParser';
import { prepareExportData } from '@/utils/exportExcel';

// =============================================================================
// APPEND TO MASTER EXCEL — V10 COMPLIANT
// Follows "Certificate Master Update Rules & Workflow" (2026-02-04):
//
//  § Column rules
//    A  Supplier Account    — write
//    B  Supplier Name       — write
//    C  Country             — write
//    D  Scope               — NEVER WRITE (shapes / icons live here)
//    E  Measure             — write
//    F  Certification       — write
//    G  Product Category    — write
//    H  Status              — FORMULA-DRIVEN: copy formula from row above
//    I  Issued              — write as Date object
//    J  Date of Expiry      — write as Date object, or "No Date" if unknown
//    K  Days to Expire      — FORMULA-DRIVEN: copy formula from row above
//    O  Comments            — write filename + cert number
//
//  § Expiry rule (§6): If expiry not stated → write exact string "No Date".
//    Never null/empty. Never invent an expiry date.
//
//  § Update Log (§12): append one summary row to the "Update Log" sheet.
// =============================================================================

// -----------------------------------------------------------------------------
// STEP 0 — Strip drawing XML from the XLSX ZIP so ExcelJS can load cleanly.
//
// Root causes of two sequential crashes:
//   1. "reading 'Anchors'" — ExcelJS drawing parser hits a null drawing anchor
//      in xl/drawings/drawingN.xml.
//   2. "reading 'Target'"  — ExcelJS sees <drawing r:id="rId1"/> in the
//      worksheet XML, looks up rId1 in the rels, gets undefined because we
//      removed the rels entry but NOT the <drawing> reference in the sheet XML.
//
// Full 4-layer fix:
//   Layer 1 — remove xl/drawings/* files (stops 'Anchors' crash)
//   Layer 2 — patch worksheet XMLs to remove <drawing>/<legacyDrawing> refs
//              (stops 'Target' crash — the reference is gone before lookup)
//   Layer 3 — patch all .rels files to remove drawing Relationship entries
//   Layer 4 — patch [Content_Types].xml to remove drawing Override entries
// -----------------------------------------------------------------------------
async function stripDrawings(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(buffer);
  const allPaths = Object.keys(zip.files);

  // Layer 1: Remove drawing XML files
  const drawingFiles = allPaths.filter(p => {
    if (p.endsWith('/')) return false;
    const lower = p.toLowerCase();
    return lower.includes('/drawings/') || lower.includes('vmldrawing');
  });

  if (drawingFiles.length === 0) {
    console.log('[stripDrawings] No drawing files — loading as-is');
    return buffer;
  }
  console.log('[stripDrawings] Removing:', drawingFiles);
  for (const p of drawingFiles) zip.remove(p);

  // Layer 2: Remove <drawing r:id="..."/> and <legacyDrawing r:id="..."/>
  // elements from worksheet XML files. Without this, ExcelJS looks up the
  // r:id in the (now-patched) rels, gets undefined, and crashes on .Target.
  const wsXmlPaths = allPaths.filter(
    p => p.startsWith('xl/worksheets/') && p.endsWith('.xml') && !p.includes('_rels')
  );
  for (const wsPath of wsXmlPaths) {
    const file = zip.file(wsPath);
    if (!file) continue;
    const original = await file.async('string');
    const patched = original
      .replace(/<drawing\b[^]*?\/>/g, '')
      .replace(/<legacyDrawing\b[^]*?\/>/g, '');
    if (patched !== original) {
      zip.file(wsPath, patched);
      console.log('[stripDrawings] Removed drawing refs from:', wsPath);
    }
  }

  // Layer 3: Patch all .rels files — remove Relationship entries whose Type
  // contains "drawing". Uses [^]*? (matches any char including newlines,
  // non-greedy) in a callback so attribute order doesn't matter.
  const stripDrawingRels = (xml: string): string =>
    xml.replace(/<Relationship\b[^]*?\/>/g, (match) => {
      if (/Type="[^"]*drawing/i.test(match)) {
        console.log('[stripDrawings] Removed rel:', match.slice(0, 80));
        return '';
      }
      return match;
    });

  for (const relPath of allPaths.filter(p => p.endsWith('.rels'))) {
    const file = zip.file(relPath);
    if (!file) continue;
    const original = await file.async('string');
    const patched = stripDrawingRels(original);
    if (patched !== original) {
      zip.file(relPath, patched);
      console.log('[stripDrawings] Patched rels:', relPath);
    }
  }

  // Layer 4: Remove drawing Override entries from [Content_Types].xml
  const ctFile = zip.file('[Content_Types].xml');
  if (ctFile) {
    const original = await ctFile.async('string');
    const patched = original.replace(/<Override\b[^]*?\/>/g, (match) => {
      if (/drawing/i.test(match)) return '';
      return match;
    });
    if (patched !== original) zip.file('[Content_Types].xml', patched);
  }

  const result = await zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  console.log(`[stripDrawings] Done — stripped ${drawingFiles.length} drawing file(s)`);
  return result;
}

// -----------------------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------------------

/**
 * Shift all cell references in a formula string down by `shift` rows.
 * e.g. shiftFormulaRow('IF(J45="No Date","",K45)', 1) → 'IF(J46="No Date","",K46)'
 * Used to clone the Status (H) and Days to Expire (K) formulas into new rows.
 */
function shiftFormulaRow(formula: string, shift: number): string {
  return formula.replace(/([A-Z]+)(\d+)/g, (_m, col, row) =>
    `${col}${parseInt(row, 10) + shift}`
  );
}

/**
 * Deep-clone cell style + data validation without copying ExcelJS internal
 * proxy IDs. `{...cell.style}` copies proxy references and corrupts XML.
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

function normalizeHeader(value: unknown): string {
  return String(value ?? '').toLowerCase().trim().replace(/\s+/g, ' ');
}

/**
 * Find the data worksheet. Tries exact name "Certs 2025" first (§2 of rulebook),
 * then scans all sheets for "Supplier Name" header.
 */
function findDataSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet | undefined {
  // Exact name match per rulebook
  const exact = workbook.getWorksheet('Certs 2025');
  if (exact) return exact;

  // Fallback: header scan
  for (const ws of workbook.worksheets) {
    for (let r = 1; r <= Math.min(10, ws.rowCount); r++) {
      let found = false;
      ws.getRow(r).eachCell({ includeEmpty: false }, (cell) => {
        const val = normalizeHeader(cell.value);
        if (val.includes('supplier name') || val.includes('supplier account')) found = true;
      });
      if (found) return ws;
    }
  }
  return (
    workbook.worksheets.find(ws => ws.name.toLowerCase() !== 'instructions') ??
    workbook.worksheets[0]
  );
}

/**
 * Scan the first 10 rows of a sheet for the header row.
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
  // V10 defaults — header row 3, 1-based column numbers (per §2 rulebook)
  return {
    headerRowNum: 3,
    columnMap: {
      'supplier account':       1,
      'supplier name':          2,
      'country':                3,
      // col 4 = Scope — intentionally NOT in map (§2: never write)
      'measure':                5,
      'certification':          6,
      'product category':       7,
      'status':                 8,   // formula-driven — DO NOT WRITE
      'issued':                 9,
      'date of expiry':         10,
      'days to expire':         11,  // formula-driven — DO NOT WRITE
      'contact person & email': 12,
      'date request sent':      13,
      'date received':          14,
      'comments':               15,
    },
  };
}

/**
 * Parse a date string to a Date object for ExcelJS (writes as proper Excel serial).
 * Returns null for empty / "Not Found" / "No Date".
 * Returns the original string if it can't be parsed as a date.
 */
function parseDateValue(dateStr: string): Date | string | null {
  if (!dateStr || dateStr === 'Not Found' || dateStr === 'No Date') return null;
  const d = new Date(dateStr);
  if (!isNaN(d.getTime())) return d;
  return dateStr;
}

// -----------------------------------------------------------------------------
// Main export
// -----------------------------------------------------------------------------
export async function appendToMasterExcel(
  originalBuffer: ArrayBuffer,
  certificates: CertificateData[],
  supplierMap?: DynamicSupplierMap
): Promise<void> {
  if (certificates.length === 0) {
    alert('No certificates to append');
    return;
  }

  // Step 1: Saurebh filter + supplier matching
  const { processedCertificates } = prepareExportData(certificates, supplierMap);
  if (processedCertificates.length === 0) {
    alert('No certificates remaining after deduplication');
    return;
  }

  // Step 2: Strip drawing XML (4-layer fix — see stripDrawings above)
  const cleanBuffer = await stripDrawings(originalBuffer);

  // Step 3: Load with ExcelJS
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(cleanBuffer);

  // Step 4: Locate the data sheet
  const ws = findDataSheet(workbook);
  if (!ws) throw new Error('No data sheet found in Master File');

  // Step 5: Build column map
  const { headerRowNum, columnMap } = findHeaderInfo(ws);

  const colIdx = (aliases: string[]): number | undefined => {
    for (const a of aliases) {
      if (columnMap[a] !== undefined) return columnMap[a];
    }
    return undefined;
  };

  // Intentionally NOT mapping: scope (D), status (H), daysToExpire (K)
  const cols = {
    supplierAccount: colIdx(['supplier account', 'account']),
    supplierName:    colIdx(['supplier name', 'supplier']),
    country:         colIdx(['country']),
    measure:         colIdx(['measure']),
    certification:   colIdx(['certification', 'certification ']),
    productCategory: colIdx(['product category', 'product category ']),
    issued:          colIdx(['issued', 'date issued', 'issue date']),
    dateOfExpiry:    colIdx(['date of expiry', 'expiry date', 'expiry']),
    comments:        colIdx(['comments', 'comment']),
  };

  // Step 6: Append one row per certificate
  const lastRowNum = ws.lastRow?.number ?? headerRowNum;
  const templateRow = ws.getRow(lastRowNum);

  for (let i = 0; i < processedCertificates.length; i++) {
    const cert = processedCertificates[i];
    const targetRowNum = lastRowNum + 1 + i;
    const rowShift    = targetRowNum - lastRowNum;
    const newRow      = ws.getRow(targetRowNum);

    // PASS 1 — Copy styles from template row (deep-safe clone)
    templateRow.eachCell({ includeEmpty: true }, (templateCell, colNumber) => {
      safeCopyStyleAndValidation(templateCell, newRow.getCell(colNumber));
    });

    // PASS 2 — Copy formulas from template row, shifting row numbers by rowShift.
    //   This preserves col H (Status) and col K (Days to Expire) as live formulas.
    //   Example: IF(J500="No Date","",…) → IF(J501="No Date","",…) for row 501.
    templateRow.eachCell({ includeEmpty: true }, (templateCell, colNumber) => {
      if (templateCell.type === ExcelJS.ValueType.Formula && templateCell.formula) {
        const shifted = shiftFormulaRow(templateCell.formula, rowShift);
        newRow.getCell(colNumber).value = { formula: shifted, result: undefined };
      }
    });

    // PASS 3 — Write data values (overrides template data values for data columns;
    //   H and K keep their formulas from PASS 2 because we never call setCell on them)
    const setCell = (colIndex: number | undefined, value: ExcelJS.CellValue | null) => {
      if (colIndex === undefined) return;
      newRow.getCell(colIndex).value = value ?? null;
    };

    setCell(cols.supplierAccount, (cert as any)._matchedAccount || null);
    setCell(cols.supplierName,    cert.supplierName    || null);
    setCell(cols.country,         cert.country         || null);
    // col D (Scope): NEVER WRITE — rulebook §2 & §5
    setCell(cols.measure,         cert.measure         || null);
    setCell(cols.certification,   cert.certification   || null);
    setCell(cols.productCategory, cert.productCategory || null);
    // col H (Status): NEVER WRITE — formula from PASS 2
    setCell(cols.issued, parseDateValue(cert.issueDate || ''));

    // Expiry rule (§6): "No Date" if not stated — never null/empty
    const expiryStr = cert.expiryDate && cert.expiryDate !== 'Not Found'
      ? cert.expiryDate : '';
    const expiryValue: ExcelJS.CellValue =
      expiryStr ? (parseDateValue(expiryStr) ?? 'No Date') : 'No Date';
    setCell(cols.dateOfExpiry, expiryValue);

    // col K (Days to Expire): NEVER WRITE — formula from PASS 2

    // Comments (O): filename + cert number if available
    const commentParts: string[] = [];
    if (cert.fileName)          commentParts.push(cert.fileName);
    if (cert.certificateNumber) commentParts.push(`Cert #${cert.certificateNumber}`);
    setCell(cols.comments, commentParts.join(' | ') || null);

    newRow.commit();
  }

  console.log(
    `[appendToMasterExcel] Appended ${processedCertificates.length} row(s) to "${ws.name}"`
  );

  // Step 7: Update Log (§12)
  const logSheet = workbook.getWorksheet('Update Log');
  if (logSheet) {
    const now = new Date();
    const timestamp = `${now.toISOString().slice(0, 10)} ${now.toTimeString().slice(0, 5)}`;
    const first = processedCertificates[0];
    const lastLog = logSheet.lastRow?.number ?? 0;
    const logRow  = logSheet.getRow(lastLog + 1);
    logRow.getCell(1).value = timestamp;
    logRow.getCell(2).value = (first as any)._matchedAccount || '';
    logRow.getCell(3).value = first.supplierName || '';
    logRow.getCell(4).value = `Inserted ${processedCertificates.length} certificate(s)`;
    logRow.getCell(5).value = 'Appended via Vexos Engine. Expiry rules applied.';
    logRow.commit();
    console.log(`[appendToMasterExcel] Update Log row added at ${lastLog + 1}`);
  } else {
    console.warn('[appendToMasterExcel] "Update Log" sheet not found — skipping');
  }

  // Step 8: Write and download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer as ArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const today = new Date().toISOString().slice(0, 10);
  saveAs(blob, `Updated_Master_File_${today}.xlsx`);
}
