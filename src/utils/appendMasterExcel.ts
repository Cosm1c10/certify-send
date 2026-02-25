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

/**
 * Find the true last data row by scanning BACKWARDS from ws.rowCount.
 *
 * Why backwards instead of forwards:
 *   After expandSharedFormulas(), formula cells throughout the sheet (e.g. rows
 *   4–847 for Status/Days columns) have ExcelJS value objects even when their
 *   displayed result is an empty string. A forward scan checking `.value` would
 *   see these as non-empty and report row 847 (or wherever formulas end) as the
 *   last data row, pushing new supplier blocks to row 10,000+.
 *
 *   Scanning backwards with `cell.text` (the rendered string, not the raw
 *   value/formula object) correctly treats empty-result formula cells as blank
 *   and stops at the first row that actually shows visible text in either the
 *   Supplier Name column or the Certification column.
 */
function findTrueLastRow(
  ws: ExcelJS.Worksheet,
  supplierNameCol: number,
  certCol: number,
  startRow: number
): number {
  const bottom = Math.max(ws.rowCount, startRow);

  for (let r = bottom; r >= startRow; r--) {
    const row        = ws.getRow(r);
    const nameText   = (row.getCell(supplierNameCol).text ?? '').trim();
    const certText   = (row.getCell(certCol).text         ?? '').trim();
    if (nameText !== '' || certText !== '') {
      console.log(`[findTrueLastRow] True last data row: ${r}`);
      return r;
    }
  }

  // No data found at all — return the row just above the first data row
  console.log(`[findTrueLastRow] No data found; returning startRow - 1 = ${startRow - 1}`);
  return startRow - 1;
}

/**
 * Find the row number at which to insert the new certificate row.
 *
 * Strategy (§8 rulebook):
 *   1. Scan col A from headerRow+1 to trueLastRow for rows matching overrideAccount.
 *   2. If found → insert at lastSupplierRow + 1 (within the supplier's block).
 *   3. If not found → insert at trueLastRow + 1 (after all existing data).
 *
 * Re-scanning on every cert ensures previous insertions (which shift row numbers) are
 * naturally accounted for without needing manual offset arithmetic.
 */
function findInsertionRow(
  ws: ExcelJS.Worksheet,
  supplierAccountCol: number,
  supplierNameCol: number,
  certCol: number,
  headerRowNum: number,
  overrideAccount: string | undefined
): { insertAt: number; supplierFound: boolean } {
  const trueLastRow = findTrueLastRow(ws, supplierNameCol, certCol, headerRowNum + 1);

  if (overrideAccount) {
    const normalizedTarget = overrideAccount.toLowerCase().trim();
    let lastSupplierRow = -1;

    for (let r = headerRowNum + 1; r <= trueLastRow; r++) {
      const cellVal = ws.getRow(r).getCell(supplierAccountCol).value;
      if (cellVal !== null && cellVal !== undefined) {
        if (String(cellVal).toLowerCase().trim() === normalizedTarget) {
          lastSupplierRow = r;
        }
      }
    }

    if (lastSupplierRow >= 0) {
      console.log(
        `[findInsertionRow] Supplier "${overrideAccount}" last row: ${lastSupplierRow} → inserting at ${lastSupplierRow + 1}`
      );
      return { insertAt: lastSupplierRow + 1, supplierFound: true };
    }
  }

  // +2 instead of +1: leaves one blank separator row between the last
  // existing supplier block and the new one, matching the sheet's visual style.
  console.log(
    `[findInsertionRow] Supplier "${overrideAccount ?? '(none)'}" not found → inserting at trueLastRow+2 = ${trueLastRow + 2}`
  );
  return { insertAt: trueLastRow + 2, supplierFound: false };
}

// -----------------------------------------------------------------------------
// Deduplication helper
// -----------------------------------------------------------------------------

/**
 * Scan a supplier's row block for an existing certificate that matches the
 * incoming cert. Returns the matching row number, or null if no match.
 *
 * Match criteria — a row is considered the SAME cert if EITHER is true:
 *
 *   A. FILENAME match in Comments (col 15).
 *      Tried first with the exact uploaded name, then again with OS copy-markers
 *      stripped ("cert (1).pdf" → "cert.pdf"). Handles both re-uploads and
 *      OS-renamed duplicates without requiring any field values to match.
 *
 *   B. CERT NAME containment (bidirectional, min 5 chars) AND PRODUCT match.
 *      No measure gate — the client's sheet often stores the cert name with a
 *      long descriptor (e.g. "ISO 9001:2015 (Quality Management System)") while
 *      the AI extracts just "ISO 9001:2015". Requiring measure equality blocked
 *      these obvious matches. The product gate still prevents "OK Compost Cups"
 *      from colliding with "OK Compost Containers".
 *
 * Column numbers are hard-coded (6/7/15) to prevent data-shift bugs.
 */
function findExistingCertRow(
  ws: ExcelJS.Worksheet,
  supplierAccountCol: number,
  headerRowNum: number,
  lastSupplierRow: number,
  overrideAccount: string | undefined,
  incomingCert: string,
  incomingMeasure: string,   // kept for future use / logging
  incomingFileName: string,
  incomingProduct: string
): number | null {
  if (!overrideAccount || lastSupplierRow < headerRowNum + 1) return null;

  const normAccount   = overrideAccount.toLowerCase().trim();
  const normInCert    = incomingCert.toLowerCase().trim();
  const normInFile    = incomingFileName.toLowerCase().trim();
  const normInProduct = incomingProduct.toLowerCase().trim();

  // Strip OS copy-markers before the extension: "cert (1).pdf" → "cert.pdf"
  const normInFileCleaned = normInFile.replace(/\s*\(\d+\)(\.\w+)$/, '$1');

  // Find the first row of this supplier's block in col A
  let blockStart = -1;
  for (let r = headerRowNum + 1; r <= lastSupplierRow; r++) {
    const v = ws.getRow(r).getCell(supplierAccountCol).value;
    if (v !== null && v !== undefined && String(v).toLowerCase().trim() === normAccount) {
      blockStart = r;
      break;
    }
  }
  if (blockStart < 0) return null;

  // Scan blockStart → lastSupplierRow (includes blank-A rows from previous inserts)
  for (let r = blockStart; r <= lastSupplierRow; r++) {
    const row = ws.getRow(r);

    // Hard-coded column indices: col 6 = Certification, col 7 = Product Category,
    // col 15 = Comments. (Measure / col 5 is no longer used as a match gate.)
    const existingCert    = (row.getCell(6).value?.toString()  ?? '').toLowerCase().trim();
    const existingProduct = (row.getCell(7).value?.toString()  ?? '').toLowerCase().trim();
    const comments        = (row.getCell(15).value?.toString() ?? '').toLowerCase();

    // --- CRITERION A: Filename in Comments (col 15) ---
    // Try exact name first; if it was OS-renamed (e.g. " (1)" appended), try
    // the cleaned version so "cert (1).pdf" still matches "cert.pdf" in comments.
    if (normInFile && comments.includes(normInFile)) {
      console.log(
        `[findExistingCertRow] Filename match at row ${r}: "${incomingFileName}" in comments`
      );
      return r;
    }
    if (normInFileCleaned !== normInFile && normInFileCleaned && comments.includes(normInFileCleaned)) {
      console.log(
        `[findExistingCertRow] Cleaned-filename match at row ${r}: ` +
        `"${normInFileCleaned}" (from "${incomingFileName}") in comments`
      );
      return r;
    }

    // --- CRITERION B: Cert name containment + Product gate ---
    // Both the AI name and the sheet name must be at least 5 chars so a
    // short token ("ISO") never accidentally matches everything.
    if (normInCert.length >= 5 && existingCert.length >= 5) {
      const isCertMatch = existingCert.includes(normInCert)
        || normInCert.includes(existingCert);

      // Empty product on either side → wildcard (legacy rows without a product
      // column still get updated rather than duplicated).
      const isProductMatch = existingProduct === normInProduct
        || existingProduct === ''
        || normInProduct === ''
        || existingProduct.includes(normInProduct)
        || normInProduct.includes(existingProduct);

      if (isCertMatch && isProductMatch) {
        console.log(
          `[findExistingCertRow] Cert-name match at row ${r}: ` +
          `existing="${existingCert}", incoming="${normInCert}", product="${existingProduct}"`
        );
        return r;
      }
    }
  }

  return null;
}

// -----------------------------------------------------------------------------
// Shared formula expansion
// -----------------------------------------------------------------------------

/**
 * Expand all shared formulas in a worksheet XML string to standalone formulas.
 *
 * Excel's XLSX format uses "shared formulas": one master cell stores the formula
 * text; every other cell in the range has an empty <f t="shared" si="N"/> clone
 * that references the master by index. ExcelJS does NOT update shared-formula
 * ranges when ws.insertRow() is called, so clones in shifted rows lose their
 * master reference and writeBuffer() throws:
 *   "Shared Formula master must exist above and or left of clone for cell KN"
 *
 * Fix: before loading with ExcelJS, convert every master+clone to an independent
 * formula with the correct row-shifted formula string, eliminating shared-formula
 * mechanics entirely.
 */
function expandSharedFormulasInXml(xml: string): string {
  // ── Phase 1: collect masters ──────────────────────────────────────────────
  // Master element: <f t="shared" ref="K4:K847" si="2">IF(J4="No Date","",…)</f>
  // inside a <c r="K4" …> element.
  const masterMap = new Map<number, { baseRow: number; formula: string }>();

  const masterRe = /<f\b([^>]*)>([^<]+)<\/f>/g;
  let fm: RegExpExecArray | null;
  while ((fm = masterRe.exec(xml)) !== null) {
    const attrs = fm[1];
    if (!attrs.includes('t="shared"')) continue;
    if (!attrs.includes('ref="'))      continue; // master must have ref=
    const siM = attrs.match(/si="(\d+)"/);
    if (!siM) continue;

    const formula  = fm[2].trim();
    const before   = xml.slice(0, fm.index);
    const cOpenIdx = before.lastIndexOf('<c ');
    if (cOpenIdx < 0) continue;
    const cellSnip = xml.slice(cOpenIdx, cOpenIdx + 120);
    const rowM     = cellSnip.match(/r="[A-Z]+(\d+)"/);
    if (!rowM) continue;

    masterMap.set(parseInt(siM[1], 10), {
      baseRow: parseInt(rowM[1], 10),
      formula,
    });
  }

  if (masterMap.size === 0) return xml;

  let cloneCount = 0;

  // ── Phase 2: replace masters + clones in a single pass ───────────────────
  const result = xml.replace(
    /<f\b([^>]*)(?:>([^<]*)<\/f>|\/>)/g,
    (fullMatch: string, attrs: string, content: string | undefined, offset: number): string => {
      if (!attrs.includes('t="shared"')) return fullMatch;

      // Master → strip shared attributes, keep formula as standalone
      if (content !== undefined && attrs.includes('ref="')) {
        return `<f>${content.trim()}</f>`;
      }

      // Clone (self-closing) → expand to shifted formula
      if (content === undefined) {
        const siM = attrs.match(/si="(\d+)"/);
        if (!siM) return fullMatch;
        const master = masterMap.get(parseInt(siM[1], 10));
        if (!master) return fullMatch;

        const before   = xml.slice(0, offset);
        const cOpenIdx = before.lastIndexOf('<c ');
        if (cOpenIdx < 0) return fullMatch;
        const cellSnip = xml.slice(cOpenIdx, cOpenIdx + 120);
        const rowM     = cellSnip.match(/r="[A-Z]+(\d+)"/);
        if (!rowM) return fullMatch;

        const cloneRow = parseInt(rowM[1], 10);
        const shift    = cloneRow - master.baseRow;
        cloneCount++;
        return `<f>${shiftFormulaRow(master.formula, shift)}</f>`;
      }

      return fullMatch;
    }
  );

  console.log(
    `[expandSharedFormulas] Expanded ${masterMap.size} master(s) + ${cloneCount} clone(s)`
  );
  return result;
}

/**
 * Pre-process XLSX buffer: expand all shared formulas to standalone formulas.
 * Prevents "Shared Formula master must exist above and or left of clone" errors
 * that ExcelJS throws during writeBuffer() when insertRow() has shifted rows.
 */
async function expandSharedFormulas(buffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip      = await JSZip.loadAsync(buffer);
  const allPaths = Object.keys(zip.files);

  const wsXmlPaths = allPaths.filter(
    p => p.startsWith('xl/worksheets/') && p.endsWith('.xml') && !p.includes('_rels')
  );

  let anyChanged = false;
  for (const wsPath of wsXmlPaths) {
    const file = zip.file(wsPath);
    if (!file) continue;
    const original = await file.async('string');
    const patched  = expandSharedFormulasInXml(original);
    if (patched !== original) {
      zip.file(wsPath, patched);
      anyChanged = true;
      console.log('[expandSharedFormulas] Patched:', wsPath);
    }
  }

  if (!anyChanged) {
    console.log('[expandSharedFormulas] No shared formulas — returning as-is');
    return buffer;
  }

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

// -----------------------------------------------------------------------------
// Smart alignment helper
// -----------------------------------------------------------------------------

/**
 * Apply column-aware alignment to every cell in a row (cols 1–15).
 *
 * Center-aligned columns: 4 (Scope), 8 (Status), 9 (Issued), 10 (Date of Expiry),
 *   11 (Days to Expire), 13 (Date Request Sent), 14 (Date Received).
 * All other columns: left-aligned (Supplier Name, Measure, Certification, etc.).
 *
 * Col 4 (Scope) additionally receives:
 *   - numFmt '@'       → forces Excel to store as text, preventing "+" / "!"
 *                        from being treated as formula operators.
 *   - font color red   → uniform ARGB 'FFFF0000' regardless of inherited style.
 */
function applySmartAlignment(row: ExcelJS.Row): void {
  const CENTER_COLS = new Set([4, 8, 9, 10, 11, 13, 14]);

  for (let col = 1; col <= 15; col++) {
    const cell = row.getCell(col);

    cell.alignment = {
      vertical:   'middle',
      horizontal: CENTER_COLS.has(col) ? 'center' : 'left',
      wrapText:   true,
    };

    if (col === 4) {
      // Force text format so "+" / "!" are never parsed as formula starters
      cell.numFmt = '@';
      // Uniform red font — override any inherited colour from the row above
      cell.font = {
        ...(cell.font ?? {}),
        color: { argb: 'FFFF0000' },
      };
    }
  }
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

  // Step 2: Pre-process XLSX ZIP before ExcelJS loads it:
  //   2a — Strip drawing XML (prevents 'Anchors' / 'Target' crashes)
  //   2b — Expand shared formulas to standalone formulas (prevents
  //         "Shared Formula master must exist above and or left of clone"
  //         errors that occur when insertRow() shifts shared-formula clones)
  const noDrawings  = await stripDrawings(originalBuffer);
  const cleanBuffer = await expandSharedFormulas(noDrawings);

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

  // Intentionally NOT written to as data: status (H), daysToExpire (K).
  // Both receive formula + pre-calculated result via PASS 2 (Protected View support).
  // scope (D) is written — ground truth mapping provides accurate ! / + values.
  // status / daysToExpire are mapped so their column numbers are known for PASS 2.
  const cols = {
    supplierAccount: colIdx(['supplier account', 'account']),
    supplierName:    colIdx(['supplier name', 'supplier']),
    country:         colIdx(['country']),
    scope:           colIdx(['scope'])           ?? 4,  // col D
    measure:         colIdx(['measure']),
    certification:   colIdx(['certification', 'certification ']),
    productCategory: colIdx(['product category', 'product category ']),
    status:          colIdx(['status'])          ?? 8,  // col H — formula + result only
    issued:          colIdx(['issued', 'date issued', 'issue date']),
    dateOfExpiry:    colIdx(['date of expiry', 'expiry date', 'expiry']),
    daysToExpire:    colIdx(['days to expire', 'days to expiry']) ?? 11, // col K — formula + result only
    comments:        colIdx(['comments', 'comment']),
  };

  // Step 6: For each certificate: update-in-place if a match exists in the
  // supplier's block, otherwise insert a new row at the bottom of that block.
  //
  // Update-in-place (ACTION 1): prevents duplicate rows when the same cert is
  // re-uploaded. Matches on certification keyword similarity OR filename in Comments.
  //
  // Protected View formula results (ACTION 2): pre-calculated Days and Status
  // values are injected as `result` alongside the formula so Excel displays them
  // even when the calculation engine is disabled (Protected View, strict mode).

  let insertCount = 0;
  let updateCount = 0;

  // In-memory tracking for this batch run (ACTIONs 2, 3, 4):
  //   processedCertTypesThisSession — blocks re-inserting the same cert type
  //     when the user uploads duplicate files in one batch.
  //   newBlockStartedFor — tracks which new-supplier keys already had their
  //     first (Account/Name/Country) row written, so subsequent certs for the
  //     same new supplier get blank A/B/C (visual grouping rule).
  const processedCertTypesThisSession = new Set<string>();
  const newBlockStartedFor = new Set<string>();

  for (const cert of processedCertificates) {
    // ACTION 4 — In-session duplicate prevention.
    // Key includes measure + cert + product category + filename so that
    // identical files re-uploaded in the same batch are skipped, but two
    // DIFFERENT products with the same cert type (e.g. OK Compost for Cups
    // and OK Compost for Containers) get their own keys and are both inserted.
    const certKey = (cert.measure || '') + '|' + (cert.certification || '')
      + '|' + (cert.productCategory || '') + '|' + (cert.fileName || '');
    if (processedCertTypesThisSession.has(certKey)) {
      console.log(
        `[appendToMasterExcel] Skipping in-session duplicate: "${certKey}" (${cert.fileName})`
      );
      continue;
    }
    processedCertTypesThisSession.add(certKey);

    const overrideAccount: string | undefined = (cert as any)._matchedAccount || undefined;

    // Re-scan each time — previous insertions shift row numbers
    const { insertAt, supplierFound } = findInsertionRow(
      ws,
      cols.supplierAccount ?? 1,
      cols.supplierName    ?? 2,
      cols.certification   ?? 6,
      headerRowNum,
      overrideAccount
    );

    // ── Pre-calculate Status & Days for Protected View ────────────────────────
    // These are injected as formula `result` so Excel shows them in Protected
    // View (where the calc engine is blocked) and in fullCalcOnLoad mode.
    const expiryForCalc = cert.expiryDate && cert.expiryDate !== 'Not Found'
      ? cert.expiryDate : '';
    let calculatedDays: number | string = '';
    let calculatedStatus: string = '';
    if (expiryForCalc) {
      const expiryMs = new Date(expiryForCalc).getTime();
      if (!isNaN(expiryMs)) {
        const days = Math.floor((expiryMs - Date.now()) / (1000 * 60 * 60 * 24));
        calculatedDays  = days;
        calculatedStatus = days < 0 ? 'Expired' : 'Up to date';
      }
    }

    // ── ACTION 1: Update-in-Place (Deduplication) ─────────────────────────────
    // Only check for duplicates when we know the supplier already has a block.
    if (supplierFound) {
      const matchRow = findExistingCertRow(
        ws,
        cols.supplierAccount ?? 1,
        headerRowNum,
        insertAt - 1,         // lastSupplierRow = one before next-insert position
        overrideAccount,
        cert.certification   || '',
        cert.measure         || '',
        cert.fileName        || '',
        cert.productCategory || ''
      );

      if (matchRow !== null) {
        // Update ONLY dates (cols I/J) and Comments (col O).
        // Col F (Certification) is intentionally NOT written — the sheet may
        // have a richer description ("ISO 9001:2015 (Quality Management System)")
        // that the AI would overwrite with a shorter form ("ISO 9001:2015").
        // Col O is appended, not overwritten, to preserve existing paragraphs.
        // Column numbers are hard-coded to prevent data-shift bugs caused by
        // column-map misresolution. Styling on existingRow is left 100% untouched.
        const existingRow = ws.getRow(matchRow);

        // col 9 (I) — Issued date
        existingRow.getCell(9).value = parseDateValue(cert.issueDate || '');

        // col 10 (J) — Date of Expiry
        const expiryStr = cert.expiryDate && cert.expiryDate !== 'Not Found'
          ? cert.expiryDate : '';
        existingRow.getCell(10).value =
          expiryStr ? (parseDateValue(expiryStr) ?? 'No Date') : 'No Date';

        // col 15 (O) — Append filename + cert number to Comments, but ONLY if
        // the filename is not already present. Guards against the same file
        // being re-uploaded and producing "file.pdf | file.pdf" spam.
        const existingComments = existingRow.getCell(15).value?.toString() ?? '';
        const fileAlreadyLogged = cert.fileName
          ? existingComments.includes(cert.fileName)
          : false;
        if (!fileAlreadyLogged) {
          const newEntry = [
            cert.fileName,
            cert.certificateNumber ? `Cert #${cert.certificateNumber}` : '',
          ].filter(Boolean).join(' | ');
          existingRow.getCell(15).value = existingComments
            ? `${existingComments} | ${newEntry}`
            : newEntry || null;
        }

        // col 4 (D) — Scope symbol: re-write value + force red font.
        // The update path skips applySmartAlignment to preserve client formatting,
        // but the red-font rule on the Scope column is non-negotiable — without it
        // the "!" / "+" symbols render black on updated rows.
        const scopeCell = existingRow.getCell(4);
        scopeCell.value  = cert.scope || '';
        scopeCell.numFmt = '@'; // prevent "+" / "!" being parsed as formula
        scopeCell.font   = { ...(scopeCell.font ?? {}), color: { argb: 'FFFF0000' } };

        existingRow.commit();
        updateCount++;
        console.log(
          `[appendToMasterExcel] Updated row ${matchRow} in-place: ` +
          `"${cert.supplierName}" → "${cert.certification}"`
        );
        continue; // ← skip insert
      }
    }

    // ── NO MATCH: insert a new row at the bottom of the supplier's block ──────
    ws.insertRow(insertAt, []);

    const newRow   = ws.getRow(insertAt);
    const aboveRow = ws.getRow(insertAt - 1); // row directly above

    // PASS 1 — Copy styles from the row directly above (deep-safe clone)
    aboveRow.eachCell({ includeEmpty: true }, (aboveCell, colNumber) => {
      safeCopyStyleAndValidation(aboveCell, newRow.getCell(colNumber));
    });

    // PASS 2 — Aggressive formula sourcing for H (Status) and K (Days to Expire).
    //
    // Scan ALL the way back to headerRowNum + 1 for a row with a live col K formula
    // (multiple "No Date" rows might sit above the insertion point, all with empty K).
    // Once found, pin every formula cell reference to insertAt and inject the
    // pre-calculated result so Protected View shows numbers without a calc engine.
    const kCol      = cols.daysToExpire; // already has ?? 11 from cols definition
    const statusCol = cols.status;       // already has ?? 8

    let formulaTemplateIdx = insertAt - 1;
    while (formulaTemplateIdx > headerRowNum + 1) {
      const kCell = ws.getRow(formulaTemplateIdx).getCell(kCol);
      if (kCell.type === ExcelJS.ValueType.Formula && kCell.formula) break;
      formulaTemplateIdx--;
    }
    const formulaTemplateRow = ws.getRow(formulaTemplateIdx);

    formulaTemplateRow.eachCell({ includeEmpty: true }, (templateCell, colNumber) => {
      if (templateCell.type === ExcelJS.ValueType.Formula && templateCell.formula) {
        const pinned = templateCell.formula.replace(
          /([A-Z]+)(\d+)/g,
          (_m, col) => `${col}${insertAt}`
        );
        // Inject pre-calculated result alongside formula.
        // Excel uses `result` as the displayed value when calc is blocked.
        if (colNumber === kCol) {
          newRow.getCell(colNumber).value = { formula: pinned, result: calculatedDays };
        } else if (colNumber === statusCol) {
          newRow.getCell(colNumber).value = { formula: pinned, result: calculatedStatus };
        } else {
          newRow.getCell(colNumber).value = { formula: pinned };
        }
      }
    });

    // PASS 3 — Write data values (H and K keep formula+result from PASS 2)
    const setCell = (colIndex: number | undefined, value: ExcelJS.CellValue | null) => {
      if (colIndex === undefined) return;
      newRow.getCell(colIndex).value = value ?? null;
    };

    // Visual grouping rule: only the FIRST row of a supplier block carries
    // Account (col 1), Name (col 2), Country (col 3). Every subsequent row
    // inside the same block MUST be strictly blank in those three columns.
    //
    // Using "" (empty string) rather than null — ExcelJS may silently skip
    // writing null on an inserted row that inherited a value from above,
    // whereas an explicit "" always overwrites the cell with an empty value.
    // Hard-coded column numbers 1/2/3 are used so the write is not dependent
    // on the column map resolving correctly.
    if (supplierFound) {
      // Existing block — enforce blanks regardless of what was inherited
      newRow.getCell(1).value = '';
      newRow.getCell(2).value = '';
      newRow.getCell(3).value = '';
    } else {
      // New supplier at the bottom.
      // Only the FIRST cert for this supplier in the current batch gets A/B/C.
      // Subsequent certs for the same supplier (keyed by account or name) get
      // blank A/B/C so the visual grouping rule is honoured across the whole batch.
      const supplierKey = overrideAccount || cert.supplierName || '';
      if (!newBlockStartedFor.has(supplierKey)) {
        setCell(cols.supplierAccount, overrideAccount     || null);
        setCell(cols.supplierName,    cert.supplierName   || null);
        setCell(cols.country,         cert.country        || null);
        newBlockStartedFor.add(supplierKey);
      } else {
        // Not the first row for this new supplier — blank A/B/C
        newRow.getCell(1).value = '';
        newRow.getCell(2).value = '';
        newRow.getCell(3).value = '';
      }
    }

    // col D (Scope): plain string — numFmt '@' in the alignment loop below forces
    // text storage, so no apostrophe hack needed.
    setCell(cols.scope, cert.scope || null);
    setCell(cols.measure,         cert.measure         || null);
    setCell(cols.certification,   cert.certification   || null);
    setCell(cols.productCategory, cert.productCategory || null);
    // col H (Status): formula + result from PASS 2 — never set as plain data
    setCell(cols.issued, parseDateValue(cert.issueDate || ''));

    // Expiry rule (§6): "No Date" if not stated — never null/empty
    const expiryStr = cert.expiryDate && cert.expiryDate !== 'Not Found'
      ? cert.expiryDate : '';
    const expiryValue: ExcelJS.CellValue =
      expiryStr ? (parseDateValue(expiryStr) ?? 'No Date') : 'No Date';
    setCell(cols.dateOfExpiry, expiryValue);
    // col K (Days to Expire): formula + result from PASS 2 — never set as plain data

    // Comments (O): filename + cert number if available
    const commentParts: string[] = [];
    if (cert.fileName)          commentParts.push(cert.fileName);
    if (cert.certificateNumber) commentParts.push(`Cert #${cert.certificateNumber}`);
    setCell(cols.comments, commentParts.join(' | ') || null);

    // Smart alignment: center data/date columns, left-align text-heavy columns.
    // Col 4 (Scope) also gets numFmt '@' + uniform red font.
    applySmartAlignment(newRow);
    newRow.commit();
    insertCount++;

    console.log(
      `[appendToMasterExcel] Inserted row at ${insertAt} for "${cert.supplierName}" ` +
      `(account: ${overrideAccount ?? 'n/a'}, existing block: ${supplierFound}, ` +
      `formula template: row ${formulaTemplateIdx})`
    );
  }

  console.log(
    `[appendToMasterExcel] Done — ${insertCount} inserted, ${updateCount} updated in "${ws.name}"`
  );

  // Step 7: Update Log (§12)
  // Scan from row 2 until col A is empty — avoids phantom rows that
  // logSheet.lastRow?.number returns when the sheet has formatting-only rows.
  const logSheet = workbook.getWorksheet('Update Log');
  if (logSheet) {
    let nextLogRow = 2;
    while (nextLogRow <= 100_000) {
      const v = logSheet.getRow(nextLogRow).getCell(1).value;
      if (v === null || v === undefined || String(v).trim() === '') break;
      nextLogRow++;
    }

    const now = new Date();
    const timestamp = `${now.toISOString().slice(0, 10)} ${now.toTimeString().slice(0, 5)}`;
    const first = processedCertificates[0];
    const logRow = logSheet.getRow(nextLogRow);
    logRow.getCell(1).value = timestamp;
    logRow.getCell(2).value = (first as any)._matchedAccount || '';
    logRow.getCell(3).value = first.supplierName || '';
    logRow.getCell(4).value = `Inserted ${processedCertificates.length} certificate(s)`;
    logRow.getCell(5).value = 'Appended via Vexos Engine. Expiry rules applied.';
    logRow.commit();
    console.log(`[appendToMasterExcel] Update Log row added at ${nextLogRow}`);
  } else {
    console.warn('[appendToMasterExcel] "Update Log" sheet not found — skipping');
  }

  // Step 8: Write and download
  // Force Excel to recalculate all formulas the moment the file is opened.
  // Without this, injected H/K formulas show as blank until the user
  // double-clicks a cell or presses Ctrl+Alt+F9.
  workbook.calcProperties.fullCalcOnLoad = true;

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer as ArrayBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const today = new Date().toISOString().slice(0, 10);
  saveAs(blob, `Updated_Master_File_${today}.xlsx`);
}
