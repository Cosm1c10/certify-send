import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { CertificateData } from '@/types/certificate';
import { DynamicSupplierMap, matchSupplier } from '@/utils/masterFileParser';

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
  // ExcelJS represents styled header cells (bold, coloured background, etc.) as
  // RichText objects: { richText: [{ text: 'Supplier Name', font: {...} }, ...] }
  // String(richTextObj) = "[object Object]" — extract plain text instead.
  if (value && typeof value === 'object' && 'richText' in value) {
    const parts = (value as any).richText;
    if (Array.isArray(parts)) {
      return parts.map((p: any) => p.text ?? '').join('').toLowerCase().trim().replace(/\s+/g, ' ');
    }
  }
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
 * Find the true last data row by scanning BACKWARDS from a capped upper bound.
 *
 * Why backwards instead of forwards:
 *   After expandSharedFormulas(), formula cells throughout the sheet (e.g. rows
 *   4–847 for Status/Days columns) have ExcelJS value objects even when their
 *   displayed result is an empty string. A forward scan checking `.value` would
 *   see these as non-empty and report row 847 (or wherever formulas end) as the
 *   last data row, pushing new supplier blocks to row 10,000+.
 *
 * Why we check cell.type and skip formula cells:
 *   Even when using `cell.text` (rendered string), formula cells can have
 *   non-empty CACHED results stored in the original file's <v> elements.
 *   These cached values survive expandSharedFormulas() and would fool the
 *   backward scan into thinking rows below the real data range are non-empty.
 *   Restricting to ValueType.String / .Number / .Date (i.e. plain data cells)
 *   eliminates formula-cache false positives entirely.
 *
 * Why three anchor columns (account, name, certification):
 *   Supplier Name (col B) can be blank for sub-rows (visual grouping rule).
 *   Certification (col F) may occasionally be formula-driven in some files.
 *   Supplier Account (col A) is always plain text and present on the FIRST
 *   row of every supplier block — checking it provides a rock-solid anchor
 *   that is independent of the other two columns.
 *
 * Scan cap:
 *   We never scan higher than startRow + MAX_SCAN_ROWS to avoid materialising
 *   tens of thousands of phantom ExcelJS Row/Cell objects in files whose
 *   formatting rows extend to row 10 000+. Any realistic master file fits well
 *   within 5 000 data rows.
 */
const MAX_SCAN_ROWS = 5_000;

function findTrueLastRow(
  ws: ExcelJS.Worksheet,
  supplierAccountCol: number,
  supplierNameCol: number,
  certCol: number,
  startRow: number
): number {
  // Cap: never scan more than MAX_SCAN_ROWS rows above startRow.
  // This prevents materialising phantom ExcelJS Row objects for every
  // formatting-only row in large files.
  const bottom = Math.min(ws.rowCount, startRow + MAX_SCAN_ROWS);

  /** Return true only for plain data cells (text/number/date) with visible content.
   *  Formula cells are explicitly excluded — their cached <v> values can make
   *  blank template rows appear non-empty even though they show nothing on screen.
   */
  function isRealData(cell: ExcelJS.Cell): boolean {
    const t = cell.type;
    if (
      t === ExcelJS.ValueType.Null      ||
      t === ExcelJS.ValueType.Formula   ||
      t === ExcelJS.ValueType.Error     ||
      t === ExcelJS.ValueType.Merge
    ) return false;
    return (cell.text ?? '').trim() !== '';
  }

  for (let r = bottom; r >= startRow; r--) {
    const row = ws.getRow(r);
    if (
      isRealData(row.getCell(supplierAccountCol)) ||
      isRealData(row.getCell(supplierNameCol))    ||
      isRealData(row.getCell(certCol))
    ) {
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
  overrideAccount: string | undefined,
  supplierNameHint?: string,   // fallback: search by col B when account search fails
  productCategoryCol: number = 7
): { insertAt: number; supplierFound: boolean; resolvedAccount: string | undefined; blockStartRow: number | undefined } {
  const trueLastRow = findTrueLastRow(ws, supplierAccountCol, supplierNameCol, certCol, headerRowNum + 1);

  // Helper: does a row carry visible cert/product data (i.e. is it a real sub-row)?
  // Must mirror findTrueLastRow's isRealData guard: skip formula/error/merge/null
  // cells because their cached <v> values look non-empty but represent no real data.
  const isSubRow = (r: number): boolean => {
    const row = ws.getRow(r);
    const certCell = row.getCell(certCol);
    const prodCell = row.getCell(productCategoryCol);
    const skipTypes = new Set([
      ExcelJS.ValueType.Null,
      ExcelJS.ValueType.Formula,
      ExcelJS.ValueType.Error,
      ExcelJS.ValueType.Merge,
    ]);
    const hasData = (cell: ExcelJS.Cell) =>
      !skipTypes.has(cell.type) && (cell.text ?? '').trim() !== '';
    return hasData(certCell) || hasData(prodCell);
  };

  // ── Primary search: account code in col A ─────────────────────────────────
  // After the anchor row is found, sub-rows (blank col A but data in col F/G)
  // are consumed as part of the same block. The block ends at the first
  // separator row (blank col A AND no cert/product data) or a different supplier.
  if (overrideAccount) {
    const normalizedTarget = overrideAccount.toLowerCase().trim();
    let firstSupplierRow = -1;
    let lastSupplierRow  = -1;
    let inBlock = false;

    for (let r = headerRowNum + 1; r <= trueLastRow; r++) {
      const colAVal = ws.getRow(r).getCell(supplierAccountCol).value;
      const colAStr = colAVal !== null && colAVal !== undefined ? String(colAVal).trim() : '';

      if (colAStr !== '') {
        if (colAStr.toLowerCase() === normalizedTarget) {
          // Anchor row (or repeat) for our supplier
          if (firstSupplierRow < 0) firstSupplierRow = r;
          lastSupplierRow = r;
          inBlock = true;
        } else if (inBlock) {
          // Different supplier account — block has ended
          break;
        }
        // else: a different supplier before ours — keep scanning
      } else if (inBlock) {
        // Col A is blank while inside our block — sub-row or separator
        if (isSubRow(r)) {
          lastSupplierRow = r;  // extend block to cover this sub-row
        } else {
          break;  // blank separator row — block has ended
        }
      }
      // else: blank row before we've found our supplier — keep scanning
    }

    if (lastSupplierRow >= 0) {
      console.log(
        `[findInsertionRow] Account "${overrideAccount}" last row: ${lastSupplierRow} → inserting at ${lastSupplierRow + 1}`
      );
      return { insertAt: lastSupplierRow + 1, supplierFound: true, resolvedAccount: overrideAccount, blockStartRow: firstSupplierRow };
    }
  }

  // ── Fallback search: supplier name in col B ────────────────────────────────
  // Runs when overrideAccount is missing (matchedAccount not in supplierMap)
  // OR when the account code in the sheet doesn't match the map's value.
  // Same block-depth logic: after the anchor row is found, consume sub-rows
  // (blank col B but data in col F/G) until a separator or new supplier appears.
  if (supplierNameHint) {
    const normalizedName = supplierNameHint.toLowerCase().trim();
    let lastSupplierRow  = -1;
    let firstSupplierRow = -1;
    let inBlock = false;

    for (let r = headerRowNum + 1; r <= trueLastRow; r++) {
      const colBVal = ws.getRow(r).getCell(supplierNameCol).value;
      const colBStr = colBVal !== null && colBVal !== undefined ? String(colBVal).trim() : '';

      if (colBStr !== '') {
        if (colBStr.toLowerCase() === normalizedName) {
          if (firstSupplierRow < 0) firstSupplierRow = r;
          lastSupplierRow = r;
          inBlock = true;
        } else if (inBlock) {
          // Different supplier name — block has ended
          break;
        }
        // else: a different supplier before ours — keep scanning
      } else if (inBlock) {
        // Col B is blank while inside our block — sub-row or separator
        if (isSubRow(r)) {
          lastSupplierRow = r;  // extend block to cover this sub-row
        } else {
          break;  // blank separator row — block has ended
        }
      }
      // else: blank row before we've found our supplier — keep scanning
    }

    if (lastSupplierRow >= 0) {
      // Extract the account code from col A of the first row of this block
      const accountCell = ws.getRow(firstSupplierRow).getCell(supplierAccountCol).value;
      const sheetAccount = accountCell ? String(accountCell).trim() || undefined : undefined;
      console.log(
        `[findInsertionRow] Name fallback "${supplierNameHint}" last row: ${lastSupplierRow} ` +
        `→ inserting at ${lastSupplierRow + 1} (account from sheet: "${sheetAccount ?? 'n/a'}")`
      );
      return { insertAt: lastSupplierRow + 1, supplierFound: true, resolvedAccount: sheetAccount ?? overrideAccount, blockStartRow: firstSupplierRow };
    }
  }

  // +2 instead of +1: leaves one blank separator row between the last
  // existing supplier block and the new one, matching the sheet's visual style.
  console.log(
    `[findInsertionRow] Supplier "${overrideAccount ?? supplierNameHint ?? '(none)'}" not found → inserting at trueLastRow+2 = ${trueLastRow + 2}`
  );
  return { insertAt: trueLastRow + 2, supplierFound: false, resolvedAccount: overrideAccount, blockStartRow: undefined };
}

// -----------------------------------------------------------------------------
// Deduplication helper
// -----------------------------------------------------------------------------

/**
 * Scan a supplier's pre-identified row block (startRow → lastSupplierRow) for
 * an existing certificate that matches the incoming cert.
 *
 * CRITICAL DESIGN RULES:
 *   • startRow and lastSupplierRow come from findInsertionRow — we NEVER
 *     re-scan Col A or Col B inside this loop. Sub-rows have a blank Col A;
 *     checking it would skip them silently and miss every update-in-place.
 *   • Column indices are taken from the parsed `cols` map (dynamic), not
 *     hard-coded, so the function works regardless of actual sheet layout.
 *   • Cert names from both the incoming cert and the sheet are normalised
 *     aggressively (lowercase + strip all non-alphanumeric chars) before
 *     comparison so "BRC", "BRCGS", "BRC Global Standards" all collapse to
 *     comparable strings.
 *   • The Measure column text is concatenated with the Certification text
 *     before matching so that "BRCGS" in either column is detected.
 *
 * Returns the matching row number, or null if no match.
 */
function findExistingCertRow(
  ws: ExcelJS.Worksheet,
  startRow: number,        // first row of supplier's block (from findInsertionRow)
  lastSupplierRow: number, // last row of supplier's block  (= insertAt - 1)
  cert: CertificateData,
  cols: Record<string, number>
): number | null {
  if (startRow < 0 || lastSupplierRow < startRow) return null;

  // Aggressive normalisation: lowercase + strip every non-alphanumeric char.
  const norm = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  const normalizedNew = norm(cert.certification || '');
  if (!normalizedNew) return null;

  // Filename check prep (Criterion A)
  const normInFile        = (cert.fileName || '').toLowerCase().trim();
  const normInFileCleaned = normInFile.replace(/\s*\(\d+\)(\.\w+)$/, '$1');

  for (let r = startRow; r <= lastSupplierRow; r++) {
    const row = ws.getRow(r);

    // --- CRITERION A: Filename in Comments ---
    // Checked first — it's the most authoritative signal (same physical file).
    const commentsVal = row.getCell(cols.comments).value;
    const comments    = commentsVal ? String(commentsVal).toLowerCase() : '';
    if (normInFile && comments.includes(normInFile)) {
      console.log(`[findExistingCertRow] Filename match at row ${r}: "${cert.fileName}"`);
      return r;
    }
    if (normInFileCleaned !== normInFile && normInFileCleaned && comments.includes(normInFileCleaned)) {
      console.log(`[findExistingCertRow] Cleaned-filename match at row ${r}: "${normInFileCleaned}"`);
      return r;
    }

    // --- CRITERION B: Cert alias match ---
    // Read Certification (col F) and Measure (col E); combine so that an
    // acronym appearing in either column is still found.
    const certCellVal    = row.getCell(cols.certification).value;
    const measureCellVal = row.getCell(cols.measure).value;

    const existingCertText    = certCellVal    ? String(certCellVal)    : '';
    const existingMeasureText = measureCellVal ? String(measureCellVal) : '';
    const combinedExisting    = norm(existingCertText) + norm(existingMeasureText);

    if (!combinedExisting) continue;

    let isMatch = false;

    // BRC / BRCGS / BRC Global Standards
    if (normalizedNew.includes('brc') && combinedExisting.includes('brc')) {
      isMatch = true;
    }
    // FSC / Forest Stewardship Council
    else if (
      (normalizedNew.includes('fsc') || normalizedNew.includes('foreststewardship')) &&
      (combinedExisting.includes('fsc') || combinedExisting.includes('foreststewardship'))
    ) {
      isMatch = true;
    }
    // GRS / Global Recycled Standard
    else if (
      (normalizedNew.includes('grs') || normalizedNew.includes('globalrecycled')) &&
      (combinedExisting.includes('grs') || combinedExisting.includes('globalrecycled'))
    ) {
      isMatch = true;
    }
    // ISO — require matching number so ISO 9001 ≠ ISO 14001
    else if (normalizedNew.includes('iso') && combinedExisting.includes('iso')) {
      const isoNumNew = normalizedNew.match(/iso(\d+)/)?.[1];
      const isoNumEx  = combinedExisting.match(/iso(\d+)/)?.[1];
      if (isoNumNew && isoNumEx && isoNumNew === isoNumEx) isMatch = true;
    }
    // GOTS / Global Organic Textile Standard
    else if (
      (normalizedNew.includes('gots') || normalizedNew.includes('globalorganic')) &&
      (combinedExisting.includes('gots') || combinedExisting.includes('globalorganic'))
    ) {
      isMatch = true;
    }
    // OEKO-TEX / Standard 100
    else if (
      (normalizedNew.includes('oekotex') || normalizedNew.includes('oekotek') || normalizedNew.includes('standard100')) &&
      (combinedExisting.includes('oekotex') || combinedExisting.includes('oekotek') || combinedExisting.includes('standard100'))
    ) {
      isMatch = true;
    }
    // RSPO / Roundtable on Sustainable Palm Oil
    else if (
      (normalizedNew.includes('rspo') || normalizedNew.includes('sustainablepalm')) &&
      (combinedExisting.includes('rspo') || combinedExisting.includes('sustainablepalm'))
    ) {
      isMatch = true;
    }
    // SEDEX
    else if (normalizedNew.includes('sedex') && combinedExisting.includes('sedex')) {
      isMatch = true;
    }
    // SMETA
    else if (normalizedNew.includes('smeta') && combinedExisting.includes('smeta')) {
      isMatch = true;
    }
    // Rainforest Alliance
    else if (normalizedNew.includes('rainforest') && combinedExisting.includes('rainforest')) {
      isMatch = true;
    }
    // Halal
    else if (normalizedNew.includes('halal') && combinedExisting.includes('halal')) {
      isMatch = true;
    }
    // Kosher
    else if (normalizedNew.includes('kosher') && combinedExisting.includes('kosher')) {
      isMatch = true;
    }
    // Generic fallback: bidirectional containment
    else if (normalizedNew.length >= 4 && combinedExisting.includes(normalizedNew)) {
      isMatch = true;
    }
    else if (combinedExisting.length >= 4 && normalizedNew.includes(combinedExisting)) {
      isMatch = true;
    }

    if (isMatch) {
      console.log(
        `[findExistingCertRow] Match at row ${r}: ` +
        `existing="${existingCertText}", incoming="${cert.certification}"`
      );
      return r;
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

  // BYPASS: Do not use prepareExportData/applySaurabhFilter because it deletes
  // certificates that share the same (or missing) certification name before
  // the Fuzzy Deduplicator can properly assess them against the existing Master File.

  const processedCertificates = certificates.map(cert => {
    // Basic supplier mapping without dropping any files
    if (supplierMap && Object.keys(supplierMap).length > 0) {
      const matchResult = matchSupplier(cert.supplierName, supplierMap, 0.75);
      if (matchResult.wasMatched) {
         (cert as any)._matchedAccount = matchResult.matchedAccount;
         cert.supplierName = matchResult.matchedName;
      }
    }
    return cert;
  });

  if (processedCertificates.length === 0) {
    alert('No certificates to process.');
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
    supplierAccount: colIdx(['supplier account', 'account'])              ?? 1,
    supplierName:    colIdx(['supplier name', 'supplier'])                ?? 2,
    country:         colIdx(['country'])                                  ?? 3,
    scope:           colIdx(['scope'])                                    ?? 4,
    measure:         colIdx(['measure'])                                  ?? 5,
    certification:   colIdx(['certification', 'certification '])          ?? 6,
    productCategory: colIdx(['product category', 'product category '])   ?? 7,
    status:          colIdx(['status'])                                   ?? 8,
    issued:          colIdx(['issued', 'date issued', 'issue date'])      ?? 9,
    dateOfExpiry:    colIdx(['date of expiry', 'expiry date', 'expiry']) ?? 10,
    daysToExpire:    colIdx(['days to expire', 'days to expiry'])        ?? 11,
    comments:        colIdx(['comments', 'comment'])                     ?? 15,
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

  // ── In-memory tracking for this batch run ────────────────────────────────────
  //
  //   newBlockStartedFor
  //     Tracks which supplier keys have already had their first (Account/Name/
  //     Country) row written, so subsequent certs in the block get blank A/B/C.
  //
  //   batchInsertPtr
  //     STRICT SEQUENTIAL INSERTION — root-cause fix for "ghost rows" when a
  //     batch contains incomplete certs.
  //
  //     Problem: each cert re-calls findInsertionRow which calls findTrueLastRow.
  //     findTrueLastRow is blind to sub-rows that have blank anchor columns
  //     (A/B/F) — e.g. the 2nd and later rows of a new supplier block when the
  //     AI failed to extract a cert name.  Every time a new incomplete row is
  //     inserted, findTrueLastRow snapbacks to the FIRST row of the block and
  //     re-anchors insertAt = lastSupplierRow + 1.  Subsequent certs keep
  //     inserting at that same slot, pushing earlier ones down — the block ends
  //     up in reverse order, or (for truly blank rows) at a completely wrong
  //     position near the bottom of the sheet.
  //
  //     Fix: after the FIRST insert for a supplier in this batch, record
  //     insertAt + 1 in batchInsertPtr.  Every subsequent cert for the same
  //     supplier reads the pointer and increments it after each insert, giving
  //     guaranteed sequential rows regardless of whether those rows have any
  //     data in the anchor columns.
  //
  //   batchSupplierWasFound
  //     Stores whether the supplier's block pre-existed in the sheet (needed to
  //     decide: run update-in-place check vs. always-insert for new suppliers).
  // ─────────────────────────────────────────────────────────────────────────────
  const newBlockStartedFor            = new Set<string>();
  const batchInsertPtr                = new Map<string, number>();
  const batchSupplierWasFound         = new Map<string, boolean>();
  const batchResolvedAccount          = new Map<string, string | undefined>();
  const batchBlockStartRow            = new Map<string, number | undefined>();

  for (const cert of processedCertificates) {
    const overrideAccount: string | undefined = (cert as any)._matchedAccount || undefined;
    const supplierKey = overrideAccount || cert.supplierName || '';

    // ── Determine insertion point ─────────────────────────────────────────────
    // First encounter for this supplier key in this batch: scan the sheet once.
    // Every subsequent encounter: use the pre-tracked pointer so that rows
    // already inserted (which may have blank anchor columns A/B/F) never confuse
    // findTrueLastRow and reverse the insertion order.
    let insertAt: number;
    let supplierFound: boolean;
    let resolvedAccount: string | undefined;
    let blockStartRow: number | undefined;

    if (batchInsertPtr.has(supplierKey)) {
      insertAt        = batchInsertPtr.get(supplierKey)!;
      supplierFound   = batchSupplierWasFound.get(supplierKey)!;
      resolvedAccount = batchResolvedAccount.get(supplierKey);
      blockStartRow   = batchBlockStartRow.get(supplierKey);
    } else {
      ({ insertAt, supplierFound, resolvedAccount, blockStartRow } = findInsertionRow(
        ws,
        cols.supplierAccount  ?? 1,
        cols.supplierName     ?? 2,
        cols.certification    ?? 6,
        headerRowNum,
        overrideAccount,
        cert.supplierName,              // name-based fallback
        cols.productCategory  ?? 7      // sub-row detection
      ));
      batchResolvedAccount.set(supplierKey, resolvedAccount);
      batchBlockStartRow.set(supplierKey, blockStartRow);
    }

    // ── Pre-calculate Status & Days for Protected View ────────────────────────
    // Injected as formula `result` so Excel shows them in Protected View.
    const expiryForCalc = cert.expiryDate && cert.expiryDate !== 'Not Found'
      ? cert.expiryDate : '';
    let calculatedDays: number | string = '';
    let calculatedStatus: string = '';
    if (expiryForCalc) {
      const expiryMs = new Date(expiryForCalc).getTime();
      if (!isNaN(expiryMs)) {
        const days = Math.floor((expiryMs - Date.now()) / (1000 * 60 * 60 * 24));
        calculatedDays   = days;
        calculatedStatus = days < 0 ? 'Expired' : 'Up to date';
      }
    }

    // ── Update-in-Place (only for pre-existing supplier blocks) ──────────────
    // For new supplier blocks we are building this session, skip the duplicate
    // check — the block didn't exist before, so there is nothing to match.
    if (supplierFound) {
      const matchRow = findExistingCertRow(
        ws,
        blockStartRow ?? -1,  // first row of supplier's block (from findInsertionRow)
        insertAt - 1,         // lastSupplierRow = one before next-insert position
        cert,
        cols
      );

      if (matchRow !== null) {
        // Update ONLY dates (cols I/J) and Comments (col O).
        // Col F (Certification) is intentionally NOT written — the sheet may
        // have a richer description ("ISO 9001:2015 (Quality Management System)")
        // that the AI would overwrite with a shorter form ("ISO 9001:2015").
        // Col O is appended, not overwritten, to preserve existing content.
        // Column numbers are hard-coded to prevent data-shift bugs.
        // Do NOT advance batchInsertPtr — no row was inserted, so the next
        // cert should still land at the same sequential position.
        const existingRow = ws.getRow(matchRow);

        // col 9 (I) — Issued date
        existingRow.getCell(9).value = parseDateValue(cert.issueDate || '');

        // col 10 (J) — Date of Expiry
        const expiryStrU = cert.expiryDate && cert.expiryDate !== 'Not Found'
          ? cert.expiryDate : '';
        existingRow.getCell(10).value =
          expiryStrU ? (parseDateValue(expiryStrU) ?? 'No Date') : 'No Date';

        // col 15 (O) — Append filename + cert number, guard against spam
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
        const scopeCell = existingRow.getCell(4);
        scopeCell.value  = cert.scope || '';
        scopeCell.numFmt = '@';
        scopeCell.font   = { ...(scopeCell.font ?? {}), color: { argb: 'FFFF0000' } };

        existingRow.commit();
        updateCount++;
        console.log(
          `[appendToMasterExcel] Updated row ${matchRow} in-place: ` +
          `"${cert.supplierName}" → "${cert.certification}"`
        );
        continue; // ← skip insert; do NOT advance batchInsertPtr
      }
    }

    // ── NEW ROW: insert strictly at the tracked sequential position ───────────
    ws.insertRow(insertAt, []);

    const newRow   = ws.getRow(insertAt);
    const aboveRow = ws.getRow(insertAt - 1);

    // Inherit row height so new rows blend in visually
    if (aboveRow.height) {
      newRow.height = aboveRow.height;
    }

    // PASS 1 — Copy styles from the row directly above (deep-safe clone)
    aboveRow.eachCell({ includeEmpty: true }, (aboveCell, colNumber) => {
      safeCopyStyleAndValidation(aboveCell, newRow.getCell(colNumber));
    });

    // PASS 2 — Aggressive formula sourcing for H (Status) and K (Days to Expire).
    // Scan backwards from insertAt-1 for a row with a live col K formula so that
    // "No Date" rows (which have empty K) above the insertion point don't block
    // formula inheritance. Pin all references to insertAt and inject results.
    const kCol      = cols.daysToExpire; // already has ?? 11 from cols definition
    const statusCol = cols.status;       // already has ?? 8

    let formulaTemplateIdx = insertAt - 1;
    while (formulaTemplateIdx > headerRowNum + 1) {
      const kCell = ws.getRow(formulaTemplateIdx).getCell(kCol);
      if (kCell.type === ExcelJS.ValueType.Formula && kCell.formula) break;
      formulaTemplateIdx--;
    }
    const formulaTemplateRow = ws.getRow(formulaTemplateIdx);

    // Track whether PASS 2 successfully wrote a live formula to H and K.
    // These flags are set inside the eachCell callback and used by the fallback
    // below — more reliable than re-querying cell.type after the fact, since
    // ExcelJS may not update the in-memory type immediately after assignment.
    let statusFormulaWritten = false;
    let daysFormulaWritten   = false;

    formulaTemplateRow.eachCell({ includeEmpty: true }, (templateCell, colNumber) => {
      if (templateCell.type === ExcelJS.ValueType.Formula && templateCell.formula) {
        // Col 4 (Scope / col D) is plain text written in PASS 3 — never propagate
        // a formula here or it could inject a stray "+" / "!" before PASS 3 runs.
        if (colNumber === 4) return;

        const pinned = templateCell.formula.replace(
          /([A-Z]+)(\d+)/g,
          (_m, col) => `${col}${insertAt}`
        );
        if (colNumber === kCol) {
          newRow.getCell(colNumber).value = { formula: pinned, result: calculatedDays };
          daysFormulaWritten = true;
        } else if (colNumber === statusCol) {
          newRow.getCell(colNumber).value = { formula: pinned, result: calculatedStatus };
          statusFormulaWritten = true;
        } else {
          newRow.getCell(colNumber).value = { formula: pinned };
        }
      }
    });

    // PASS 2 fallback — only runs when PASS 2 did NOT write a formula to H/K.
    // Writes the pre-calculated primitive value directly so those cells are
    // never blank. Mutually exclusive with the formula path above.
    if (!statusFormulaWritten && calculatedStatus !== '') {
      newRow.getCell(statusCol ?? 8).value = calculatedStatus;
    }
    if (!daysFormulaWritten && calculatedDays !== '') {
      newRow.getCell(kCol ?? 11).value = calculatedDays;
    }

    // PASS 3 — Write data values (H and K keep formula+result from PASS 2)
    const setCell = (colIndex: number | undefined, value: ExcelJS.CellValue | null) => {
      if (colIndex === undefined) return;
      newRow.getCell(colIndex).value = value ?? null;
    };

    // Visual grouping: only the FIRST row of a supplier block carries A/B/C.
    // Hard-coded cols 1/2/3 — independent of the column map.
    if (supplierFound) {
      // Existing block — enforce blanks regardless of what was inherited
      newRow.getCell(1).value = '';
      newRow.getCell(2).value = '';
      newRow.getCell(3).value = '';
    } else {
      // New supplier at the bottom: only the first cert of this key gets A/B/C.
      if (!newBlockStartedFor.has(supplierKey)) {
        setCell(cols.supplierAccount, overrideAccount   || null);
        setCell(cols.supplierName,    cert.supplierName || null);
        setCell(cols.country,         cert.country      || null);
        newBlockStartedFor.add(supplierKey);
      } else {
        newRow.getCell(1).value = '';
        newRow.getCell(2).value = '';
        newRow.getCell(3).value = '';
      }
    }

    // col D (Scope): numFmt '@' forced by applySmartAlignment below
    setCell(cols.scope,           cert.scope           || null);
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

    applySmartAlignment(newRow);
    newRow.commit();
    insertCount++;

    // ── Advance the per-supplier pointer ──────────────────────────────────────
    // Record (or update) the NEXT available row for this supplier key so the
    // following cert lands directly below — even if it has no data in the
    // anchor columns (A/B/F) that findTrueLastRow uses to detect real rows.
    batchInsertPtr.set(supplierKey, insertAt + 1);
    if (!batchSupplierWasFound.has(supplierKey)) {
      batchSupplierWasFound.set(supplierKey, supplierFound);
    }

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
