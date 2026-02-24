import * as XLSX from 'xlsx';

// =============================================================================
// MASTER FILE PARSER
// Parses the client's Master Excel file to build DYNAMIC_SUPPLIER_MAP
// Uses SheetJS (xlsx) for robust parsing of various Excel formats
// =============================================================================

export interface SupplierEntry {
  officialName: string;      // Exact name as it appears in Master File
  country?: string;          // Country if available
  supplierAccount?: string;  // Account code if available
}

export interface DynamicSupplierMap {
  // Keyed by skeleton key (normalized supplier name)
  [skeletonKey: string]: SupplierEntry;
}

export interface MasterFileData {
  supplierMap: DynamicSupplierMap;
  totalSuppliers: number;
  fileName: string;
}

/**
 * Strip Unicode diacritics/accents from a string.
 * e.g., "Ömeroğlu" → "Omeroglu", "Şti" → "Sti", "Ürün" → "Urun"
 * Uses Unicode NFD decomposition + removal of combining marks.
 * Also handles special Turkish characters that NFD doesn't fully cover.
 */
function stripDiacritics(text: string): string {
  // NFD decompose then remove combining marks (handles most accented chars)
  let result = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  // Handle special Turkish characters not fully covered by NFD
  result = result
    .replace(/ğ/g, 'g').replace(/Ğ/g, 'G')
    .replace(/ı/g, 'i').replace(/İ/g, 'I')
    .replace(/ş/g, 's').replace(/Ş/g, 'S')
    .replace(/ç/g, 'c').replace(/Ç/g, 'C')
    .replace(/ü/g, 'u').replace(/Ü/g, 'U')
    .replace(/ö/g, 'o').replace(/Ö/g, 'O')
    // Polish
    .replace(/ł/g, 'l').replace(/Ł/g, 'L')
    .replace(/ą/g, 'a').replace(/ę/g, 'e')
    .replace(/ź/g, 'z').replace(/ż/g, 'z')
    .replace(/ń/g, 'n').replace(/ś/g, 's')
    // German
    .replace(/ä/g, 'a').replace(/Ä/g, 'A')
    .replace(/ß/g, 'ss');

  return result;
}

/**
 * Generate a "skeleton key" from a supplier name for fuzzy matching.
 *
 * Rules:
 * 1. Strip diacritics/accents (ö→o, ğ→g, ş→s, etc.)
 * 2. Convert to lowercase
 * 3. Remove all punctuation
 * 4. Remove common company suffixes (Co, Ltd, Inc, GmbH, etc.)
 * 5. Remove extra whitespace
 * 6. Sort words alphabetically (for word-order independence)
 */
export function generateSkeletonKey(name: string): string {
  if (!name) return '';

  // Strip diacritics FIRST, then lowercase
  let key = stripDiacritics(name).toLowerCase();

  // Remove common company suffixes (comprehensive list for global suppliers)
  const suffixes = [
    // English
    'company', 'co', 'corp', 'corporation', 'inc', 'incorporated', 'ltd', 'limited',
    'llc', 'llp', 'lp', 'plc', 'pvt', 'private',
    // German
    'gmbh', 'ag', 'kg', 'ohg', 'gbr', 'e.v', 'ev', 'mbh',
    // French
    'sa', 'sarl', 'sas', 'sasu', 'snc', 'eurl',
    // Italian
    'spa', 'srl', 'snc', 'sas',
    // Spanish
    'sl', 'sa', 'slu', 'sau',
    // Dutch
    'bv', 'nv', 'vof', 'cv',
    // Turkish (after diacritic stripping: ş→s, ı→i, ö→o, ü→u, ğ→g, ç→c)
    'as', 'anonim', 'sirketi', 'tic', 'ticaret', 'san', 'sanayi', 'sti',
    'tarim', 'urun', 'gida', 've', 'dis',
    // Chinese
    '有限公司', '股份有限公司', '集团',
    // Polish
    'sp', 'zoo', 'spzoo', 'spółka',
    // Other common terms
    'international', 'intl', 'trading', 'group', 'holding', 'holdings', 'industries', 'industrial',
    'manufacturing', 'mfg', 'products', 'services', 'solutions', 'enterprises', 'enterprise',
    'packaging', 'paper', 'plastic', 'plastics', 'chemicals',
  ];

  // Remove punctuation and special characters, but keep spaces
  key = key.replace(/[.,\-_'"()&+!?@#$%^*=[\]{}|\\/<>:;]/g, ' ');

  // Split into words
  let words = key.split(/\s+/).filter(w => w.length > 0);

  // Remove suffixes
  words = words.filter(word => !suffixes.includes(word));

  // Remove single letters and numbers (likely abbreviations or address numbers)
  words = words.filter(word => word.length > 1 && !/^\d+$/.test(word));

  // Sort alphabetically for word-order independence
  words.sort();

  // Join with single space
  return words.join(' ').trim();
}

/**
 * Calculate similarity between two skeleton keys using Levenshtein distance
 * Returns a score from 0 to 1 (1 = identical)
 */
export function calculateSimilarity(key1: string, key2: string): number {
  if (key1 === key2) return 1;
  if (key1.length === 0 || key2.length === 0) return 0;

  // Check if one is a substring of the other
  if (key1.includes(key2) || key2.includes(key1)) {
    return 0.9;
  }

  // Levenshtein distance
  const matrix: number[][] = [];

  for (let i = 0; i <= key1.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= key2.length; j++) {
    matrix[0][j] = j;
  }

  for (let i = 1; i <= key1.length; i++) {
    for (let j = 1; j <= key2.length; j++) {
      if (key1[i - 1] === key2[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  const distance = matrix[key1.length][key2.length];
  const maxLength = Math.max(key1.length, key2.length);
  return 1 - (distance / maxLength);
}

/**
 * Parse the Master Excel file and build DYNAMIC_SUPPLIER_MAP
 * Uses SheetJS (xlsx) for robust parsing - handles files with drawings, anchors, etc.
 */
export async function parseMasterFile(file: File): Promise<MasterFileData> {
  const arrayBuffer = await file.arrayBuffer();

  // Parse with SheetJS - much more robust than ExcelJS for reading
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(arrayBuffer, { type: 'array' });
  } catch (error) {
    console.error('Failed to parse Excel file:', error);
    throw new Error(`Unable to parse Excel file. Please ensure it is a valid .xlsx or .xls format.`);
  }

  const supplierMap: DynamicSupplierMap = {};
  let totalSuppliers = 0;

  // Smart sheet selection: find the sheet that actually contains supplier data
  // Priority: 1) sheet with "Supplier Name" header, 2) sheet named "Certificates", 3) skip "Instructions"
  console.log(`[MASTER FILE] Available sheets: ${workbook.SheetNames.join(', ')}`);

  let sheetName: string | undefined;
  let data: any[][] = [];

  // Strategy 1: Find a sheet that contains "Supplier Name" in the first 10 rows
  for (const name of workbook.SheetNames) {
    const ws = workbook.Sheets[name];
    if (!ws) continue;
    const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
    for (let r = 0; r < Math.min(10, rows.length); r++) {
      const row = rows[r] || [];
      for (const cell of row) {
        const val = String(cell || '').toLowerCase().trim();
        if (val.includes('supplier name') || val.includes('supplier account')) {
          sheetName = name;
          data = rows;
          console.log(`[MASTER FILE] Found supplier data in sheet "${name}" at row ${r}`);
          break;
        }
      }
      if (sheetName) break;
    }
    if (sheetName) break;
  }

  // Strategy 2: Fallback to sheet named "Certificates" or similar
  if (!sheetName) {
    sheetName = workbook.SheetNames.find(name =>
      name.toLowerCase() === 'certificates' || name.toLowerCase().includes('certificate')
    );
  }

  // Strategy 3: Use first sheet that isn't "Instructions"
  if (!sheetName) {
    sheetName = workbook.SheetNames.find(name =>
      name.toLowerCase() !== 'instructions'
    ) || workbook.SheetNames[0];
  }

  if (!sheetName) {
    throw new Error('No worksheet found in Master File');
  }

  // If we haven't loaded data yet (strategies 2 or 3), load it now
  if (data.length === 0) {
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      throw new Error('Unable to read worksheet from Master File');
    }
    data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  }

  if (data.length < 2) {
    throw new Error('Master File appears to be empty or has no data rows');
  }

  // Find the header row — it may not be row 0 (title rows, merged cells, etc.)
  let headerRowIndex = -1;
  let supplierNameCol = -1;
  let countryCol = -1;
  let supplierAccountCol = -1;

  // Scan first 10 rows for the header
  for (let r = 0; r < Math.min(10, data.length); r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      const value = String(row[c] || '').toLowerCase().trim();
      if (value.includes('supplier name') || value === 'supplier') {
        supplierNameCol = c;
        headerRowIndex = r;
      }
      if (value === 'country') {
        countryCol = c;
      }
      if (value.includes('supplier account') || value === 'account') {
        supplierAccountCol = c;
      }
    }
    if (headerRowIndex !== -1) break;
  }

  // V7 Master File expected columns:
  // Col 0: Supplier Account, Col 1: Supplier Name, Col 2: Country (0-indexed)
  if (supplierNameCol === -1) {
    // Fallback to V7 defaults — assume header at row 0
    headerRowIndex = 0;
    supplierAccountCol = 0;
    supplierNameCol = 1;
    countryCol = 2;
    console.log('Using V7 default column positions (header not found by name)');
  }

  console.log(`Master File: sheet "${sheetName}", header at row ${headerRowIndex}, columns — Name: ${supplierNameCol}, Country: ${countryCol}, Account: ${supplierAccountCol}`);

  // Iterate through data rows (skip header row)
  for (let i = headerRowIndex + 1; i < data.length; i++) {
    const row = data[i];
    if (!row) continue;

    const supplierName = String(row[supplierNameCol] || '').trim();
    if (!supplierName) continue;

    const skeletonKey = generateSkeletonKey(supplierName);
    if (!skeletonKey) continue;

    // Only add if not already in map (first occurrence wins)
    if (!supplierMap[skeletonKey]) {
      supplierMap[skeletonKey] = {
        officialName: supplierName,
        country: countryCol >= 0 ? String(row[countryCol] || '').trim() : undefined,
        supplierAccount: supplierAccountCol >= 0 ? String(row[supplierAccountCol] || '').trim() : undefined,
      };
      totalSuppliers++;
    }
  }

  console.log(`Parsed Master File: ${totalSuppliers} unique suppliers`);

  return {
    supplierMap,
    totalSuppliers,
    fileName: file.name,
  };
}

/**
 * Calculate word-overlap score between two skeleton keys.
 * Returns the fraction of words in the shorter key that appear in the longer key.
 * A single shared distinctive word (like a brand name) can trigger a match.
 */
function calculateWordOverlap(key1: string, key2: string): number {
  const words1 = key1.split(/\s+/).filter(w => w.length > 2);
  const words2 = key2.split(/\s+/).filter(w => w.length > 2);

  if (words1.length === 0 || words2.length === 0) return 0;

  const shorter = words1.length <= words2.length ? words1 : words2;
  const longer = words1.length <= words2.length ? words2 : words1;
  const longerSet = new Set(longer);

  let exactMatches = 0;
  let fuzzyMatches = 0;

  for (const word of shorter) {
    if (longerSet.has(word)) {
      exactMatches++;
    } else {
      // Check fuzzy word match (e.g., "sterilizasyon" ≈ "sterilization")
      for (const lWord of longer) {
        if (calculateSimilarity(word, lWord) >= 0.8) {
          fuzzyMatches++;
          break;
        }
      }
    }
  }

  const totalMatches = exactMatches + fuzzyMatches * 0.8;
  return totalMatches / shorter.length;
}

/**
 * Find the best matching supplier from the DYNAMIC_SUPPLIER_MAP
 * Uses multi-strategy matching:
 *   1. Exact skeleton key match (confidence 1.0)
 *   2. Levenshtein fuzzy match on full key (threshold 0.75)
 *   3. Word-overlap match — shared distinctive words like brand names (threshold 0.5)
 *
 * Returns the official name if a match is found, otherwise returns the original name
 */
export function matchSupplier(
  extractedName: string,
  supplierMap: DynamicSupplierMap,
  threshold: number = 0.75
): { matchedName: string; wasMatched: boolean; confidence: number; matchedAccount?: string } {
  if (!extractedName || Object.keys(supplierMap).length === 0) {
    return { matchedName: extractedName, wasMatched: false, confidence: 0 };
  }

  const extractedKey = generateSkeletonKey(extractedName);
  if (!extractedKey) {
    return { matchedName: extractedName, wasMatched: false, confidence: 0 };
  }

  // Strategy 1: Exact skeleton key match
  if (supplierMap[extractedKey]) {
    console.log(`Exact match: "${extractedName}" -> "${supplierMap[extractedKey].officialName}"`);
    return {
      matchedName: supplierMap[extractedKey].officialName,
      wasMatched: true,
      confidence: 1,
      matchedAccount: supplierMap[extractedKey].supplierAccount,
    };
  }

  // Strategy 2 & 3: Fuzzy match (Levenshtein) + Word-overlap match
  let bestMatch: SupplierEntry | null = null;
  let bestScore = 0;
  let matchStrategy = '';

  for (const [key, entry] of Object.entries(supplierMap)) {
    // Levenshtein on full skeleton key
    const levenshteinScore = calculateSimilarity(extractedKey, key);
    if (levenshteinScore > bestScore && levenshteinScore >= threshold) {
      bestScore = levenshteinScore;
      bestMatch = entry;
      matchStrategy = 'levenshtein';
    }

    // Word-overlap: catches cases where names use different languages
    // but share distinctive words (e.g., "omeroglu" or "sanita")
    const overlapScore = calculateWordOverlap(extractedKey, key);
    // Apply 0.85 confidence cap for word-overlap matches
    const adjustedOverlap = Math.min(overlapScore * 0.85, 0.95);
    if (adjustedOverlap > bestScore && overlapScore >= 0.5) {
      bestScore = adjustedOverlap;
      bestMatch = entry;
      matchStrategy = 'word-overlap';
    }
  }

  if (bestMatch) {
    console.log(`${matchStrategy} matched "${extractedName}" -> "${bestMatch.officialName}" (confidence: ${bestScore.toFixed(2)})`);
    return {
      matchedName: bestMatch.officialName,
      wasMatched: true,
      confidence: bestScore,
      matchedAccount: bestMatch.supplierAccount,
    };
  }

  // No match found - this is a NEW supplier
  console.log(`NEW supplier detected: "${extractedName}"`);
  return { matchedName: extractedName, wasMatched: false, confidence: 0 };
}
