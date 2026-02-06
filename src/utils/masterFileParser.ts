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
 * Generate a "skeleton key" from a supplier name for fuzzy matching.
 *
 * Rules:
 * 1. Convert to lowercase
 * 2. Remove all punctuation
 * 3. Remove common company suffixes (Co, Ltd, Inc, GmbH, etc.)
 * 4. Remove extra whitespace
 * 5. Sort words alphabetically (for word-order independence)
 */
export function generateSkeletonKey(name: string): string {
  if (!name) return '';

  let key = name.toLowerCase();

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
    // Turkish
    'as', 'a.s', 'a.ş', 'aş', 'anonim', 'şirketi', 'sirketi', 'tic', 'ticaret', 'san', 'sanayi', 'ltd', 'sti', 'şti',
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

  // Get the first sheet (or 'Certificates' if it exists)
  let sheetName = workbook.SheetNames.find(name =>
    name.toLowerCase() === 'certificates' || name.toLowerCase().includes('certificate')
  );
  if (!sheetName) {
    sheetName = workbook.SheetNames[0];
  }

  if (!sheetName) {
    throw new Error('No worksheet found in Master File');
  }

  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new Error('Unable to read worksheet from Master File');
  }

  // Convert to JSON array for easier processing
  const data: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  if (data.length < 2) {
    throw new Error('Master File appears to be empty or has no data rows');
  }

  // Find header columns (row 0)
  const headerRow = data[0] || [];
  let supplierNameCol = -1;
  let countryCol = -1;
  let supplierAccountCol = -1;

  headerRow.forEach((cell: any, index: number) => {
    const value = String(cell || '').toLowerCase().trim();
    if (value.includes('supplier name') || value === 'supplier') {
      supplierNameCol = index;
    }
    if (value === 'country') {
      countryCol = index;
    }
    if (value.includes('supplier account') || value === 'account') {
      supplierAccountCol = index;
    }
  });

  // V7 Master File expected columns:
  // Col 0: Supplier Account, Col 1: Supplier Name, Col 2: Country (0-indexed)
  if (supplierNameCol === -1) {
    // Fallback to V7 defaults (0-indexed)
    supplierAccountCol = 0;
    supplierNameCol = 1;
    countryCol = 2;
    console.log('Using V7 default column positions');
  }

  console.log(`Master File columns - Supplier Name: ${supplierNameCol}, Country: ${countryCol}, Account: ${supplierAccountCol}`);

  // Iterate through data rows (skip header at index 0)
  for (let i = 1; i < data.length; i++) {
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
 * Find the best matching supplier from the DYNAMIC_SUPPLIER_MAP
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

  // Exact skeleton key match
  if (supplierMap[extractedKey]) {
    return {
      matchedName: supplierMap[extractedKey].officialName,
      wasMatched: true,
      confidence: 1,
      matchedAccount: supplierMap[extractedKey].supplierAccount,
    };
  }

  // Fuzzy matching
  let bestMatch: SupplierEntry | null = null;
  let bestScore = 0;

  for (const [key, entry] of Object.entries(supplierMap)) {
    const score = calculateSimilarity(extractedKey, key);
    if (score > bestScore && score >= threshold) {
      bestScore = score;
      bestMatch = entry;
    }
  }

  if (bestMatch) {
    console.log(`Fuzzy matched "${extractedName}" -> "${bestMatch.officialName}" (confidence: ${bestScore.toFixed(2)})`);
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
