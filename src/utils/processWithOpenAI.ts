// =============================================================================
// LOCAL PROCESSING: processWithOpenAI.ts
// Version: FINAL PRODUCTION v4.0 (AGGRESSIVE VALIDATION)
// =============================================================================

const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;
const MAX_INPUT_LENGTH = 15000;
const MODEL_TEXT = "gpt-4o-mini";
const MODEL_IMAGE = "gpt-4o";

// =============================================================================
// 1. VALID MEASURES WHITELIST (From Compliance Knowledge Base)
// =============================================================================
const VALID_MEASURES = [
  // Food Contact Materials (FCM)
  "(EC) No 1935/2004",         // Framework regulation for all FCM
  "(EC) No 2023/2006",         // GMP for FCM manufacturing
  "(EC) No 10/2011",           // Plastic materials specific
  // PPE and Medical Devices
  "EU Regulation 2016/425",    // PPE Regulation (Gloves)
  "EU MDR 2017/745",           // Medical Device Regulation
  // Environmental & Safety
  "EU Waste Framework Directive (2008/98/EC)",  // ISO 14001
  "EU Directive 89/391/EEC",   // ISO 45001 Occupational Safety
  "EU GDPR",                   // ISO 27001 Information Security
  // Certification Standards
  "FSC",                       // Forest Stewardship Council
  "EN 13432",                  // Compostable packaging
  "EN 13430",                  // Recyclable packaging
  "EN 14287",                  // Aluminium/Foil for food contact
  // Measuring Instruments
  "EU Directive 2014/32/EU",   // Measuring Instruments Directive (MID)
  // Fallback
  "National Regulation",
];

// =============================================================================
// 1b. CERTIFICATION TO MEASURE MAPPING (From Compliance Knowledge Base)
// =============================================================================
const CERT_MEASURE_MAP: Record<string, { measure: string; scope: string }> = {
  // GMP & Quality Management -> (EC) No 2023/2006 (General !)
  "iso 9001": { measure: "(EC) No 2023/2006", scope: "!" },
  "iso 22000": { measure: "(EC) No 2023/2006", scope: "!" },
  "fssc 22000": { measure: "(EC) No 2023/2006", scope: "!" },
  "fssc22000": { measure: "(EC) No 2023/2006", scope: "!" },
  "brc": { measure: "(EC) No 2023/2006", scope: "!" },
  "brcgs": { measure: "(EC) No 2023/2006", scope: "!" },
  "haccp": { measure: "(EC) No 2023/2006", scope: "!" },
  "gmp": { measure: "(EC) No 2023/2006", scope: "!" },
  "ifs": { measure: "(EC) No 2023/2006", scope: "!" },
  // Environmental -> EU Waste Framework (General !)
  "iso 14001": { measure: "EU Waste Framework Directive (2008/98/EC)", scope: "!" },
  // Occupational Safety -> EU Directive 89/391/EEC (General !)
  "iso 45001": { measure: "EU Directive 89/391/EEC", scope: "!" },
  "ohsas 18001": { measure: "EU Directive 89/391/EEC", scope: "!" },
  // Information Security -> EU GDPR (General !)
  "iso 27001": { measure: "EU GDPR", scope: "!" },
  // Forest Certification (General !)
  "fsc": { measure: "FSC", scope: "!" },
  "pefc": { measure: "FSC", scope: "!" },
  // Gloves PPE -> EU Regulation 2016/425 (Specific +)
  "en 455": { measure: "EU Regulation 2016/425", scope: "+" },
  "en 374": { measure: "EU Regulation 2016/425", scope: "+" },
  "en 420": { measure: "EU Regulation 2016/425", scope: "+" },
  "en 388": { measure: "EU Regulation 2016/425", scope: "+" },
  "en 16523": { measure: "EU Regulation 2016/425", scope: "+" },
  "en iso 374": { measure: "EU Regulation 2016/425", scope: "+" },
  // Compostable -> EN 13432 (Specific +)
  "en 13432": { measure: "EN 13432", scope: "+" },
  "din certco": { measure: "EN 13432", scope: "+" },
  "ok compost": { measure: "EN 13432", scope: "+" },
  "compostable": { measure: "EN 13432", scope: "+" },
  // Recyclable -> EN 13430 (Specific +)
  "en 13430": { measure: "EN 13430", scope: "+" },
  "recyclass": { measure: "EN 13430", scope: "+" },
  "recyclable": { measure: "EN 13430", scope: "+" },
  // Plastics -> (EC) No 10/2011 (Specific +)
  "eu 10/2011": { measure: "(EC) No 10/2011", scope: "+" },
  "10/2011": { measure: "(EC) No 10/2011", scope: "+" },
  // Aluminium -> EN 14287 (Specific +)
  "en 14287": { measure: "EN 14287", scope: "+" },
  "en14287": { measure: "EN 14287", scope: "+" },
  // Measuring Instruments -> EU Directive 2014/32/EU (Specific +)
  "2014/32": { measure: "EU Directive 2014/32/EU", scope: "+" },
  "mid": { measure: "EU Directive 2014/32/EU", scope: "+" },
  "measuring instruments": { measure: "EU Directive 2014/32/EU", scope: "+" },
  // Food Contact General -> (EC) No 1935/2004 (Specific +)
  "1935/2004": { measure: "(EC) No 1935/2004", scope: "+" },
  "declaration of conformity": { measure: "(EC) No 1935/2004", scope: "+" },
  "declaration of compliance": { measure: "(EC) No 1935/2004", scope: "+" },
  "doc": { measure: "(EC) No 1935/2004", scope: "+" },
  "coc": { measure: "(EC) No 1935/2004", scope: "+" },
  "migration": { measure: "(EC) No 1935/2004", scope: "+" },
  "food contact": { measure: "(EC) No 1935/2004", scope: "+" },
  "food grade": { measure: "(EC) No 1935/2004", scope: "+" },
  // Medical Device -> EU MDR 2017/745 (Specific +)
  "2017/745": { measure: "EU MDR 2017/745", scope: "+" },
  "mdr": { measure: "EU MDR 2017/745", scope: "+" },
  "medical device": { measure: "EU MDR 2017/745", scope: "+" },
};

// =============================================================================
// 2. GARBAGE VALUES BLACKLIST
// =============================================================================
const GARBAGE_VALUES = [
  "string",
  "null",
  "undefined",
  "unknown",
  "n/a",
  "none",
  "not applicable",
  "not available",
  "see certificate",
  "as per certificate",
  "refer to",
];

// =============================================================================
// 3. AGGRESSIVE SANITIZE
// =============================================================================
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined) return defaultVal;

  const str = String(val).trim();
  const lower = str.toLowerCase();

  // CRITICAL: Explicit "string" check (JSON schema leak)
  if (lower === "string" || str === "String") {
    console.warn("BLOCKED: JSON schema leak detected - 'string' value");
    return defaultVal;
  }

  if (GARBAGE_VALUES.some(g => lower === g)) {
    console.warn("BLOCKED: Garbage value detected -", str);
    return defaultVal;
  }

  if (str === "" || str === "null" || str === "undefined") {
    return defaultVal;
  }

  return str;
}

// =============================================================================
// 4. VALIDATE MEASURE (STRICT - whitelist only)
// =============================================================================
function validateMeasure(measure: string, certification?: string): string {
  if (!measure || measure.toLowerCase() === "string") return inferMeasureFromCert(certification);

  const lower = measure.toLowerCase();

  for (const valid of VALID_MEASURES) {
    if (lower.includes(valid.toLowerCase()) || valid.toLowerCase().includes(lower)) {
      return valid;
    }
  }

  if (measure.length > 40) {
    console.warn("Measure rejected (too long) - forcing lookup:", measure.substring(0, 40) + "...");
    return inferMeasureFromCert(certification);
  }

  const suspiciousWords = ["manufacturing", "selling", "production", "products", "services", "company", "limited", "trading", "industry"];
  if (suspiciousWords.some(w => lower.includes(w))) {
    console.warn("Measure rejected (description) - forcing lookup:", measure);
    return inferMeasureFromCert(certification);
  }

  console.warn("Measure not in whitelist - forcing lookup:", measure);
  return inferMeasureFromCert(certification);
}

// =============================================================================
// 4b. INFER MEASURE FROM CERTIFICATION (Uses CERT_MEASURE_MAP)
// =============================================================================
function inferMeasureFromCert(certification?: string): string {
  if (!certification) return "National Regulation";

  const cert = certification.toLowerCase();

  // Check against comprehensive mapping (priority order matters)
  // 1. Specific ISO standards first (45001, 14001, 27001 before generic 9001)
  if (cert.includes("45001")) return "EU Directive 89/391/EEC";
  if (cert.includes("14001")) return "EU Waste Framework Directive (2008/98/EC)";
  if (cert.includes("27001")) return "EU GDPR";

  // 2. Check all keys in CERT_MEASURE_MAP
  for (const [key, value] of Object.entries(CERT_MEASURE_MAP)) {
    if (cert.includes(key)) {
      return value.measure;
    }
  }

  // 3. FSC (careful not to match FSSC)
  if (cert.match(/\bfsc\b/) && !cert.includes("fssc")) {
    return "FSC";
  }

  // 4. Generic quality management certs
  if (cert.includes("9001") || cert.includes("22000") || cert.includes("brc") || cert.includes("haccp")) {
    return "(EC) No 2023/2006";
  }

  // 5. Analysis Report / Test Report without specific standard -> Food Contact
  if (cert.includes("analysis report") || cert.includes("test report") || cert.includes("migration")) {
    return "(EC) No 1935/2004";
  }

  return "National Regulation";
}

// =============================================================================
// 5. INFER COUNTRY
// =============================================================================
function inferCountry(fullText: string, supplierName: string): string | null {
  const combined = (fullText + " " + supplierName).toLowerCase();

  if (combined.includes("china") || combined.includes("anhui") || combined.includes("guangdong") ||
      combined.includes("shanghai") || combined.includes("beijing") || combined.includes("shenzhen") ||
      combined.includes("changsha") || combined.includes("zhejiang") || combined.includes("jiangsu") ||
      combined.includes("hunan") || combined.includes("wenzhou") || combined.includes("fujian") ||
      combined.includes("shaoneng") || combined.includes("intco") || combined.match(/[\u4e00-\u9fff]/)) {
    return "China";
  }

  if (combined.includes("turkey") || combined.includes("türkiye") || combined.includes("turkiye") ||
      combined.includes("istanbul") || combined.includes("ankara") || combined.includes("izmir") ||
      combined.includes("boran") || combined.includes("mopack") || combined.includes("ilke") ||
      combined.includes("ambalaj") || combined.includes("san. ve tic")) {
    return "Turkey";
  }

  if (combined.includes("germany") || combined.includes("deutschland") || combined.includes("gmbh")) {
    return "Germany";
  }

  if (combined.includes("united kingdom") || combined.includes("england") || combined.includes("london") ||
      combined.match(/\buk\b/)) {
    return "UK";
  }

  if (combined.includes("ireland") || combined.includes("dublin") || combined.includes("cork")) {
    return "Ireland";
  }

  if (combined.includes("netherlands") || combined.includes("holland") || combined.includes("amsterdam")) {
    return "Netherlands";
  }

  if (combined.includes("italy") || combined.includes("italia") || combined.includes("s.r.l")) {
    return "Italy";
  }

  if (combined.includes("france") || combined.includes("paris")) {
    return "France";
  }

  if (combined.includes("poland") || combined.includes("polska")) {
    return "Poland";
  }

  if (combined.includes("spain") || combined.includes("españa")) {
    return "Spain";
  }

  return null;
}

// =============================================================================
// 6. DESPERATE DATE SCAVENGER (Multi-pattern extraction)
// =============================================================================
function extractDateFromText(text: string): string | null {
  const dates: string[] = [];

  // PATTERN 1: European format DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY
  const euroPattern = /(\d{1,2})[./-](\d{1,2})[./-](20\d{2})/g;
  let match: RegExpExecArray | null;
  while ((match = euroPattern.exec(text)) !== null) {
    const before = text.substring(Math.max(0, match.index - 5), match.index);
    if (before.match(/\d+$/) || before.match(/[:/]$/)) continue;

    const day = match[1].padStart(2, "0");
    const month = match[2].padStart(2, "0");
    const year = match[3];

    const monthNum = parseInt(month, 10);
    if (monthNum >= 1 && monthNum <= 12) {
      dates.push(`${year}-${month}-${day}`);
    }
  }

  // PATTERN 2: ISO format YYYY-MM-DD
  const isoPattern = /(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})/g;
  while ((match = isoPattern.exec(text)) !== null) {
    const year = match[1];
    const month = match[2].padStart(2, "0");
    const day = match[3].padStart(2, "0");

    const monthNum = parseInt(month, 10);
    const dayNum = parseInt(day, 10);
    if (monthNum >= 1 && monthNum <= 12 && dayNum >= 1 && dayNum <= 31) {
      dates.push(`${year}-${month}-${day}`);
    }
  }

  // PATTERN 3: Month name format (January 15, 2024 or 15 January 2024)
  const monthNames = "(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:tember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)";
  const monthNamePattern = new RegExp(`(\\d{1,2})[\\s.,/-]*(${monthNames})[\\s.,/-]*(20\\d{2})|(${monthNames})[\\s.,/-]*(\\d{1,2})[\\s.,/-]*(20\\d{2})`, "gi");
  while ((match = monthNamePattern.exec(text)) !== null) {
    let day: string, monthStr: string, year: string;
    if (match[1]) {
      // Format: 15 January 2024
      day = match[1];
      monthStr = match[2];
      year = match[3];
    } else {
      // Format: January 15, 2024
      monthStr = match[4];
      day = match[5];
      year = match[6];
    }

    const monthMap: { [key: string]: string } = {
      jan: "01", january: "01", feb: "02", february: "02", mar: "03", march: "03",
      apr: "04", april: "04", may: "05", jun: "06", june: "06", jul: "07", july: "07",
      aug: "08", august: "08", sep: "09", september: "09", oct: "10", october: "10",
      nov: "11", november: "11", dec: "12", december: "12"
    };
    const month = monthMap[monthStr.toLowerCase()];
    if (month) {
      dates.push(`${year}-${month}-${day.padStart(2, "0")}`);
    }
  }

  // PATTERN 4: Year-only as last resort (use Jan 1st)
  if (dates.length === 0) {
    const yearOnlyPattern = /\b(20[12]\d)\b/g;
    while ((match = yearOnlyPattern.exec(text)) !== null) {
      dates.push(`${match[1]}-01-01`);
    }
  }

  // Return FIRST valid date (most likely to be issue date)
  return dates.length > 0 ? dates[0] : null;
}

// =============================================================================
// 7. CALCULATE EXPIRY
// =============================================================================
function calculateExpiry(dateIssued: string): string {
  if (!dateIssued || !dateIssued.match(/^\d{4}-\d{2}-\d{2}$/)) return "";
  try {
    const parts = dateIssued.split("-");
    const year = parseInt(parts[0], 10) + 3;
    return `${year}-${parts[1]}-${parts[2]}`;
  } catch {
    return "";
  }
}

// =============================================================================
// 7b. EXTRACT YEAR FROM CERTIFICATION STRING
// Handles: "EN 455-1:2020", "ISO 9001:2015", "EN 455-1:2020+A1:2022"
// =============================================================================
function extractYearFromCertification(certification: string): string | null {
  if (!certification) return null;

  // Match year patterns like :2020, :2015, +A1:2022, etc.
  const yearPattern = /[:\-+](20[12]\d)\b/g;
  const years: number[] = [];

  let match: RegExpExecArray | null;
  while ((match = yearPattern.exec(certification)) !== null) {
    years.push(parseInt(match[1], 10));
  }

  // Return the LATEST year found (most recent revision)
  if (years.length > 0) {
    const latestYear = Math.max(...years);
    return `${latestYear}-01-01`;  // Use Jan 1st of that year
  }

  return null;
}

// =============================================================================
// 7c. GET TODAY'S DATE (for DoC default)
// =============================================================================
function getTodayISO(): string {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, "0");
  const day = String(today.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

// =============================================================================
// 8. DETECT CERTIFICATE TYPE FROM RAW TEXT
// =============================================================================
function detectCertTypeFromText(fullText: string): { scope: string; measure: string; certification: string } | null {
  const t = fullText.toLowerCase();

  // PRIORITY 1: GLOVES
  // First check for EU MDR 2017/745 (Medical Device Regulation - higher priority for medical gloves)
  if (t.includes("2017/745") || t.includes("mdr") || t.includes("medical device regulation")) {
    const certName = t.includes("en 455") ? "EN 455" : "Medical Device";
    return { scope: "+", measure: "EU MDR 2017/745", certification: certName };
  }

  // PPE Regulation (EU) 2016/425 for non-medical gloves
  if (t.includes("en 455") || t.includes("en 374") || t.includes("en 420") || t.includes("en 388") ||
      t.includes("module b") || t.includes("cat iii") || t.includes("category iii") ||
      t.includes("2016/425") || t.includes("nitrile glove") || t.includes("examination glove")) {
    const certName = t.includes("en 455") ? "EN 455" : t.includes("en 374") ? "EN 374" : t.includes("en 420") ? "EN 420" : "EN 388";
    return { scope: "+", measure: "EU Regulation 2016/425", certification: certName };
  }

  // PRIORITY 2: ISO 45001 (MUST CHECK BEFORE FACTORY CERTS!)
  if (t.includes("iso 45001") || t.includes("45001:") || t.match(/\b45001\b/)) {
    console.log("DETECTED: ISO 45001 - Mapping to EU Directive 89/391/EEC");
    return { scope: "!", measure: "EU Directive 89/391/EEC", certification: "ISO 45001" };
  }

  // PRIORITY 3: ISO 14001
  if (t.includes("iso 14001") || t.includes("14001:") || t.match(/\b14001\b/)) {
    return { scope: "!", measure: "EU Waste Framework Directive (2008/98/EC)", certification: "ISO 14001" };
  }

  // PRIORITY 4: ISO 27001
  if (t.includes("iso 27001") || t.includes("27001:") || t.match(/\b27001\b/)) {
    return { scope: "!", measure: "EU GDPR", certification: "ISO 27001" };
  }

  // PRIORITY 5: FSC (NOT FSSC)
  if (t.match(/\bfsc\b/) && !t.includes("fssc")) {
    return { scope: "!", measure: "FSC", certification: "FSC" };
  }

  // PRIORITY 6: ISO 9001 / BRC / BRCGS / ISO 22000 / FSSC / GMP (AFTER specific ISO checks)
  if (t.includes("iso 9001") || t.match(/\b9001:/) || t.includes("brcgs") || t.includes("brc global") ||
      t.includes("iso 22000") || t.match(/\b22000:/) || t.includes("fssc 22000") || t.includes("fssc22000") ||
      t.match(/\bgmp\b/) || t.includes("good manufacturing")) {
    const certName = t.includes("brc") ? "BRC" : t.includes("fssc") ? "FSSC 22000" : t.includes("22000") ? "ISO 22000" : "ISO 9001";
    return { scope: "!", measure: "(EC) No 2023/2006", certification: certName };
  }

  // EU 10/2011
  if (t.includes("10/2011") || t.includes("eu 10/2011")) {
    return { scope: "+", measure: "(EC) No 10/2011", certification: "EU 10/2011" };
  }

  // =========================================================================
  // PRIORITY 7: EN 13432 Compostable (MUST check BEFORE DoC/Migration!)
  // =========================================================================
  if (t.includes("en 13432") || t.includes("13432") || t.includes("compostable")) {
    return { scope: "+", measure: "EN 13432", certification: "EN 13432" };
  }

  // =========================================================================
  // PRIORITY 8: EN 13430 Recyclable (MUST check BEFORE DoC/Migration!)
  // =========================================================================
  if (t.includes("en 13430") || t.includes("13430") || t.includes("recyclable") || t.includes("recyclass")) {
    return { scope: "+", measure: "EN 13430", certification: "EN 13430" };
  }

  // =========================================================================
  // PRIORITY 9: DIN CERTCO (check AFTER specific EN standards)
  // =========================================================================
  if (t.includes("din certco") && !t.includes("13432") && !t.includes("13430")) {
    // DIN CERTCO without specific standard - default to compostable
    return { scope: "+", measure: "EN 13432", certification: "DIN CERTCO" };
  }

  // Declaration of Conformity / Migration
  if (t.includes("declaration of conformity") || t.includes("declaration of compliance") ||
      t.includes("declarative") || t.includes("conformity declaration") ||
      t.includes("migration") || t.includes("food contact") || t.includes("1935/2004")) {
    return { scope: "+", measure: "(EC) No 1935/2004", certification: "Declaration of Conformity" };
  }

  // Business License
  if (t.includes("business license") || t.includes("operating permit") || t.includes("trade license")) {
    return { scope: "!", measure: "National Regulation", certification: "Business License" };
  }

  return null;
}

// =============================================================================
// 9. APPLY BUSINESS LOGIC (AGGRESSIVE)
// =============================================================================
function applyBusinessLogic(data: any, fullText: string): any {
  // STEP 1: Detect from raw text FIRST
  const detected = detectCertTypeFromText(fullText);

  if (detected) {
    data.scope = detected.scope;
    data.measure = detected.measure;

    if (!data.certification || GARBAGE_VALUES.some(g => data.certification.toLowerCase().includes(g))) {
      data.certification = detected.certification;
    }

    if (!data.certificate_number || data.certificate_number === "string" || data.certificate_number === "") {
      data.certificate_number = detected.certification;
    }

    if (detected.measure === "EU Regulation 2016/425") {
      if (!data.product_category || data.product_category === "string" || data.product_category === "Goods") {
        data.product_category = "Gloves";
      }
    }
  }

  // STEP 2: Country inference
  if (!data.country || data.country === "Unknown" || data.country === "string") {
    const inferredCountry = inferCountry(fullText, data.supplier_name || "");
    if (inferredCountry) data.country = inferredCountry;
  }

  // STEP 3: Date fallback (multi-source)
  if (!data.date_issued || data.date_issued === "string" || data.date_issued === "") {
    // 3a: Try extracting from fullText first
    let extractedDate = extractDateFromText(fullText);

    // 3b: If no date from text, try certification field (e.g., "EN 455-1:2020")
    if (!extractedDate && data.certification) {
      extractedDate = extractYearFromCertification(data.certification);
      if (extractedDate) {
        console.log("Date extracted from certification string:", extractedDate);
      }
    }

    // 3c: If still no date and it's a DoC/Declaration, use today's date
    if (!extractedDate) {
      const certLower = (data.certification || "").toLowerCase();
      const measureLower = (data.measure || "").toLowerCase();
      const isDoC = certLower.includes("declaration") || certLower.includes("conformity") ||
                    certLower.includes("doc") || measureLower.includes("1935/2004") ||
                    certLower.includes("en 455") || certLower.includes("en 374");

      if (isDoC) {
        extractedDate = getTodayISO();
        console.log("DoC detected without date - using today's date:", extractedDate);
      }
    }

    if (extractedDate) data.date_issued = extractedDate;
  }

  // STEP 4: 3-Year Rule
  if (data.date_issued && (!data.date_expired || data.date_expired === "string" || data.date_expired === "")) {
    if (data.measure === "EU Regulation 2016/425" || data.measure === "(EC) No 1935/2004") {
      data.date_expired = calculateExpiry(data.date_issued);
    }
  }

  // =========================================================================
  // STEP 5: CERTIFICATION-BASED MEASURE OVERRIDE (Uses CERT_MEASURE_MAP)
  // This runs BEFORE whitelist validation to force correct mappings
  // =========================================================================
  const certLower = (data.certification || "").toLowerCase();

  // Priority 1: Specific ISO standards (must check before generic patterns)
  if (certLower.includes("45001")) {
    console.log("OVERRIDE: ISO 45001 -> EU Directive 89/391/EEC");
    data.measure = "EU Directive 89/391/EEC";
    data.scope = "!";
  }
  else if (certLower.includes("14001")) {
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    data.scope = "!";
  }
  else if (certLower.includes("27001")) {
    data.measure = "EU GDPR";
    data.scope = "!";
  }
  // Priority 2: FSC (careful not to match FSSC)
  else if (certLower.match(/\bfsc\b/) && !certLower.includes("fssc")) {
    data.measure = "FSC";
    data.scope = "!";
  }
  // Priority 3: Use CERT_MEASURE_MAP for all other patterns
  else {
    let matched = false;
    for (const [key, value] of Object.entries(CERT_MEASURE_MAP)) {
      if (certLower.includes(key)) {
        console.log(`OVERRIDE: "${key}" matched -> ${value.measure} (${value.scope})`);
        data.measure = value.measure;
        data.scope = value.scope;
        matched = true;
        break;
      }
    }
    // Priority 4: Generic quality certs -> GMP
    if (!matched && (certLower.includes("9001") || certLower.includes("22000") ||
        certLower.includes("brc") || certLower.includes("haccp"))) {
      data.measure = "(EC) No 2023/2006";
      data.scope = "!";
    }
  }

  // STEP 6: Validate measure (whitelist + cert-based fallback)
  data.measure = validateMeasure(data.measure, data.certification);

  // STEP 7: Keep certificate_number CLEAN (no artificial uniqueness - dedup handled in Excel export)
  // Client requirement: Return raw certificate number as extracted

  // STEP 8: Default scope
  if (!data.scope || data.scope === "string") {
    data.scope = "!";
  }

  // STEP 9: Apply 3-Year Rule if still no expiry (catches cases where measure changed)
  if (data.date_issued && (!data.date_expired || data.date_expired === "")) {
    data.date_expired = calculateExpiry(data.date_issued);
    console.log("Applied 3-Year Rule fallback: expiry =", data.date_expired);
  }

  return data;
}

// =============================================================================
// 10. TYPES
// =============================================================================
interface CertificateExtractionResult {
  supplier_name: string;
  certificate_number: string;
  country: string;
  scope: string;
  measure: string;
  certification: string;
  product_category: string;
  date_issued: string;
  date_expired: string;
  status?: string;
}

// =============================================================================
// 11. SYSTEM PROMPT (Compliance Knowledge Base Integrated)
// =============================================================================
const systemPrompt = `You are a Food Contact Materials (FCM) Compliance Data Extraction Engine.
Extract certificate data from documents in ANY language and return ONLY valid JSON.

=== EXTRACTION RULES ===
- supplier_name: Company name from "Applicant", "Manufacturer", "Holder", or letterhead
- certificate_number: The certificate/report/document number (e.g., "QC14335Q001", "LNE-38792")
- country: Country from address (China, Turkey, Serbia, Germany, UK, etc.)
- certification: Standard/certificate type (e.g., "ISO 9001:2015", "FSSC 22000", "Declaration of Compliance")
- product_category: Product description (e.g., "Plastic cups", "Paper packaging", "Gloves")
- date_issued: Issue date as YYYY-MM-DD (European format: "09.10.2024" = October 9th, NOT September 10th)
- date_expired: Expiry date as YYYY-MM-DD, or null if "validity until revocation" or not specified

=== SCOPE CLASSIFICATION (CRITICAL) ===
"!" = GENERAL MEASURE (Facility-wide, applies to all products from manufacturer)
"+" = SPECIFIC MEASURE (Product-specific, applies only to listed products)

=== MEASURE MAPPINGS (Use EXACTLY these values) ===
GENERAL MEASURES (!):
- ISO 9001, ISO 22000, FSSC 22000, BRC, BRCGS, HACCP, GMP, IFS → "(EC) No 2023/2006"
- ISO 14001 (Environmental) → "EU Waste Framework Directive (2008/98/EC)"
- ISO 45001, OHSAS 18001 (Safety) → "EU Directive 89/391/EEC"
- ISO 27001 (Information Security) → "EU GDPR"
- FSC, PEFC (Forest certification) → "FSC"

SPECIFIC MEASURES (+):
- Declaration of Conformity, DoC, CoC, Migration Report, Food Contact, 1935/2004 → "(EC) No 1935/2004"
- Plastic materials, EU 10/2011, Specific Migration → "(EC) No 10/2011"
- EN 455, EN 374, EN 420, EN 388 (Gloves PPE) → "EU Regulation 2016/425"
- EN 13432, DIN CERTCO, OK Compost (Compostable) → "EN 13432"
- EN 13430, RecyClass (Recyclable) → "EN 13430"
- EN 14287 (Aluminium/Foil) → "EN 14287"
- EU MDR 2017/745, Medical Device → "EU MDR 2017/745"
- EU 2014/32/EU, MID, Measuring Instruments → "EU Directive 2014/32/EU"

=== DATE FORMATS (Handle all) ===
- European: DD.MM.YYYY or DD/MM/YYYY (e.g., "09.10.2024" = October 9th)
- ISO: YYYY-MM-DD
- Text: "9 October 2024", "October 9, 2024"
- Serbian/European: "04.06.2025" = June 4th, 2025

=== CRITICAL RULES ===
1. NEVER output "string" as a value - extract actual data or use null
2. For "validity until revocation" documents, set date_expired to null
3. Look for expiry in: "Valid until", "Expiry", "Date of Expiry", "Gültig bis"
4. Look for issue in: "Issue date", "Date of issue", "Issued", "Ausstellungsdatum"
5. Certificate numbers often appear after "Certificate No:", "Report No:", "Cert. No."
6. LANGUAGE RULE: If the document contains multiple languages (e.g., Chinese on Page 1, English on Page 2), YOU MUST ONLY EXTRACT THE ENGLISH TEXT. Prioritize the English translation for Supplier Name, Product Category, and Certification Name. Never output Chinese characters.

Return JSON:
{
  "supplier_name": "DIVI d.o.o.",
  "certificate_number": "QC14335Q001",
  "country": "Serbia",
  "scope": "!",
  "measure": "(EC) No 2023/2006",
  "certification": "ISO 9001:2015",
  "product_category": "Plastic packaging",
  "date_issued": "2025-06-04",
  "date_expired": "2028-06-03"
}`;

// =============================================================================
// 12. MAIN EXPORT
// =============================================================================
export async function processWithOpenAI(
  base64Image: string,
  filename?: string,
  textContent?: string
): Promise<CertificateExtractionResult> {
  try {
    if (!OPENAI_API_KEY) {
      throw new Error("OpenAI API key not configured");
    }

    const isTextOnly = !!textContent && !base64Image;
    console.log(`Processing: ${filename || "unknown"} | Mode: ${isTextOnly ? "TEXT" : "IMAGE"}`);

    let userContent: Array<{ type: string; text?: string; image_url?: { url: string } }>;
    let truncatedText = "";

    if (isTextOnly && textContent) {
      truncatedText = textContent.length > MAX_INPUT_LENGTH
        ? textContent.slice(0, MAX_INPUT_LENGTH) + "\n[TRUNCATED]"
        : textContent;

      userContent = [
        { type: "text", text: `Extract certificate data from (${filename || "unknown"}):\n\n${truncatedText}` },
      ];
    } else {
      const imageUrl = base64Image.startsWith("data:") ? base64Image : `data:image/jpeg;base64,${base64Image}`;
      userContent = [
        { type: "image_url", image_url: { url: imageUrl } },
        { type: "text", text: `Extract certificate data from (${filename || "unknown"}).` },
      ];
    }

    const selectedModel = isTextOnly ? MODEL_TEXT : MODEL_IMAGE;
    console.log(`Using model: ${selectedModel}`);

    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: selectedModel,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userContent },
        ],
        max_tokens: 800,
        temperature: 0.1,
        response_format: { type: "json_object" },
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || "OpenAI API failed");
    }

    const data = await response.json();
    const rawContent = data.choices[0]?.message?.content;

    console.log("Raw AI Response:", rawContent?.substring(0, 300));

    if (!rawContent) {
      return {
        supplier_name: "Error: No AI Response",
        certificate_number: "Error",
        country: "Unknown",
        scope: "!",
        measure: "Manual Review Required",
        certification: "AI returned empty",
        product_category: "Unknown",
        date_issued: "",
        date_expired: "",
        status: "Error",
      };
    }

    let parsedData: any;
    try {
      parsedData = JSON.parse(rawContent);
    } catch {
      console.error("JSON Parse Error");
      parsedData = {};
    }

    // Apply AGGRESSIVE business logic
    parsedData = applyBusinessLogic(parsedData, truncatedText || textContent || "");

    // Final sanitization
    return {
      supplier_name: sanitize(parsedData.supplier_name, "Unknown Supplier"),
      certificate_number: sanitize(parsedData.certificate_number, "Certificate"),
      country: sanitize(parsedData.country, "Unknown"),
      scope: sanitize(parsedData.scope, "!"),
      measure: sanitize(parsedData.measure, "National Regulation"),
      certification: sanitize(parsedData.certification, "Certificate"),
      product_category: sanitize(parsedData.product_category, "Goods"),
      date_issued: sanitize(parsedData.date_issued, ""),
      date_expired: sanitize(parsedData.date_expired, ""),
      status: "Success",
    };

  } catch (error) {
    console.error("Processing Error:", error);
    return {
      supplier_name: "Error: Processing Failed",
      certificate_number: "Error",
      country: "Unknown",
      scope: "!",
      measure: "Manual Review Required",
      certification: "System Error",
      product_category: "Unknown",
      date_issued: "",
      date_expired: "",
      status: "Error",
    };
  }
}
