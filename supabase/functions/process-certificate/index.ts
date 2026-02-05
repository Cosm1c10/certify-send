// =============================================================================
// SUPABASE EDGE FUNCTION: process-certificate
// Version: FINAL PRODUCTION v4.0 (AGGRESSIVE VALIDATION)
// =============================================================================
// deno-lint-ignore-file no-explicit-any
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";

declare const Deno: {
  env: {
    get(key: string): string | undefined;
  };
};

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
};

// =============================================================================
// 1. VALID MEASURES WHITELIST (Only these are allowed)
// =============================================================================
const VALID_MEASURES = [
  "EU Regulation 2016/425",
  "(EC) No 2023/2006",
  "(EC) No 1935/2004",
  "(EC) No 10/2011",
  "EU Waste Framework Directive (2008/98/EC)",
  "EU Directive 89/391/EEC",
  "EU GDPR",
  "FSC",
  "EN 13432",
  "EN 13430",
  "National Regulation",
];

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
// 3. AGGRESSIVE SANITIZE (Blocks garbage + validates)
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

  // Block garbage values
  if (GARBAGE_VALUES.some(g => lower === g)) {
    console.warn("BLOCKED: Garbage value detected -", str);
    return defaultVal;
  }

  // Block empty
  if (str === "" || str === "null" || str === "undefined") {
    return defaultVal;
  }

  return str;
}

// =============================================================================
// 4. VALIDATE MEASURE (Must be from whitelist - STRICT)
// =============================================================================
function validateMeasure(measure: string, certification?: string): string {
  if (!measure || measure.toLowerCase() === "string") return inferMeasureFromCert(certification);

  const lower = measure.toLowerCase();

  // Check if it's a valid measure (case-insensitive)
  for (const valid of VALID_MEASURES) {
    if (lower.includes(valid.toLowerCase()) || valid.toLowerCase().includes(lower)) {
      return valid; // Return the properly formatted version
    }
  }

  // If measure is too long (>40 chars), it's a description - FORCE LOOKUP
  if (measure.length > 40) {
    console.warn("Measure rejected (too long) - forcing lookup:", measure.substring(0, 40) + "...");
    return inferMeasureFromCert(certification);
  }

  // If it contains suspicious words, FORCE LOOKUP
  const suspiciousWords = ["manufacturing", "selling", "production", "products", "services", "company", "limited", "trading", "industry"];
  if (suspiciousWords.some(w => lower.includes(w))) {
    console.warn("Measure rejected (description) - forcing lookup:", measure);
    return inferMeasureFromCert(certification);
  }

  // Not in whitelist and not a description - still force lookup
  console.warn("Measure not in whitelist - forcing lookup:", measure);
  return inferMeasureFromCert(certification);
}

// =============================================================================
// 4b. INFER MEASURE FROM CERTIFICATION (Fallback)
// =============================================================================
function inferMeasureFromCert(certification?: string): string {
  if (!certification) return "National Regulation";

  const cert = certification.toLowerCase();

  // Gloves
  if (cert.includes("en 455") || cert.includes("en 374") || cert.includes("en 420") || cert.includes("en 388")) {
    return "EU Regulation 2016/425";
  }

  // ISO 45001 - SAFETY (must check BEFORE 9001)
  if (cert.includes("45001") || cert.includes("iso 45001")) {
    return "EU Directive 89/391/EEC";
  }

  // ISO 14001 - Environmental
  if (cert.includes("14001") || cert.includes("iso 14001")) {
    return "EU Waste Framework Directive (2008/98/EC)";
  }

  // ISO 27001 - Information Security
  if (cert.includes("27001") || cert.includes("iso 27001")) {
    return "EU GDPR";
  }

  // ISO 9001 / BRC / GMP
  if (cert.includes("9001") || cert.includes("brc") || cert.includes("22000") || cert.includes("fssc") || cert.includes("gmp")) {
    return "(EC) No 2023/2006";
  }

  // DoC / Migration
  if (cert.includes("declaration") || cert.includes("doc") || cert.includes("migration") || cert.includes("conformity")) {
    return "(EC) No 1935/2004";
  }

  // FSC
  if (cert.includes("fsc") && !cert.includes("fssc")) {
    return "FSC";
  }

  // EN 13432 Compostable
  if (cert.includes("13432") || cert.includes("compostable")) {
    return "EN 13432";
  }

  // EN 13430 Recyclable
  if (cert.includes("13430") || cert.includes("recyclable") || cert.includes("recyclass")) {
    return "EN 13430";
  }

  // EU 10/2011
  if (cert.includes("10/2011")) {
    return "(EC) No 10/2011";
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
// 8. DETECT CERTIFICATE TYPE FROM RAW TEXT (TEXT-FIRST APPROACH)
// =============================================================================
function detectCertTypeFromText(fullText: string): { scope: string; measure: string; certification: string } | null {
  const t = fullText.toLowerCase();

  // =========================================================================
  // PRIORITY 1: GLOVES (Highest Priority)
  // =========================================================================
  if (t.includes("en 455") || t.includes("en 374") || t.includes("en 420") || t.includes("en 388") ||
      t.includes("module b") || t.includes("cat iii") || t.includes("category iii") ||
      t.includes("2016/425") || t.includes("nitrile glove") || t.includes("examination glove")) {
    const certName = t.includes("en 455") ? "EN 455" : t.includes("en 374") ? "EN 374" : t.includes("en 420") ? "EN 420" : "EN 388";
    return { scope: "+", measure: "EU Regulation 2016/425", certification: certName };
  }

  // =========================================================================
  // PRIORITY 2: ISO 45001 (MUST CHECK BEFORE FACTORY CERTS!)
  // Occupational Health & Safety - NOT GMP!
  // =========================================================================
  if (t.includes("iso 45001") || t.includes("45001:") || t.match(/\b45001\b/)) {
    console.log("DETECTED: ISO 45001 - Mapping to EU Directive 89/391/EEC");
    return { scope: "!", measure: "EU Directive 89/391/EEC", certification: "ISO 45001" };
  }

  // =========================================================================
  // PRIORITY 3: ISO 14001 (Environmental)
  // =========================================================================
  if (t.includes("iso 14001") || t.includes("14001:") || t.match(/\b14001\b/)) {
    return { scope: "!", measure: "EU Waste Framework Directive (2008/98/EC)", certification: "ISO 14001" };
  }

  // =========================================================================
  // PRIORITY 4: ISO 27001 (Information Security)
  // =========================================================================
  if (t.includes("iso 27001") || t.includes("27001:") || t.match(/\b27001\b/)) {
    return { scope: "!", measure: "EU GDPR", certification: "ISO 27001" };
  }

  // =========================================================================
  // PRIORITY 5: FSC (Forestry - NOT FSSC!)
  // =========================================================================
  if (t.match(/\bfsc\b/) && !t.includes("fssc")) {
    return { scope: "!", measure: "FSC", certification: "FSC" };
  }

  // =========================================================================
  // PRIORITY 6: ISO 9001 / BRC / BRCGS / ISO 22000 / FSSC / GMP
  // Factory/Manufacturing certs - AFTER specific ISO checks
  // =========================================================================
  if (t.includes("iso 9001") || t.match(/\b9001:/) || t.includes("brcgs") || t.includes("brc global") ||
      t.includes("iso 22000") || t.match(/\b22000:/) || t.includes("fssc 22000") || t.includes("fssc22000") ||
      t.match(/\bgmp\b/) || t.includes("good manufacturing")) {
    const certName = t.includes("brc") ? "BRC" : t.includes("fssc") ? "FSSC 22000" : t.includes("22000") ? "ISO 22000" : "ISO 9001";
    return { scope: "!", measure: "(EC) No 2023/2006", certification: certName };
  }

  // EU 10/2011 Plastics
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
// 9. APPLY BUSINESS LOGIC (AGGRESSIVE - TEXT-FIRST)
// =============================================================================
function applyBusinessLogic(data: any, fullText: string): any {
  // STEP 1: Detect cert type from RAW TEXT first (most reliable)
  const detected = detectCertTypeFromText(fullText);

  if (detected) {
    // FORCE the values - don't trust AI
    data.scope = detected.scope;
    data.measure = detected.measure;

    // Only override certification if AI gave garbage
    if (!data.certification || GARBAGE_VALUES.some(g => data.certification.toLowerCase().includes(g))) {
      data.certification = detected.certification;
    }

    // Set certificate_number for dedup
    if (!data.certificate_number || data.certificate_number === "string" || data.certificate_number === "") {
      data.certificate_number = detected.certification;
    }

    // Gloves product category
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

  // STEP 3: Date fallback
  if (!data.date_issued || data.date_issued === "string" || data.date_issued === "") {
    const extractedDate = extractDateFromText(fullText);
    if (extractedDate) data.date_issued = extractedDate;
  }

  // STEP 4: 3-Year Rule for test reports
  if (data.date_issued && (!data.date_expired || data.date_expired === "string" || data.date_expired === "")) {
    if (data.measure === "EU Regulation 2016/425" || data.measure === "(EC) No 1935/2004") {
      data.date_expired = calculateExpiry(data.date_issued);
    }
  }

  // =========================================================================
  // STEP 5: CERTIFICATION-BASED MEASURE OVERRIDE (Catches text-first misses)
  // This runs BEFORE whitelist validation to force correct mappings
  // =========================================================================
  const certLower = (data.certification || "").toLowerCase();

  // ISO 45001 -> Safety (NOT GMP!)
  if (certLower.includes("45001") || certLower.includes("iso 45001")) {
    console.log("OVERRIDE: ISO 45001 detected in certification - forcing EU Directive 89/391/EEC");
    data.measure = "EU Directive 89/391/EEC";
    data.scope = "!";
  }
  // ISO 14001 -> Environmental
  else if (certLower.includes("14001") || certLower.includes("iso 14001")) {
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    data.scope = "!";
  }
  // ISO 27001 -> GDPR
  else if (certLower.includes("27001") || certLower.includes("iso 27001")) {
    data.measure = "EU GDPR";
    data.scope = "!";
  }
  // FSC (not FSSC!)
  else if (certLower.match(/\bfsc\b/) && !certLower.includes("fssc")) {
    data.measure = "FSC";
    data.scope = "!";
  }
  // EN 13432 Compostable
  else if (certLower.includes("13432") || certLower.includes("compostable")) {
    data.measure = "EN 13432";
    data.scope = "+";
  }
  // EN 13430 Recyclable
  else if (certLower.includes("13430") || certLower.includes("recyclable")) {
    data.measure = "EN 13430";
    data.scope = "+";
  }
  // Gloves
  else if (certLower.includes("en 455") || certLower.includes("en 374") || certLower.includes("en 420") || certLower.includes("en 388")) {
    data.measure = "EU Regulation 2016/425";
    data.scope = "+";
  }

  // STEP 6: Validate measure (must be from whitelist, fallback to cert-based lookup)
  data.measure = validateMeasure(data.measure, data.certification);

  // STEP 7: Ensure certificate_number is UNIQUE (append certification for dedup)
  // This prevents Excel from dropping rows with same cert number but different standards
  const certNum = sanitize(data.certificate_number, "");
  const certName = sanitize(data.certification, "Certificate");
  if (certNum && certNum !== certName && !certNum.includes("(")) {
    // Append certification name to make unique: "2690-2023-003529-W1 (EN 13430)"
    data.certificate_number = `${certNum} (${certName})`;
  } else if (!certNum) {
    data.certificate_number = certName;
  }

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
// 10. PROCESS AI RESPONSE
// =============================================================================
function processAIResponse(rawJSON: any, fullInput: string): any {
  let data = rawJSON || {};

  // Apply aggressive business logic
  data = applyBusinessLogic(data, fullInput);

  // FINAL SANITIZATION with defaults
  return {
    supplier_name: sanitize(data.supplier_name, "Unknown Supplier"),
    certificate_number: sanitize(data.certificate_number, "Certificate"),
    country: sanitize(data.country, "Unknown"),
    scope: sanitize(data.scope, "!"),
    measure: sanitize(data.measure, "National Regulation"),
    certification: sanitize(data.certification, "Certificate"),
    product_category: sanitize(data.product_category, "Goods"),
    date_issued: sanitize(data.date_issued, ""),
    date_expired: sanitize(data.date_expired, ""),
    status: "Success",
  };
}

// =============================================================================
// 11. SYSTEM PROMPT (No "string" types - use examples instead)
// =============================================================================
const systemPrompt = `You are a Compliance Data Extraction Engine. Extract certificate data and return ONLY valid JSON.

RULES:
- supplier_name: Company name from "Applicant", "Manufacturer", or letterhead
- certificate_number: The certificate/report number (e.g., "CERT-2024-001")
- country: Country from address (e.g., "China", "Turkey", "Germany")
- certification: Standard name (e.g., "ISO 9001", "EN 455", "BRC")
- product_category: Product description (e.g., "Gloves", "Paper Cups", "Packaging")
- date_issued: Issue date as YYYY-MM-DD (European: "09.10.2024" = October 9th)
- date_expired: Expiry date as YYYY-MM-DD, or null if not found

CLASSIFICATION:
- Gloves (EN 455/374/420): scope="+", measure="EU Regulation 2016/425"
- ISO 9001/BRC/GMP: scope="!", measure="(EC) No 2023/2006"
- DoC/Migration/1935: scope="+", measure="(EC) No 1935/2004"

IMPORTANT: Never output the word "string" as a value. Extract actual data or use null.

Return JSON:
{
  "supplier_name": "Actual Company Name",
  "certificate_number": "CERT-123 or null",
  "country": "Country Name or null",
  "scope": "!" or "+",
  "measure": "Regulation Name",
  "certification": "Standard Name",
  "product_category": "Product Description",
  "date_issued": "2024-01-15 or null",
  "date_expired": "2027-01-15 or null"
}`;

// =============================================================================
// 12. CONFIGURATION
// =============================================================================
const MAX_INPUT_LENGTH = 15000;
const OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";
const MODEL_TEXT = "gpt-4o-mini";
const MODEL_IMAGE = "gpt-4o";

// =============================================================================
// 13. MAIN HANDLER
// =============================================================================
serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const input = await req.json();

    const rawText = input.text || input.fileContent || "";
    const rawImage = input.image || "";
    const fileName = input.filename || input.fileName || "Unknown";

    const isTextMode = !!rawText && !rawImage;

    if (!rawText && !rawImage) {
      return new Response(
        JSON.stringify({ error: "Missing 'text', 'fileContent', or 'image'" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    console.log(`Processing: ${fileName} | Mode: ${isTextMode ? "TEXT" : "IMAGE"}`);

    const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY");
    if (!OPENAI_API_KEY) {
      return new Response(
        JSON.stringify({ error: "OPENAI_API_KEY not configured" }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    let userContent: any[];
    let truncatedText = "";

    if (isTextMode) {
      truncatedText = rawText.length > MAX_INPUT_LENGTH
        ? rawText.slice(0, MAX_INPUT_LENGTH) + "\n[TRUNCATED]"
        : rawText;

      userContent = [
        { type: "text", text: `Extract certificate data from (${fileName}):\n\n${truncatedText}` },
      ];
    } else {
      const imageUrl = rawImage.startsWith("data:") ? rawImage : `data:image/jpeg;base64,${rawImage}`;
      userContent = [
        { type: "image_url", image_url: { url: imageUrl } },
        { type: "text", text: `Extract certificate data from (${fileName}).` },
      ];
    }

    const selectedModel = isTextMode ? MODEL_TEXT : MODEL_IMAGE;
    console.log(`Using model: ${selectedModel}`);

    const response = await fetch(OPENAI_API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${OPENAI_API_KEY}`,
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
      const errorData = await response.json();
      console.error("OpenAI API Error:", errorData);
      throw new Error(errorData.error?.message || `OpenAI API Error: ${response.status}`);
    }

    const aiResponse = await response.json();
    const rawContent = aiResponse.choices?.[0]?.message?.content;

    console.log("Raw AI Response:", rawContent?.substring(0, 300));

    if (!rawContent) {
      return new Response(
        JSON.stringify({
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
        }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    let parsedData: any;
    try {
      parsedData = JSON.parse(rawContent);
    } catch {
      console.error("JSON Parse Error");
      parsedData = {};
    }

    // Process with AGGRESSIVE business logic
    const result = processAIResponse(parsedData, truncatedText || rawText);

    return new Response(JSON.stringify(result), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });

  } catch (error) {
    console.error("Processing Error:", error);

    return new Response(
      JSON.stringify({
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
      }),
      { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});
