// =============================================================================
// LOCAL PROCESSING: processWithOpenAI.ts
// Version: FINAL PRODUCTION v4.0 (AGGRESSIVE VALIDATION)
// =============================================================================

const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;
const MAX_INPUT_LENGTH = 15000;
const MODEL_TEXT = "gpt-4o-mini";
const MODEL_IMAGE = "gpt-4o";

// =============================================================================
// 1. VALID MEASURES WHITELIST
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
// 3. AGGRESSIVE SANITIZE
// =============================================================================
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined) return defaultVal;

  const str = String(val).trim();
  const lower = str.toLowerCase();

  if (GARBAGE_VALUES.some(g => lower === g || lower.includes(g))) {
    return defaultVal;
  }

  if (str === "" || str === "null" || str === "undefined") {
    return defaultVal;
  }

  return str;
}

// =============================================================================
// 4. VALIDATE MEASURE
// =============================================================================
function validateMeasure(measure: string): string {
  if (!measure) return "National Regulation";

  const lower = measure.toLowerCase();

  for (const valid of VALID_MEASURES) {
    if (lower.includes(valid.toLowerCase()) || valid.toLowerCase().includes(lower)) {
      return valid;
    }
  }

  if (measure.length > 50) {
    console.warn("Measure rejected (too long):", measure.substring(0, 50) + "...");
    return "National Regulation";
  }

  const suspiciousWords = ["manufacturing", "selling", "production", "products", "services", "company", "limited"];
  if (suspiciousWords.some(w => lower.includes(w))) {
    console.warn("Measure rejected (looks like description):", measure);
    return "National Regulation";
  }

  return measure;
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
// 6. EXTRACT DATE FROM TEXT
// =============================================================================
function extractDateFromText(text: string): string | null {
  const euroPattern = /(\d{1,2})[./-](\d{1,2})[./-](20\d{2})/g;
  const dates: string[] = [];

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

  return dates.length > 0 ? dates[dates.length - 1] : null;
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
// 8. DETECT CERTIFICATE TYPE FROM RAW TEXT
// =============================================================================
function detectCertTypeFromText(fullText: string): { scope: string; measure: string; certification: string } | null {
  const t = fullText.toLowerCase();

  // GLOVES
  if (t.includes("en 455") || t.includes("en 374") || t.includes("en 420") || t.includes("en 388") ||
      t.includes("module b") || t.includes("cat iii") || t.includes("category iii") ||
      t.includes("2016/425") || t.includes("nitrile glove") || t.includes("examination glove")) {
    const certName = t.includes("en 455") ? "EN 455" : t.includes("en 374") ? "EN 374" : t.includes("en 420") ? "EN 420" : "EN 388";
    return { scope: "+", measure: "EU Regulation 2016/425", certification: certName };
  }

  // ISO 14001
  if (t.includes("iso 14001") || t.includes("14001:")) {
    return { scope: "!", measure: "EU Waste Framework Directive (2008/98/EC)", certification: "ISO 14001" };
  }

  // ISO 45001
  if (t.includes("iso 45001") || t.includes("45001:")) {
    return { scope: "!", measure: "EU Directive 89/391/EEC", certification: "ISO 45001" };
  }

  // ISO 27001
  if (t.includes("iso 27001") || t.includes("27001:")) {
    return { scope: "!", measure: "EU GDPR", certification: "ISO 27001" };
  }

  // FSC
  if (t.match(/\bfsc\b/) && !t.includes("fssc")) {
    return { scope: "!", measure: "FSC", certification: "FSC" };
  }

  // ISO 9001 / BRC / BRCGS / ISO 22000 / FSSC / GMP
  if (t.includes("iso 9001") || t.includes("9001:") || t.includes("brcgs") || t.includes("brc global") ||
      t.includes("iso 22000") || t.includes("22000:") || t.includes("fssc 22000") || t.includes("fssc22000") ||
      t.match(/\bgmp\b/) || t.includes("good manufacturing")) {
    const certName = t.includes("brc") ? "BRC" : t.includes("fssc") ? "FSSC 22000" : t.includes("22000") ? "ISO 22000" : "ISO 9001";
    return { scope: "!", measure: "(EC) No 2023/2006", certification: certName };
  }

  // EU 10/2011
  if (t.includes("10/2011") || t.includes("eu 10/2011")) {
    return { scope: "+", measure: "(EC) No 10/2011", certification: "EU 10/2011" };
  }

  // EN 13432
  if (t.includes("en 13432") || t.includes("13432") || t.includes("compostable") || t.includes("din certco")) {
    return { scope: "+", measure: "EN 13432", certification: "EN 13432" };
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

  // STEP 3: Date fallback
  if (!data.date_issued || data.date_issued === "string" || data.date_issued === "") {
    const extractedDate = extractDateFromText(fullText);
    if (extractedDate) data.date_issued = extractedDate;
  }

  // STEP 4: 3-Year Rule
  if (data.date_issued && (!data.date_expired || data.date_expired === "string" || data.date_expired === "")) {
    if (data.measure === "EU Regulation 2016/425" || data.measure === "(EC) No 1935/2004") {
      data.date_expired = calculateExpiry(data.date_issued);
    }
  }

  // STEP 5: Validate measure
  data.measure = validateMeasure(data.measure);

  // STEP 6: Certificate number for dedup
  if (!data.certificate_number || data.certificate_number === "string" || data.certificate_number === "") {
    data.certificate_number = data.certification || "Certificate";
  }

  // STEP 7: Default scope
  if (!data.scope || data.scope === "string") {
    data.scope = "!";
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
// 11. SYSTEM PROMPT
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
