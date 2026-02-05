// =============================================================================
// LOCAL PROCESSING: processWithOpenAI.ts
// Version: FINAL PRODUCTION v3.0 (Dedup Fix + Country Inference + Date Fallback)
// =============================================================================

const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;
const MAX_INPUT_LENGTH = 50000;

// =============================================================================
// 1. SANITIZE
// =============================================================================
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined || val === "null" || val === "undefined") {
    return defaultVal;
  }
  return String(val).trim();
}

// =============================================================================
// 2. INFER COUNTRY (Checks BOTH text AND supplier_name)
// =============================================================================
function inferCountry(fullText: string, supplierName: string): string | null {
  const combined = (fullText + " " + supplierName).toLowerCase();

  // CHINA
  if (
    combined.includes("china") || combined.includes("anhui") || combined.includes("guangdong") ||
    combined.includes("shanghai") || combined.includes("beijing") || combined.includes("shenzhen") ||
    combined.includes("changsha") || combined.includes("zhejiang") || combined.includes("jiangsu") ||
    combined.includes("hunan") || combined.includes("wenzhou") || combined.includes("fujian") ||
    combined.includes("shaoneng") || combined.includes("intco") ||
    combined.match(/[\u4e00-\u9fff]/) // Chinese characters
  ) {
    return "China";
  }

  // TURKEY
  if (
    combined.includes("turkey") || combined.includes("türkiye") || combined.includes("istanbul") ||
    combined.includes("ankara") || combined.includes("boran") || combined.includes("mopack") ||
    combined.includes("san. ve tic")
  ) {
    return "Turkey";
  }

  // GERMANY
  if (combined.includes("germany") || combined.includes("deutschland") || combined.includes("gmbh")) {
    return "Germany";
  }

  // UK
  if (combined.includes("united kingdom") || combined.includes("england") || combined.includes("london") || combined.match(/\buk\b/)) {
    return "UK";
  }

  // IRELAND
  if (combined.includes("ireland") || combined.includes("dublin") || combined.includes("cork")) {
    return "Ireland";
  }

  // NETHERLANDS
  if (combined.includes("netherlands") || combined.includes("holland") || combined.includes("amsterdam")) {
    return "Netherlands";
  }

  // ITALY
  if (combined.includes("italy") || combined.includes("italia") || combined.includes("s.r.l")) {
    return "Italy";
  }

  // FRANCE
  if (combined.includes("france") || combined.includes("paris")) {
    return "France";
  }

  // POLAND
  if (combined.includes("poland") || combined.includes("polska") || combined.includes("warsaw")) {
    return "Poland";
  }

  // SPAIN
  if (combined.includes("spain") || combined.includes("españa") || combined.includes("madrid")) {
    return "Spain";
  }

  return null;
}

// =============================================================================
// 3. EXTRACT DATE FROM TEXT (Regex Fallback)
// =============================================================================
function extractDateFromText(text: string): string | null {
  const euroPattern = /(\d{1,2})[./-](\d{1,2})[./-](20\d{2})/g;
  const isoPattern = /(20\d{2})-(\d{1,2})-(\d{1,2})/g;

  const dates: { date: string; index: number }[] = [];

  let match: RegExpExecArray | null;
  while ((match = euroPattern.exec(text)) !== null) {
    const before = text.substring(Math.max(0, match.index - 5), match.index);
    if (before.match(/\d+$/) || before.match(/[:/]$/)) continue;

    const day = match[1].padStart(2, "0");
    const month = match[2].padStart(2, "0");
    const year = match[3];

    const monthNum = parseInt(month, 10);
    if (monthNum >= 1 && monthNum <= 12) {
      dates.push({ date: `${year}-${month}-${day}`, index: match.index });
    }
  }

  while ((match = isoPattern.exec(text)) !== null) {
    const year = match[1];
    const month = match[2].padStart(2, "0");
    const day = match[3].padStart(2, "0");

    const monthNum = parseInt(month, 10);
    if (monthNum >= 1 && monthNum <= 12) {
      dates.push({ date: `${year}-${month}-${day}`, index: match.index });
    }
  }

  if (dates.length > 0) {
    return dates[dates.length - 1].date;
  }

  return null;
}

// =============================================================================
// 4. CALCULATE EXPIRY (Issue Date + 3 Years)
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
// 5. APPLY BUSINESS LOGIC (Master List + Dedup Fix)
// =============================================================================
function applyBusinessLogic(data: any, fullText: string): any {
  const cert = (data.certification || "").toLowerCase();
  const prod = (data.product_category || "").toLowerCase();
  const supp = (data.supplier_name || "").toLowerCase();
  const currentMeasure = (data.measure || "").toLowerCase();
  const textLower = fullText.toLowerCase();

  // Country inference
  if (!data.country || data.country === "Unknown" || data.country === "") {
    const inferredCountry = inferCountry(fullText, data.supplier_name || "");
    if (inferredCountry) data.country = inferredCountry;
  }

  // Date fallback
  if (!data.date_issued || data.date_issued === "null" || data.date_issued === "") {
    const extractedDate = extractDateFromText(fullText);
    if (extractedDate) data.date_issued = extractedDate;
  }

  // GLOVES
  const isGlove =
    cert.includes("en 455") || cert.includes("en 374") || cert.includes("en 420") ||
    cert.includes("en 388") || cert.includes("module b") || cert.includes("cat iii") ||
    prod.includes("glove") || prod.includes("nitrile") || supp.includes("intco") ||
    textLower.includes("en 455") || textLower.includes("en 374");

  if (isGlove) {
    data.scope = "+";
    data.measure = "EU Regulation 2016/425";
    if (!data.product_category || data.product_category === "Goods") data.product_category = "Gloves";

    // DEDUP FIX: Copy certification to certificate_number
    if (!data.certificate_number || data.certificate_number === "" || data.certificate_number === "null") {
      if (textLower.includes("en 455")) data.certificate_number = "EN 455";
      else if (textLower.includes("en 374")) data.certificate_number = "EN 374";
      else if (textLower.includes("en 420")) data.certificate_number = "EN 420";
      else if (data.certification) data.certificate_number = data.certification;
      else data.certificate_number = "Glove Test Report";
    }

    if (data.date_issued && (!data.date_expired || data.date_expired === "null")) {
      data.date_expired = calculateExpiry(data.date_issued);
    }
    return data;
  }

  // ISO 14001
  if (cert.includes("iso 14001") || cert.includes("14001") || textLower.includes("iso 14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    if (!data.certificate_number) data.certificate_number = data.certification || "ISO 14001";
    return data;
  }

  // ISO 45001
  if (cert.includes("iso 45001") || cert.includes("45001") || textLower.includes("iso 45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
    if (!data.certificate_number) data.certificate_number = data.certification || "ISO 45001";
    return data;
  }

  // ISO 27001
  if (cert.includes("iso 27001") || cert.includes("27001")) {
    data.scope = "!";
    data.measure = "EU GDPR";
    if (!data.certificate_number) data.certificate_number = data.certification || "ISO 27001";
    return data;
  }

  // FSC
  if ((cert.includes("fsc") && !cert.includes("fssc")) || (textLower.includes("fsc") && !textLower.includes("fssc"))) {
    data.scope = "!";
    data.measure = "FSC";
    if (!data.certificate_number) data.certificate_number = data.certification || "FSC";
    return data;
  }

  // Factory Certs
  const isFactoryCert =
    cert.includes("iso 9001") || cert.includes("9001") ||
    cert.includes("brc") || cert.includes("brcgs") ||
    cert.includes("iso 22000") || cert.includes("fssc") || cert.includes("gmp") ||
    textLower.includes("iso 9001") || textLower.includes("brcgs");

  if (isFactoryCert) {
    data.scope = "!";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 2023/2006";
    }
    if (!data.certificate_number) data.certificate_number = data.certification || "Management Certificate";
    return data;
  }

  // EU 10/2011
  if (cert.includes("10/2011") || textLower.includes("10/2011")) {
    data.scope = "+";
    data.measure = "(EC) No 10/2011";
    if (!data.certificate_number) data.certificate_number = data.certification || "EU 10/2011";
    return data;
  }

  // EN 13432
  if (cert.includes("en 13432") || cert.includes("compostable") || textLower.includes("en 13432")) {
    data.scope = "+";
    data.measure = "EN 13432";
    if (!data.certificate_number) data.certificate_number = data.certification || "EN 13432";
    return data;
  }

  // DoC / Migration
  const isProductTest =
    cert.includes("declaration of conformity") || cert.includes("declaration of compliance") ||
    cert.includes("doc") || cert.includes("migration") || cert.includes("food contact") ||
    textLower.includes("declaration of conformity") || textLower.includes("migration");

  if (isProductTest) {
    data.scope = "+";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 1935/2004";
    }
    if (!data.certificate_number) data.certificate_number = data.certification || "DoC";
    if (data.date_issued && (!data.date_expired || data.date_expired === "null")) {
      data.date_expired = calculateExpiry(data.date_issued);
    }
    return data;
  }

  // Business License
  if (cert.includes("business license") || textLower.includes("business license")) {
    data.scope = "!";
    data.measure = "National Regulation";
    if (!data.certificate_number) data.certificate_number = data.certification || "Business License";
    return data;
  }

  // Default
  if (!data.certificate_number || data.certificate_number === "" || data.certificate_number === "null") {
    data.certificate_number = data.certification || "Certificate";
  }
  if (!data.scope) data.scope = "!";
  if (!data.measure) data.measure = "National Regulation";

  return data;
}

// =============================================================================
// 6. EXTRACT JSON
// =============================================================================
function extractJSON(text: string, fullInput: string): any {
  console.log("Raw AI Response:", text.substring(0, 500));

  let cleanText = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  let data: any = {};

  try {
    const firstBrace = cleanText.indexOf("{");
    const lastBrace = cleanText.lastIndexOf("}");
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      data = JSON.parse(cleanText.substring(firstBrace, lastBrace + 1));
    }
  } catch {
    console.warn("JSON.parse failed, attempting regex...");
  }

  const scavenge = (key: string): string | null => {
    const regex = new RegExp(`"${key}"\\s*:\\s*(?:"([^"]*)"|null)`, "i");
    const match = cleanText.match(regex);
    return match && match[1] ? match[1] : null;
  };

  if (!data.supplier_name) data.supplier_name = scavenge("supplier_name");
  if (!data.certificate_number) data.certificate_number = scavenge("certificate_number");
  if (!data.country) data.country = scavenge("country");
  if (!data.scope) data.scope = scavenge("scope");
  if (!data.measure) data.measure = scavenge("measure");
  if (!data.certification) data.certification = scavenge("certification");
  if (!data.product_category) data.product_category = scavenge("product_category");
  if (!data.date_issued) data.date_issued = scavenge("date_issued");
  if (!data.date_expired) data.date_expired = scavenge("date_expired");

  data = applyBusinessLogic(data, fullInput);

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
// 7. TYPES
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
// 8. SYSTEM PROMPT
// =============================================================================
const systemPrompt = `
You are a Compliance Data Extraction Engine.
CRITICAL: OUTPUT RAW JSON ONLY. NO MARKDOWN.

### EXTRACTION RULES
- supplier_name: Look for "Applicant", "Manufacturer", company letterhead.
- certificate_number: Extract the certificate/report number if visible.
- country: Extract from address. If unclear, return null.
- certification: The standard name (e.g., "EN 455", "ISO 9001").
- date_issued: Issue date. Format: YYYY-MM-DD. "09.10.2024" = Oct 9th.
- date_expired: Expiry date. If not found, return null.

### OUTPUT
{
  "supplier_name": "string",
  "certificate_number": "string or null",
  "country": "string or null",
  "scope": "! or +",
  "measure": "string",
  "certification": "string",
  "product_category": "string",
  "date_issued": "YYYY-MM-DD or null",
  "date_expired": "YYYY-MM-DD or null"
}
`;

// =============================================================================
// 9. MAIN EXPORT
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

    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: "gpt-4o",
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userContent },
        ],
        max_tokens: 1500,
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || "OpenAI API failed");
    }

    const data = await response.json();
    const rawContent = data.choices[0]?.message?.content;

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

    return extractJSON(rawContent, truncatedText || textContent || "");
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
