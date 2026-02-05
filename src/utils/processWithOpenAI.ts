// =============================================================================
// LOCAL PROCESSING: processWithOpenAI.ts
// Version: FINAL PRODUCTION v2.0 (Smart Logic)
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
// 2. INFER COUNTRY
// =============================================================================
function inferCountry(text: string): string | null {
  const t = text.toLowerCase();

  if (t.includes("china") || t.includes("anhui") || t.includes("guangdong") ||
      t.includes("shanghai") || t.includes("beijing") || t.includes("shenzhen") ||
      t.includes("changsha") || t.includes("zhejiang") || t.includes("jiangsu") ||
      t.includes("shaoneng") || t.includes("intco") || t.includes("hunan")) {
    return "China";
  }

  if (t.includes("turkey") || t.includes("türkiye") || t.includes("turkiye") ||
      t.includes("istanbul") || t.includes("ankara") || t.includes("boran") || t.includes("mopack")) {
    return "Turkey";
  }

  if (t.includes("germany") || t.includes("deutschland") || t.includes("gmbh")) {
    return "Germany";
  }

  if (t.includes("united kingdom") || t.includes("england") || t.includes("london") || t.match(/\buk\b/)) {
    return "UK";
  }

  if (t.includes("ireland") || t.includes("dublin") || t.includes("cork")) {
    return "Ireland";
  }

  if (t.includes("netherlands") || t.includes("holland") || t.includes("amsterdam")) {
    return "Netherlands";
  }

  if (t.includes("italy") || t.includes("italia") || t.includes("milan")) {
    return "Italy";
  }

  if (t.includes("france") || t.includes("paris")) {
    return "France";
  }

  if (t.includes("poland") || t.includes("polska") || t.includes("warsaw")) {
    return "Poland";
  }

  return null;
}

// =============================================================================
// 3. SMART DATE EXTRACTION
// =============================================================================
function extractDateFromText(text: string, type: "issue" | "expiry"): string | null {
  const datePatterns = [
    /(\d{1,2})[./-](\d{1,2})[./-](20\d{2})/g,
    /(20\d{2})-(\d{1,2})-(\d{1,2})/g,
  ];

  const issueKeywords = /(?:date|dated|issued|issue date|signed|signature|valid from|effective)/i;
  const expiryKeywords = /(?:expir|valid until|valid to|validity|expires|expiration)/i;
  const keywords = type === "issue" ? issueKeywords : expiryKeywords;

  const allDates: { date: string; index: number; nearKeyword: boolean }[] = [];

  for (const pattern of datePatterns) {
    let match: RegExpExecArray | null;
    const regex = new RegExp(pattern.source, pattern.flags);
    while ((match = regex.exec(text)) !== null) {
      const contextStart = Math.max(0, match.index - 100);
      const context = text.substring(contextStart, match.index + match[0].length);

      // Skip regulation years
      const beforeDate = text.substring(Math.max(0, match.index - 10), match.index);
      if (beforeDate.match(/[:/]\s*$/) || beforeDate.match(/\d+\s*$/)) {
        continue;
      }

      const nearKeyword = keywords.test(context);
      let isoDate = "";

      if (match[0].match(/^\d{1,2}[./-]\d{1,2}[./-]20\d{2}$/)) {
        const parts = match[0].split(/[./-]/);
        isoDate = `${parts[2]}-${parts[1].padStart(2, "0")}-${parts[0].padStart(2, "0")}`;
      } else if (match[0].match(/^20\d{2}-\d{1,2}-\d{1,2}$/)) {
        isoDate = match[0];
      }

      if (isoDate && isoDate.match(/^20\d{2}-\d{2}-\d{2}$/)) {
        allDates.push({ date: isoDate, index: match.index, nearKeyword });
      }
    }
  }

  if (allDates.length === 0) return null;

  const keywordDates = allDates.filter((d) => d.nearKeyword);
  if (keywordDates.length > 0) {
    return type === "issue" ? keywordDates[0].date : keywordDates[keywordDates.length - 1].date;
  }

  return allDates[allDates.length - 1].date;
}

// =============================================================================
// 4. CALCULATE EXPIRY
// =============================================================================
function calculateExpiry(dateIssued: string): string {
  if (!dateIssued) return "";
  try {
    const parts = dateIssued.split("-");
    if (parts.length === 3) {
      const year = parseInt(parts[0], 10) + 3;
      return `${year}-${parts[1]}-${parts[2]}`;
    }
  } catch (e) {
    console.warn("Could not calculate expiry:", e);
  }
  return "";
}

// =============================================================================
// 5. APPLY BUSINESS LOGIC
// =============================================================================
function applyBusinessLogic(data: any, fullText: string): any {
  const cert = (data.certification || "").toLowerCase();
  const prod = (data.product_category || "").toLowerCase();
  const supp = (data.supplier_name || "").toLowerCase();
  const currentMeasure = (data.measure || "").toLowerCase();
  const textLower = fullText.toLowerCase();

  // Country inference
  if (!data.country || data.country === "Unknown" || data.country === "") {
    const inferredCountry = inferCountry(fullText);
    if (inferredCountry) data.country = inferredCountry;
  }

  // Date extraction
  if (!data.date_issued || data.date_issued === "" || data.date_issued === "null") {
    const extractedDate = extractDateFromText(fullText, "issue");
    if (extractedDate) data.date_issued = extractedDate;
  }

  if (!data.date_expired || data.date_expired === "" || data.date_expired === "null") {
    const extractedExpiry = extractDateFromText(fullText, "expiry");
    if (extractedExpiry) data.date_expired = extractedExpiry;
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
    if (!data.date_expired && data.date_issued) data.date_expired = calculateExpiry(data.date_issued);
    return data;
  }

  // ISO 14001
  if (cert.includes("iso 14001") || cert.includes("14001") || textLower.includes("iso 14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    return data;
  }

  // ISO 45001
  if (cert.includes("iso 45001") || cert.includes("45001") || textLower.includes("iso 45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
    return data;
  }

  // ISO 27001
  if (cert.includes("iso 27001") || cert.includes("27001")) {
    data.scope = "!";
    data.measure = "EU GDPR";
    return data;
  }

  // FSC
  if ((cert.includes("fsc") && !cert.includes("fssc")) || textLower.match(/\bfsc\b/)) {
    data.scope = "!";
    data.measure = "FSC";
    return data;
  }

  // ISO 9001 / BRC / GMP
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
    return data;
  }

  // EU 10/2011
  if (cert.includes("10/2011") || textLower.includes("10/2011")) {
    data.scope = "+";
    data.measure = "(EC) No 10/2011";
    return data;
  }

  // EN 13432
  if (cert.includes("en 13432") || cert.includes("compostable") || textLower.includes("en 13432")) {
    data.scope = "+";
    data.measure = "EN 13432";
    return data;
  }

  // DoC / Migration
  const isProductTest =
    cert.includes("declaration of conformity") || cert.includes("doc") ||
    cert.includes("migration") || cert.includes("food contact") ||
    textLower.includes("declaration of conformity") || textLower.includes("migration");

  if (isProductTest) {
    data.scope = "+";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 1935/2004";
    }
    return data;
  }

  // Business License
  if (cert.includes("business license") || textLower.includes("business license")) {
    data.scope = "!";
    data.measure = "National Regulation";
    return data;
  }

  // Default
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
  } catch (e) {
    console.warn("JSON.parse failed, attempting regex...");
  }

  const scavenge = (key: string): string | null => {
    const regex = new RegExp(`"${key}"\\s*:\\s*(?:"([^"]*)"|null)`, "i");
    const match = cleanText.match(regex);
    return match && match[1] ? match[1] : null;
  };

  if (!data.supplier_name) data.supplier_name = scavenge("supplier_name");
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

### DATE RULES
- Format: YYYY-MM-DD
- European: DD.MM.YYYY → "09.10.2024" = October 9th.
- IGNORE years in regulation names ("2016" in "2016/425").
- PRIORITIZE dates near: "Date", "Issued", "Signed".

### CLASSIFICATION
- Gloves (EN 455/374/420): scope="+", measure="EU Regulation 2016/425"
- ISO 9001/BRC/GMP: scope="!", measure="(EC) No 2023/2006"
- DoC/Migration: scope="+", measure="(EC) No 1935/2004"

### OUTPUT
{
  "supplier_name": "string",
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
        { type: "text", text: `Extract compliance data from (${filename || "unknown"}):\n\n${truncatedText}` },
      ];
    } else {
      const imageUrl = base64Image.startsWith("data:") ? base64Image : `data:image/jpeg;base64,${base64Image}`;
      userContent = [
        { type: "image_url", image_url: { url: imageUrl } },
        { type: "text", text: `Extract compliance data from (${filename || "unknown"}).` },
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
