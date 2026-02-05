// =============================================================================
// LOCAL PROCESSING: processWithOpenAI.ts
// Version: FINAL PRODUCTION (Feb 2026)
// Client: Catering Disposables (Jun & Saurebh)
// =============================================================================

const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;
const MAX_INPUT_LENGTH = 50000;

// =============================================================================
// 1. SANITIZE FUNCTION (Fixes Excel Drop Bug)
// =============================================================================
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined || val === "null" || val === "undefined") {
    return defaultVal;
  }
  return String(val).trim();
}

// =============================================================================
// 2. APPLY BUSINESS LOGIC (Jun & Saurebh's Rules)
// =============================================================================
function applyBusinessLogic(data: any): any {
  const cert = (data.certification || "").toLowerCase();
  const prod = (data.product_category || "").toLowerCase();
  const supp = (data.supplier_name || "").toLowerCase();
  const currentMeasure = (data.measure || "").toLowerCase();

  // RULE C: GLOVES - HIGHEST PRIORITY
  const isGlove =
    cert.includes("en 455") ||
    cert.includes("en 374") ||
    cert.includes("en 420") ||
    cert.includes("en 388") ||
    cert.includes("module b") ||
    cert.includes("cat iii") ||
    cert.includes("category iii") ||
    cert.includes("2016/425") ||
    prod.includes("glove") ||
    prod.includes("nitrile") ||
    supp.includes("intco");

  if (isGlove) {
    data.scope = "+";
    data.measure = "EU Regulation 2016/425";
    if (!data.product_category || data.product_category === "Goods" || data.product_category === "Unknown") {
      data.product_category = "Gloves";
    }
    return data;
  }

  // RULE A: GENERAL (!) - Factory Certificates
  if (cert.includes("iso 14001") || cert.includes("14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    return data;
  }

  if (cert.includes("iso 45001") || cert.includes("45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
    return data;
  }

  if (cert.includes("iso 27001") || cert.includes("27001")) {
    data.scope = "!";
    data.measure = "EU GDPR";
    return data;
  }

  if (cert.includes("fsc") && !cert.includes("fssc")) {
    data.scope = "!";
    data.measure = "FSC";
    return data;
  }

  const isFactoryCert =
    cert.includes("iso 9001") ||
    cert.includes("9001") ||
    cert.includes("brc") ||
    cert.includes("brcgs") ||
    cert.includes("iso 22000") ||
    cert.includes("22000") ||
    cert.includes("fssc") ||
    cert.includes("gmp") ||
    cert.includes("good manufacturing");

  if (isFactoryCert) {
    data.scope = "!";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 2023/2006";
    }
    return data;
  }

  // RULE B: SPECIFIC (+) - Product Certificates
  if (cert.includes("10/2011") || cert.includes("eu 10/2011") || cert.includes("plastic")) {
    data.scope = "+";
    data.measure = "(EC) No 10/2011";
    return data;
  }

  if (cert.includes("en 13432") || cert.includes("13432") || cert.includes("compostable") || cert.includes("din certco")) {
    data.scope = "+";
    data.measure = "EN 13432";
    return data;
  }

  if (cert.includes("en 13430") || cert.includes("13430") || cert.includes("iso 14021") || cert.includes("recyclable")) {
    data.scope = "+";
    data.measure = "EN 13430";
    return data;
  }

  const isProductTest =
    cert.includes("declaration of conformity") ||
    cert.includes("doc") ||
    cert.includes("migration") ||
    cert.includes("heavy metal") ||
    cert.includes("food contact") ||
    cert.includes("1935/2004") ||
    cert.includes("microwave") ||
    cert.includes("dishwasher");

  if (isProductTest) {
    data.scope = "+";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 1935/2004";
    }
    return data;
  }

  // DEFAULT
  if (cert.includes("business license") || cert.includes("operating permit")) {
    data.scope = "!";
    data.measure = "National Regulation";
    return data;
  }

  if (!data.scope) data.scope = "!";
  if (!data.measure) data.measure = "National Regulation";

  return data;
}

// =============================================================================
// 3. CALCULATE EXPIRY (3-Year Rule)
// =============================================================================
function calculateExpiry(dateIssued: string | null): string {
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
// 4. EXTRACT JSON (Self-Healing Parser)
// =============================================================================
function extractJSON(text: string): any {
  console.log("Raw AI Response:", text.substring(0, 500));

  let cleanText = text.replace(/```json/gi, "").replace(/```/g, "").trim();
  let data: any = {};

  // Attempt 1: Standard JSON Parse
  try {
    const firstBrace = cleanText.indexOf("{");
    const lastBrace = cleanText.lastIndexOf("}");
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      data = JSON.parse(cleanText.substring(firstBrace, lastBrace + 1));
    }
  } catch (e) {
    console.warn("JSON.parse failed, attempting regex extraction...");
  }

  // Attempt 2: Regex Extraction
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

  // Apply Business Logic
  data = applyBusinessLogic(data);

  // Apply 3-Year Rule
  if (!data.date_expired && data.date_issued) {
    const cert = (data.certification || "").toLowerCase();
    const isTestReport =
      cert.includes("en 455") ||
      cert.includes("en 374") ||
      cert.includes("migration") ||
      cert.includes("test") ||
      cert.includes("report");
    if (isTestReport) {
      data.date_expired = calculateExpiry(data.date_issued);
    }
  }

  // Final Sanitization
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
// 5. TYPES
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
// 6. SYSTEM PROMPT
// =============================================================================
const systemPrompt = `
You are a Compliance Data Extraction Engine for Catering Disposables Ltd.
CRITICAL: OUTPUT RAW JSON ONLY. NO EXPLANATIONS. NO MARKDOWN.

### DOCUMENT HANDLING
- Treat Test Reports (EN 455, EN 374, Migration Tests) as VALID certificates.
- For multi-language documents (Chinese + English), extract English data.
- Supplier: Look for "Applicant", "Manufacturer", or company name.

### DATE RULES
- Input format: DD.MM.YYYY (European). Output: YYYY-MM-DD.
- "09.10.2024" = October 9th, 2024. NEVER September.
- If no expiry date exists, calculate: Issue Date + 3 Years.

### CLASSIFICATION RULES
- GENERAL (!): ISO 9001, BRC, ISO 22000, FSSC, GMP → Measure: "(EC) No 2023/2006"
- ISO 14001 → "EU Waste Framework Directive (2008/98/EC)"
- ISO 45001 → "EU Directive 89/391/EEC"
- SPECIFIC (+): DoC, Migration, Food Contact → Measure: "(EC) No 1935/2004"
- GLOVES (EN 455, EN 374, EN 420) → Scope: "+", Measure: "EU Regulation 2016/425"

### OUTPUT JSON SCHEMA
{
  "supplier_name": "string",
  "country": "string",
  "scope": "! or +",
  "measure": "string",
  "certification": "string",
  "product_category": "string",
  "date_issued": "YYYY-MM-DD",
  "date_expired": "YYYY-MM-DD"
}
`;

// =============================================================================
// 7. MAIN EXPORT
// =============================================================================
export async function processWithOpenAI(
  base64Image: string,
  filename?: string,
  textContent?: string
): Promise<CertificateExtractionResult> {
  try {
    if (!OPENAI_API_KEY) {
      throw new Error("OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local");
    }

    const isTextOnly = !!textContent && !base64Image;
    console.log(`Processing: ${filename || "unknown"} | Mode: ${isTextOnly ? "TEXT" : "IMAGE"}`);

    let userContent: Array<{ type: string; text?: string; image_url?: { url: string } }>;

    if (isTextOnly && textContent) {
      const truncatedText =
        textContent.length > MAX_INPUT_LENGTH
          ? textContent.slice(0, MAX_INPUT_LENGTH) + "\n\n[TRUNCATED]"
          : textContent;

      userContent = [
        { type: "text", text: `Extract compliance data from this file (${filename || "unknown"}):\n\n${truncatedText}` },
      ];
    } else {
      const imageUrl = base64Image.startsWith("data:") ? base64Image : `data:image/jpeg;base64,${base64Image}`;

      userContent = [
        { type: "image_url", image_url: { url: imageUrl } },
        { type: "text", text: `Extract compliance data from this certificate (${filename || "unknown"}).` },
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
      console.error("OpenAI API Error:", error);
      throw new Error(error.error?.message || "OpenAI API request failed");
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

    return extractJSON(rawContent);
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
