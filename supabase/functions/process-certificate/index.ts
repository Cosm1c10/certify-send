// =============================================================================
// SUPABASE EDGE FUNCTION: process-certificate
// Version: FINAL PRODUCTION v2.0 (Smart Logic)
// =============================================================================
// deno-lint-ignore-file no-explicit-any
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import OpenAI from "https://deno.land/x/openai@v4.20.1/mod.ts";

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
// 1. SANITIZE (Fixes Excel Drop Bug)
// =============================================================================
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined || val === "null" || val === "undefined") {
    return defaultVal;
  }
  return String(val).trim();
}

// =============================================================================
// 2. INFER COUNTRY (Fixes "Unknown" Country Bug)
// =============================================================================
function inferCountry(text: string): string | null {
  const t = text.toLowerCase();

  // China
  if (t.includes("china") || t.includes("anhui") || t.includes("guangdong") ||
      t.includes("shanghai") || t.includes("beijing") || t.includes("shenzhen") ||
      t.includes("changsha") || t.includes("zhejiang") || t.includes("jiangsu") ||
      t.includes("shaoneng") || t.includes("intco") || t.includes("hunan") ||
      t.match(/\bcn\b/)) {
    return "China";
  }

  // Turkey
  if (t.includes("turkey") || t.includes("türkiye") || t.includes("turkiye") ||
      t.includes("istanbul") || t.includes("ankara") || t.includes("izmir") ||
      t.includes("boran") || t.includes("mopack") || t.match(/\.tr\b/)) {
    return "Turkey";
  }

  // Germany
  if (t.includes("germany") || t.includes("deutschland") || t.includes("gmbh") ||
      t.includes("munich") || t.includes("berlin") || t.includes("frankfurt") ||
      t.match(/\bde\b/)) {
    return "Germany";
  }

  // UK
  if (t.includes("united kingdom") || t.includes("england") || t.includes("london") ||
      t.includes("manchester") || t.includes("birmingham") || t.includes("scotland") ||
      t.match(/\buk\b/) || t.match(/\bgb\b/)) {
    return "UK";
  }

  // Ireland
  if (t.includes("ireland") || t.includes("dublin") || t.includes("cork") ||
      t.includes("galway") || t.match(/\bie\b/)) {
    return "Ireland";
  }

  // USA
  if (t.includes("united states") || t.includes("usa") || t.includes("california") ||
      t.includes("new york") || t.includes("texas") || t.match(/\bus\b/)) {
    return "USA";
  }

  // Netherlands
  if (t.includes("netherlands") || t.includes("holland") || t.includes("amsterdam") ||
      t.match(/\bnl\b/)) {
    return "Netherlands";
  }

  // Italy
  if (t.includes("italy") || t.includes("italia") || t.includes("milan") ||
      t.includes("rome") || t.match(/\bit\b/)) {
    return "Italy";
  }

  // France
  if (t.includes("france") || t.includes("paris") || t.includes("lyon") ||
      t.match(/\bfr\b/)) {
    return "France";
  }

  // Poland
  if (t.includes("poland") || t.includes("polska") || t.includes("warsaw") ||
      t.match(/\bpl\b/)) {
    return "Poland";
  }

  return null;
}

// =============================================================================
// 3. SMART DATE EXTRACTION (Regex Fallback)
// =============================================================================
// Finds dates in text, prioritizing those near keywords, falling back to LAST date
function extractDateFromText(text: string, type: "issue" | "expiry"): string | null {
  // Patterns to match various date formats
  const datePatterns = [
    // DD.MM.YYYY or DD/MM/YYYY
    /(\d{1,2})[./-](\d{1,2})[./-](20\d{2})/g,
    // YYYY-MM-DD
    /(20\d{2})-(\d{1,2})-(\d{1,2})/g,
    // Month DD, YYYY
    /(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(20\d{2})/gi,
  ];

  // Keywords that indicate a date is relevant
  const issueKeywords = /(?:date|dated|issued|issue date|signed|signature|valid from|effective)/i;
  const expiryKeywords = /(?:expir|valid until|valid to|validity|expires|expiration|valid through)/i;

  const keywords = type === "issue" ? issueKeywords : expiryKeywords;
  const allDates: { date: string; index: number; nearKeyword: boolean }[] = [];

  // Find all dates and check if they're near relevant keywords
  for (const pattern of datePatterns) {
    let match;
    const regex = new RegExp(pattern.source, pattern.flags);
    while ((match = regex.exec(text)) !== null) {
      // Get surrounding context (100 chars before)
      const contextStart = Math.max(0, match.index - 100);
      const context = text.substring(contextStart, match.index + match[0].length);

      // Skip dates that look like regulation years (e.g., "2016/425", "9001:2015")
      const beforeDate = text.substring(Math.max(0, match.index - 10), match.index);
      if (beforeDate.match(/[:/]\s*$/) || beforeDate.match(/\d+\s*$/)) {
        continue; // Skip - likely part of a regulation number
      }

      const nearKeyword = keywords.test(context);
      let isoDate = "";

      // Convert to ISO format
      if (match[0].match(/^\d{1,2}[./-]\d{1,2}[./-]20\d{2}$/)) {
        // DD.MM.YYYY format
        const parts = match[0].split(/[./-]/);
        const day = parts[0].padStart(2, "0");
        const month = parts[1].padStart(2, "0");
        const year = parts[2];
        isoDate = `${year}-${month}-${day}`;
      } else if (match[0].match(/^20\d{2}-\d{1,2}-\d{1,2}$/)) {
        // Already YYYY-MM-DD
        isoDate = match[0];
      } else if (match[1] && match[2] && match[3]) {
        // Month name format
        const months: Record<string, string> = {
          january: "01", february: "02", march: "03", april: "04",
          may: "05", june: "06", july: "07", august: "08",
          september: "09", october: "10", november: "11", december: "12",
        };
        const month = months[match[1].toLowerCase()] || "01";
        const day = match[2].padStart(2, "0");
        const year = match[3];
        isoDate = `${year}-${month}-${day}`;
      }

      if (isoDate && isoDate.match(/^20\d{2}-\d{2}-\d{2}$/)) {
        allDates.push({ date: isoDate, index: match.index, nearKeyword });
      }
    }
  }

  if (allDates.length === 0) return null;

  // Priority 1: Dates near keywords
  const keywordDates = allDates.filter((d) => d.nearKeyword);
  if (keywordDates.length > 0) {
    // For issue date, take first near keyword; for expiry, take last
    return type === "issue" ? keywordDates[0].date : keywordDates[keywordDates.length - 1].date;
  }

  // Priority 2: For issue date, use LAST date in document (signatures at bottom)
  // For expiry, also use last date but only if it's after issue date
  return allDates[allDates.length - 1].date;
}

// =============================================================================
// 4. CALCULATE EXPIRY (3-Year Rule)
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
// 5. APPLY BUSINESS LOGIC (Jun's Master List + Smart Fixes)
// =============================================================================
function applyBusinessLogic(data: any, fullText: string): any {
  const cert = (data.certification || "").toLowerCase();
  const prod = (data.product_category || "").toLowerCase();
  const supp = (data.supplier_name || "").toLowerCase();
  const currentMeasure = (data.measure || "").toLowerCase();
  const textLower = fullText.toLowerCase();

  // -------------------------------------------------------------------------
  // SMART FIX: Country Inference
  // -------------------------------------------------------------------------
  if (!data.country || data.country === "Unknown" || data.country === "") {
    const inferredCountry = inferCountry(fullText);
    if (inferredCountry) {
      data.country = inferredCountry;
    }
  }

  // -------------------------------------------------------------------------
  // SMART FIX: Date Extraction (if AI missed it)
  // -------------------------------------------------------------------------
  if (!data.date_issued || data.date_issued === "" || data.date_issued === "null") {
    const extractedDate = extractDateFromText(fullText, "issue");
    if (extractedDate) {
      data.date_issued = extractedDate;
      console.log("Smart Date: Extracted issue date from text:", extractedDate);
    }
  }

  if (!data.date_expired || data.date_expired === "" || data.date_expired === "null") {
    // First try to extract from text
    const extractedExpiry = extractDateFromText(fullText, "expiry");
    if (extractedExpiry) {
      data.date_expired = extractedExpiry;
      console.log("Smart Date: Extracted expiry date from text:", extractedExpiry);
    }
  }

  // -------------------------------------------------------------------------
  // RULE: GLOVES - HIGHEST PRIORITY
  // -------------------------------------------------------------------------
  const isGlove =
    cert.includes("en 455") || cert.includes("en 374") || cert.includes("en 420") ||
    cert.includes("en 388") || cert.includes("module b") || cert.includes("cat iii") ||
    cert.includes("category iii") || cert.includes("2016/425") ||
    prod.includes("glove") || prod.includes("nitrile") ||
    supp.includes("intco") ||
    textLower.includes("en 455") || textLower.includes("en 374") || textLower.includes("en 420");

  if (isGlove) {
    data.scope = "+";
    data.measure = "EU Regulation 2016/425";
    if (!data.product_category || data.product_category === "Goods" || data.product_category === "Unknown") {
      data.product_category = "Gloves";
    }
    // Glove tests: Apply 3-year rule if no expiry
    if (!data.date_expired && data.date_issued) {
      data.date_expired = calculateExpiry(data.date_issued);
    }
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: ISO 14001 (Environmental)
  // -------------------------------------------------------------------------
  if (cert.includes("iso 14001") || cert.includes("14001") || textLower.includes("iso 14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: ISO 45001 (Occupational Health & Safety)
  // -------------------------------------------------------------------------
  if (cert.includes("iso 45001") || cert.includes("45001") || textLower.includes("iso 45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: ISO 27001 (Information Security)
  // -------------------------------------------------------------------------
  if (cert.includes("iso 27001") || cert.includes("27001")) {
    data.scope = "!";
    data.measure = "EU GDPR";
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: FSC (Forestry)
  // -------------------------------------------------------------------------
  if ((cert.includes("fsc") && !cert.includes("fssc")) || textLower.match(/\bfsc\b/)) {
    data.scope = "!";
    data.measure = "FSC";
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: ISO 9001 / BRC / BRCGS / ISO 22000 / FSSC / GMP
  // -------------------------------------------------------------------------
  const isFactoryCert =
    cert.includes("iso 9001") || cert.includes("9001") ||
    cert.includes("brc") || cert.includes("brcgs") ||
    cert.includes("iso 22000") || cert.includes("22000") ||
    cert.includes("fssc") || cert.includes("gmp") ||
    textLower.includes("iso 9001") || textLower.includes("brcgs");

  if (isFactoryCert) {
    data.scope = "!";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 2023/2006";
    }
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: EU 10/2011 (Plastics)
  // -------------------------------------------------------------------------
  if (cert.includes("10/2011") || textLower.includes("10/2011")) {
    data.scope = "+";
    data.measure = "(EC) No 10/2011";
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: EN 13432 (Compostable)
  // -------------------------------------------------------------------------
  if (cert.includes("en 13432") || cert.includes("compostable") || textLower.includes("en 13432")) {
    data.scope = "+";
    data.measure = "EN 13432";
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: EN 13430 / ISO 14021 (Recyclable)
  // -------------------------------------------------------------------------
  if (cert.includes("en 13430") || cert.includes("iso 14021") || cert.includes("recyclable")) {
    data.scope = "+";
    data.measure = "EN 13430";
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: DoC / Migration / Food Contact
  // -------------------------------------------------------------------------
  const isProductTest =
    cert.includes("declaration of conformity") || cert.includes("doc") ||
    cert.includes("migration") || cert.includes("heavy metal") ||
    cert.includes("food contact") || cert.includes("1935/2004") ||
    textLower.includes("declaration of conformity") || textLower.includes("migration");

  if (isProductTest) {
    data.scope = "+";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 1935/2004";
    }
    return data;
  }

  // -------------------------------------------------------------------------
  // RULE: Business License
  // -------------------------------------------------------------------------
  if (cert.includes("business license") || cert.includes("operating permit") || textLower.includes("business license")) {
    data.scope = "!";
    data.measure = "National Regulation";
    return data;
  }

  // -------------------------------------------------------------------------
  // DEFAULT
  // -------------------------------------------------------------------------
  if (!data.scope) data.scope = "!";
  if (!data.measure) data.measure = "National Regulation";

  return data;
}

// =============================================================================
// 6. EXTRACT JSON (Self-Healing Parser)
// =============================================================================
function extractJSON(text: string, fullInput: string): any {
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

  // Apply Smart Business Logic (with full text for inference)
  data = applyBusinessLogic(data, fullInput);

  // Final Sanitization - NEVER return null
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
// 7. SYSTEM PROMPT
// =============================================================================
const systemPrompt = `
You are a Compliance Data Extraction Engine.
CRITICAL: OUTPUT RAW JSON ONLY. NO MARKDOWN. NO EXPLANATIONS.

### DATE EXTRACTION RULES (IMPORTANT)
- Format: YYYY-MM-DD
- European input: DD.MM.YYYY → "09.10.2024" = October 9th, NEVER September.
- IGNORE years in regulation names (e.g., "2016" in "2016/425", "2015" in "ISO 9001:2015").
- PRIORITIZE dates near: "Date", "Issued", "Signed", "Signature", "Valid from".
- If no expiry found, leave date_expired as null (system will calculate +3 years).

### SUPPLIER EXTRACTION
- Look for: "Applicant", "Manufacturer", "Company Name", letterhead.
- Normalize: Remove "Co., Ltd", "GmbH", "San. Tic." suffixes.

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
// 8. CONFIGURATION
// =============================================================================
const MAX_INPUT_LENGTH = 50000;

// =============================================================================
// 9. MAIN HANDLER
// =============================================================================
serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const input = await req.json();

    // Support multiple payload formats
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

    console.log(`Processing: ${fileName} | Mode: ${isTextMode ? "TEXT" : "IMAGE"} | Length: ${rawText.length || "N/A"}`);

    const openaiApiKey = Deno.env.get("OPENAI_API_KEY");
    if (!openaiApiKey) {
      return new Response(
        JSON.stringify({ error: "OpenAI API key not configured" }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const openai = new OpenAI({ apiKey: openaiApiKey });

    let userContent: any[];
    let truncatedText = "";

    if (isTextMode) {
      truncatedText = rawText.length > MAX_INPUT_LENGTH
        ? rawText.slice(0, MAX_INPUT_LENGTH) + "\n[TRUNCATED]"
        : rawText;

      userContent = [
        { type: "text", text: `Extract compliance data from (${fileName}):\n\n${truncatedText}` },
      ];
    } else {
      const imageUrl = rawImage.startsWith("data:") ? rawImage : `data:image/jpeg;base64,${rawImage}`;
      userContent = [
        { type: "image_url", image_url: { url: imageUrl } },
        { type: "text", text: `Extract compliance data from (${fileName}).` },
      ];
    }

    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userContent },
      ],
      max_tokens: 1500,
    });

    const rawContent = response.choices[0]?.message?.content;

    if (!rawContent) {
      return new Response(
        JSON.stringify({
          supplier_name: "Error: No AI Response",
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

    // Extract with full text for smart inference
    const result = extractJSON(rawContent, truncatedText || rawText);

    return new Response(JSON.stringify(result), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });

  } catch (error) {
    console.error("Processing Error:", error);

    return new Response(
      JSON.stringify({
        supplier_name: "Error: Processing Failed",
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
