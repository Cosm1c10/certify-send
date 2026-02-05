// =============================================================================
// SUPABASE EDGE FUNCTION: process-certificate
// Version: FINAL PRODUCTION v3.0 (Dedup Fix + Country Inference + Date Fallback)
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
// 1. SANITIZE (Prevents Excel Null Errors)
// =============================================================================
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined || val === "null" || val === "undefined") {
    return defaultVal;
  }
  return String(val).trim();
}

// =============================================================================
// 2. INFER COUNTRY (Checks BOTH full text AND supplier_name)
// =============================================================================
function inferCountry(fullText: string, supplierName: string): string | null {
  // Combine both sources for maximum coverage
  const combined = (fullText + " " + supplierName).toLowerCase();

  // CHINA - Provinces, cities, and company indicators
  if (
    combined.includes("china") ||
    combined.includes("anhui") ||
    combined.includes("guangdong") ||
    combined.includes("shanghai") ||
    combined.includes("beijing") ||
    combined.includes("shenzhen") ||
    combined.includes("changsha") ||
    combined.includes("zhejiang") ||
    combined.includes("jiangsu") ||
    combined.includes("hunan") ||
    combined.includes("wenzhou") ||
    combined.includes("fujian") ||
    combined.includes("shandong") ||
    combined.includes("hebei") ||
    combined.includes("henan") ||
    combined.includes("shaoneng") ||
    combined.includes("intco") ||
    combined.match(/\bcn\b/) ||
    combined.match(/[\u4e00-\u9fff]/) // Chinese characters
  ) {
    return "China";
  }

  // TURKEY
  if (
    combined.includes("turkey") ||
    combined.includes("türkiye") ||
    combined.includes("turkiye") ||
    combined.includes("istanbul") ||
    combined.includes("ankara") ||
    combined.includes("izmir") ||
    combined.includes("bursa") ||
    combined.includes("boran") ||
    combined.includes("mopack") ||
    combined.includes("san. ve tic") ||
    combined.match(/\.tr\b/)
  ) {
    return "Turkey";
  }

  // GERMANY
  if (
    combined.includes("germany") ||
    combined.includes("deutschland") ||
    combined.includes("gmbh") ||
    combined.includes("munich") ||
    combined.includes("berlin") ||
    combined.includes("frankfurt") ||
    combined.includes("hamburg") ||
    combined.match(/\bde\b/)
  ) {
    return "Germany";
  }

  // UK
  if (
    combined.includes("united kingdom") ||
    combined.includes("england") ||
    combined.includes("london") ||
    combined.includes("manchester") ||
    combined.includes("birmingham") ||
    combined.includes("scotland") ||
    combined.includes("wales") ||
    combined.match(/\buk\b/) ||
    combined.match(/\bgb\b/)
  ) {
    return "UK";
  }

  // IRELAND
  if (
    combined.includes("ireland") ||
    combined.includes("dublin") ||
    combined.includes("cork") ||
    combined.includes("galway") ||
    combined.includes("limerick") ||
    combined.match(/\bie\b/)
  ) {
    return "Ireland";
  }

  // USA
  if (
    combined.includes("united states") ||
    combined.includes("usa") ||
    combined.includes("california") ||
    combined.includes("new york") ||
    combined.includes("texas") ||
    combined.includes("florida") ||
    combined.match(/\bus\b/)
  ) {
    return "USA";
  }

  // NETHERLANDS
  if (
    combined.includes("netherlands") ||
    combined.includes("holland") ||
    combined.includes("amsterdam") ||
    combined.includes("rotterdam") ||
    combined.match(/\bnl\b/)
  ) {
    return "Netherlands";
  }

  // ITALY
  if (
    combined.includes("italy") ||
    combined.includes("italia") ||
    combined.includes("milan") ||
    combined.includes("rome") ||
    combined.includes("s.r.l") ||
    combined.match(/\bit\b/)
  ) {
    return "Italy";
  }

  // FRANCE
  if (
    combined.includes("france") ||
    combined.includes("paris") ||
    combined.includes("lyon") ||
    combined.includes("marseille") ||
    combined.match(/\bfr\b/)
  ) {
    return "France";
  }

  // POLAND
  if (
    combined.includes("poland") ||
    combined.includes("polska") ||
    combined.includes("warsaw") ||
    combined.includes("krakow") ||
    combined.match(/\bpl\b/)
  ) {
    return "Poland";
  }

  // SPAIN
  if (
    combined.includes("spain") ||
    combined.includes("españa") ||
    combined.includes("madrid") ||
    combined.includes("barcelona") ||
    combined.match(/\bes\b/)
  ) {
    return "Spain";
  }

  return null;
}

// =============================================================================
// 3. EXTRACT DATE FROM TEXT (Regex Fallback)
// =============================================================================
function extractDateFromText(text: string): string | null {
  // Pattern 1: DD.MM.YYYY or DD/MM/YYYY or DD-MM-YYYY
  const euroPattern = /(\d{1,2})[./-](\d{1,2})[./-](20\d{2})/g;
  // Pattern 2: YYYY-MM-DD (ISO)
  const isoPattern = /(20\d{2})-(\d{1,2})-(\d{1,2})/g;

  const dates: { date: string; index: number }[] = [];

  // Find European format dates
  let match: RegExpExecArray | null;
  while ((match = euroPattern.exec(text)) !== null) {
    // Skip if this looks like a regulation number (e.g., "2016/425")
    const before = text.substring(Math.max(0, match.index - 5), match.index);
    if (before.match(/\d+$/) || before.match(/[:/]$/)) continue;

    const day = match[1].padStart(2, "0");
    const month = match[2].padStart(2, "0");
    const year = match[3];

    // Validate month
    const monthNum = parseInt(month, 10);
    if (monthNum >= 1 && monthNum <= 12) {
      dates.push({ date: `${year}-${month}-${day}`, index: match.index });
    }
  }

  // Find ISO format dates
  while ((match = isoPattern.exec(text)) !== null) {
    const year = match[1];
    const month = match[2].padStart(2, "0");
    const day = match[3].padStart(2, "0");

    const monthNum = parseInt(month, 10);
    if (monthNum >= 1 && monthNum <= 12) {
      dates.push({ date: `${year}-${month}-${day}`, index: match.index });
    }
  }

  // Return the LAST date found (signatures/dates typically at bottom)
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

  // =========================================================================
  // FIX #1: COUNTRY INFERENCE (Check both text AND supplier name)
  // =========================================================================
  if (!data.country || data.country === "Unknown" || data.country === "") {
    const inferredCountry = inferCountry(fullText, data.supplier_name || "");
    if (inferredCountry) {
      data.country = inferredCountry;
      console.log("Country inferred:", inferredCountry);
    }
  }

  // =========================================================================
  // FIX #2: DATE FALLBACK (Regex extraction if AI missed it)
  // =========================================================================
  if (!data.date_issued || data.date_issued === "null" || data.date_issued === "") {
    const extractedDate = extractDateFromText(fullText);
    if (extractedDate) {
      data.date_issued = extractedDate;
      console.log("Date extracted via regex:", extractedDate);
    }
  }

  // =========================================================================
  // RULE 1: GLOVES (Highest Priority - "Anhui Intco" Files)
  // =========================================================================
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
    supp.includes("intco") ||
    textLower.includes("en 455") ||
    textLower.includes("en 374") ||
    textLower.includes("en 420");

  if (isGlove) {
    data.scope = "+";
    data.measure = "EU Regulation 2016/425";

    if (!data.product_category || data.product_category === "Goods" || data.product_category === "Unknown") {
      data.product_category = "Gloves";
    }

    // FIX #3: DEDUP FIX - Copy certification to certificate_number if missing
    if (!data.certificate_number || data.certificate_number === "" || data.certificate_number === "null") {
      // Extract the specific EN standard as the certificate number
      if (textLower.includes("en 455")) data.certificate_number = "EN 455";
      else if (textLower.includes("en 374")) data.certificate_number = "EN 374";
      else if (textLower.includes("en 420")) data.certificate_number = "EN 420";
      else if (textLower.includes("en 388")) data.certificate_number = "EN 388";
      else if (data.certification) data.certificate_number = data.certification;
      else data.certificate_number = "Glove Test Report";
    }

    // 3-Year Rule for test reports
    if (data.date_issued && (!data.date_expired || data.date_expired === "null" || data.date_expired === "")) {
      data.date_expired = calculateExpiry(data.date_issued);
    }

    return data;
  }

  // =========================================================================
  // RULE 2: ISO 14001 (Environmental)
  // =========================================================================
  if (cert.includes("iso 14001") || cert.includes("14001") || textLower.includes("iso 14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    if (!data.certificate_number) data.certificate_number = data.certification || "ISO 14001";
    return data;
  }

  // =========================================================================
  // RULE 3: ISO 45001 (Health & Safety)
  // =========================================================================
  if (cert.includes("iso 45001") || cert.includes("45001") || textLower.includes("iso 45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
    if (!data.certificate_number) data.certificate_number = data.certification || "ISO 45001";
    return data;
  }

  // =========================================================================
  // RULE 4: ISO 27001 (Information Security)
  // =========================================================================
  if (cert.includes("iso 27001") || cert.includes("27001")) {
    data.scope = "!";
    data.measure = "EU GDPR";
    if (!data.certificate_number) data.certificate_number = data.certification || "ISO 27001";
    return data;
  }

  // =========================================================================
  // RULE 5: FSC (Forestry)
  // =========================================================================
  if ((cert.includes("fsc") && !cert.includes("fssc")) || (textLower.includes("fsc") && !textLower.includes("fssc"))) {
    data.scope = "!";
    data.measure = "FSC";
    if (!data.certificate_number) data.certificate_number = data.certification || "FSC";
    return data;
  }

  // =========================================================================
  // RULE 6: Factory Certs (ISO 9001, BRC, BRCGS, ISO 22000, FSSC)
  // =========================================================================
  const isFactoryCert =
    cert.includes("iso 9001") ||
    cert.includes("9001") ||
    cert.includes("brc") ||
    cert.includes("brcgs") ||
    cert.includes("iso 22000") ||
    cert.includes("22000") ||
    cert.includes("fssc") ||
    cert.includes("gmp") ||
    textLower.includes("iso 9001") ||
    textLower.includes("brcgs");

  if (isFactoryCert) {
    data.scope = "!";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 2023/2006";
    }
    if (!data.certificate_number) data.certificate_number = data.certification || "Management Certificate";
    return data;
  }

  // =========================================================================
  // RULE 7: EU 10/2011 (Plastics)
  // =========================================================================
  if (cert.includes("10/2011") || textLower.includes("10/2011")) {
    data.scope = "+";
    data.measure = "(EC) No 10/2011";
    if (!data.certificate_number) data.certificate_number = data.certification || "EU 10/2011";
    return data;
  }

  // =========================================================================
  // RULE 8: EN 13432 (Compostable)
  // =========================================================================
  if (cert.includes("en 13432") || cert.includes("compostable") || textLower.includes("en 13432")) {
    data.scope = "+";
    data.measure = "EN 13432";
    if (!data.certificate_number) data.certificate_number = data.certification || "EN 13432";
    return data;
  }

  // =========================================================================
  // RULE 9: DoC / Migration / Food Contact
  // =========================================================================
  const isProductTest =
    cert.includes("declaration of conformity") ||
    cert.includes("declaration of compliance") ||
    cert.includes("doc") ||
    cert.includes("migration") ||
    cert.includes("food contact") ||
    cert.includes("1935/2004") ||
    textLower.includes("declaration of conformity") ||
    textLower.includes("migration");

  if (isProductTest) {
    data.scope = "+";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 1935/2004";
    }
    if (!data.certificate_number) data.certificate_number = data.certification || "DoC";

    // 3-Year Rule for DoC if no expiry
    if (data.date_issued && (!data.date_expired || data.date_expired === "null" || data.date_expired === "")) {
      data.date_expired = calculateExpiry(data.date_issued);
    }

    return data;
  }

  // =========================================================================
  // RULE 10: Business License
  // =========================================================================
  if (cert.includes("business license") || textLower.includes("business license")) {
    data.scope = "!";
    data.measure = "National Regulation";
    if (!data.certificate_number) data.certificate_number = data.certification || "Business License";
    return data;
  }

  // =========================================================================
  // DEFAULT: Ensure certificate_number is never empty
  // =========================================================================
  if (!data.certificate_number || data.certificate_number === "" || data.certificate_number === "null") {
    data.certificate_number = data.certification || "Certificate";
  }

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
  } catch {
    console.warn("JSON.parse failed, attempting regex extraction...");
  }

  // Attempt 2: Regex Extraction
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

  // Apply Business Logic (with full text for inference)
  data = applyBusinessLogic(data, fullInput);

  // FINAL SANITIZATION - ALL FIELDS MUST BE STRINGS
  return {
    supplier_name: sanitize(data.supplier_name, "Unknown Supplier"),
    certificate_number: sanitize(data.certificate_number, "Certificate"), // CRITICAL: Prevents dedup
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
CRITICAL: OUTPUT RAW JSON ONLY. NO MARKDOWN.

### EXTRACTION RULES
- supplier_name: Look for "Applicant", "Manufacturer", company letterhead.
- certificate_number: Extract the certificate/report number if visible.
- country: Extract from address. If unclear, return null.
- certification: The standard name (e.g., "EN 455", "ISO 9001", "BRC").
- product_category: Product description from the certificate.
- date_issued: Issue/Signed date. Format: YYYY-MM-DD. "09.10.2024" = Oct 9th (European).
- date_expired: Expiry date. If not found, return null.

### OUTPUT SCHEMA
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
        { type: "text", text: `Extract certificate data from (${fileName}):\n\n${truncatedText}` },
      ];
    } else {
      const imageUrl = rawImage.startsWith("data:") ? rawImage : `data:image/jpeg;base64,${rawImage}`;
      userContent = [
        { type: "image_url", image_url: { url: imageUrl } },
        { type: "text", text: `Extract certificate data from (${fileName}).` },
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
