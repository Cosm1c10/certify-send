// =============================================================================
// SUPABASE EDGE FUNCTION: process-certificate
// Version: FINAL PRODUCTION v3.1 (OpenAI GPT-4o + JSON Mode)
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
  const combined = (fullText + " " + supplierName).toLowerCase();

  // CHINA
  if (
    combined.includes("china") || combined.includes("anhui") || combined.includes("guangdong") ||
    combined.includes("shanghai") || combined.includes("beijing") || combined.includes("shenzhen") ||
    combined.includes("changsha") || combined.includes("zhejiang") || combined.includes("jiangsu") ||
    combined.includes("hunan") || combined.includes("wenzhou") || combined.includes("fujian") ||
    combined.includes("shandong") || combined.includes("hebei") || combined.includes("henan") ||
    combined.includes("shaoneng") || combined.includes("intco") ||
    combined.match(/\bcn\b/) ||
    combined.match(/[\u4e00-\u9fff]/)
  ) {
    return "China";
  }

  // TURKEY
  if (
    combined.includes("turkey") || combined.includes("türkiye") || combined.includes("turkiye") ||
    combined.includes("istanbul") || combined.includes("ankara") || combined.includes("izmir") ||
    combined.includes("boran") || combined.includes("mopack") || combined.includes("san. ve tic")
  ) {
    return "Turkey";
  }

  // GERMANY
  if (combined.includes("germany") || combined.includes("deutschland") || combined.includes("gmbh") ||
      combined.includes("munich") || combined.includes("berlin")) {
    return "Germany";
  }

  // UK
  if (combined.includes("united kingdom") || combined.includes("england") || combined.includes("london") ||
      combined.match(/\buk\b/) || combined.match(/\bgb\b/)) {
    return "UK";
  }

  // IRELAND
  if (combined.includes("ireland") || combined.includes("dublin") || combined.includes("cork")) {
    return "Ireland";
  }

  // USA
  if (combined.includes("united states") || combined.includes("usa") || combined.includes("california")) {
    return "USA";
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

  // =========================================================================
  // FIX #1: COUNTRY INFERENCE
  // =========================================================================
  if (!data.country || data.country === "Unknown" || data.country === "") {
    const inferredCountry = inferCountry(fullText, data.supplier_name || "");
    if (inferredCountry) {
      data.country = inferredCountry;
      console.log("Country inferred:", inferredCountry);
    }
  }

  // =========================================================================
  // FIX #2: DATE FALLBACK
  // =========================================================================
  if (!data.date_issued || data.date_issued === "null" || data.date_issued === "") {
    const extractedDate = extractDateFromText(fullText);
    if (extractedDate) {
      data.date_issued = extractedDate;
      console.log("Date extracted via regex:", extractedDate);
    }
  }

  // =========================================================================
  // RULE 1: GLOVES (Highest Priority)
  // =========================================================================
  const isGlove =
    cert.includes("en 455") || cert.includes("en 374") || cert.includes("en 420") ||
    cert.includes("en 388") || cert.includes("module b") || cert.includes("cat iii") ||
    cert.includes("category iii") || cert.includes("2016/425") ||
    prod.includes("glove") || prod.includes("nitrile") || supp.includes("intco") ||
    textLower.includes("en 455") || textLower.includes("en 374") || textLower.includes("en 420");

  if (isGlove) {
    data.scope = "+";
    data.measure = "EU Regulation 2016/425";

    if (!data.product_category || data.product_category === "Goods" || data.product_category === "Unknown") {
      data.product_category = "Gloves";
    }

    // FIX #3: DEDUP - Copy certification to certificate_number
    if (!data.certificate_number || data.certificate_number === "" || data.certificate_number === "null") {
      if (textLower.includes("en 455")) data.certificate_number = "EN 455";
      else if (textLower.includes("en 374")) data.certificate_number = "EN 374";
      else if (textLower.includes("en 420")) data.certificate_number = "EN 420";
      else if (textLower.includes("en 388")) data.certificate_number = "EN 388";
      else if (data.certification) data.certificate_number = data.certification;
      else data.certificate_number = "Glove Test Report";
    }

    // 3-Year Rule
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
  // RULE 4: ISO 27001
  // =========================================================================
  if (cert.includes("iso 27001") || cert.includes("27001")) {
    data.scope = "!";
    data.measure = "EU GDPR";
    if (!data.certificate_number) data.certificate_number = data.certification || "ISO 27001";
    return data;
  }

  // =========================================================================
  // RULE 5: FSC
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
  // RULE 9: DoC / Migration
  // =========================================================================
  const isProductTest =
    cert.includes("declaration of conformity") || cert.includes("declaration of compliance") ||
    cert.includes("doc") || cert.includes("migration") || cert.includes("food contact") ||
    cert.includes("1935/2004") ||
    textLower.includes("declaration of conformity") || textLower.includes("migration");

  if (isProductTest) {
    data.scope = "+";
    if (!data.measure || currentMeasure === "" || currentMeasure === "national regulation") {
      data.measure = "(EC) No 1935/2004";
    }
    if (!data.certificate_number) data.certificate_number = data.certification || "DoC";

    // 3-Year Rule for DoC
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
  // DEFAULT
  // =========================================================================
  if (!data.certificate_number || data.certificate_number === "" || data.certificate_number === "null") {
    data.certificate_number = data.certification || "Certificate";
  }
  if (!data.scope) data.scope = "!";
  if (!data.measure) data.measure = "National Regulation";

  return data;
}

// =============================================================================
// 6. PROCESS AI RESPONSE
// =============================================================================
function processAIResponse(rawJSON: any, fullInput: string): any {
  let data = rawJSON || {};

  // Apply Business Logic
  data = applyBusinessLogic(data, fullInput);

  // FINAL SANITIZATION
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
// 7. SYSTEM PROMPT
// =============================================================================
const systemPrompt = `You are a Compliance Data Extraction Engine. Extract certificate data and return ONLY valid JSON.

EXTRACTION RULES:
- supplier_name: Find "Applicant", "Manufacturer", or company name from letterhead
- certificate_number: The certificate/report number if visible
- country: Extract from address
- certification: Standard name (e.g., "EN 455", "ISO 9001", "BRC")
- product_category: Product description
- date_issued: Issue date as YYYY-MM-DD. European format: "09.10.2024" = October 9th
- date_expired: Expiry date as YYYY-MM-DD, or null if not found

CLASSIFICATION:
- Gloves (EN 455/374/420): scope="+", measure="EU Regulation 2016/425"
- ISO 9001/BRC/GMP: scope="!", measure="(EC) No 2023/2006"
- DoC/Migration: scope="+", measure="(EC) No 1935/2004"

Return this JSON structure:
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
}`;

// =============================================================================
// 8. CONFIGURATION (OPTIMIZED FOR SPEED)
// =============================================================================
const MAX_INPUT_LENGTH = 15000; // Reduced from 50k - faster processing
const OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";
const MODEL_TEXT = "gpt-4o-mini"; // Fast model for text extraction
const MODEL_IMAGE = "gpt-4o";     // Vision model for images

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

    // =========================================================================
    // GET OPENAI API KEY
    // =========================================================================
    const OPENAI_API_KEY = Deno.env.get("OPENAI_API_KEY");
    if (!OPENAI_API_KEY) {
      return new Response(
        JSON.stringify({ error: "OPENAI_API_KEY not configured" }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // =========================================================================
    // BUILD REQUEST
    // =========================================================================
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

    // =========================================================================
    // CALL OPENAI WITH JSON MODE (Model selected based on input type)
    // =========================================================================
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
        max_tokens: 800,  // Reduced - we only need ~300-500 for JSON output
        temperature: 0.1, // Lower = faster, more deterministic
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

    console.log("Raw AI Response:", rawContent?.substring(0, 500));

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

    // =========================================================================
    // PARSE JSON RESPONSE (JSON Mode guarantees valid JSON)
    // =========================================================================
    let parsedData: any;
    try {
      parsedData = JSON.parse(rawContent);
    } catch (e) {
      console.error("JSON Parse Error (unexpected with JSON mode):", e);
      parsedData = {};
    }

    // Process with business logic and sanitization
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
