// =============================================================================
// SUPABASE EDGE FUNCTION: process-certificate
// Version: FINAL PRODUCTION (Feb 2026)
// Client: Catering Disposables (Jun & Saurebh)
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
// 1. SANITIZE FUNCTION (Fixes Excel Drop Bug)
// =============================================================================
// Converts null/undefined/"null" to empty string so Excel rows never disappear.
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined || val === "null" || val === "undefined") {
    return defaultVal;
  }
  return String(val).trim();
}

// =============================================================================
// 2. APPLY BUSINESS LOGIC (Jun & Saurebh's Rules - Feb 2026)
// =============================================================================
// This function OVERWRITES AI guesses with hard-coded regulatory mappings.
function applyBusinessLogic(data: any): any {
  const cert = (data.certification || "").toLowerCase();
  const prod = (data.product_category || "").toLowerCase();
  const supp = (data.supplier_name || "").toLowerCase();
  const currentMeasure = (data.measure || "").toLowerCase();

  // -------------------------------------------------------------------------
  // RULE C: GLOVES (Saurebh's Requirement) - HIGHEST PRIORITY
  // -------------------------------------------------------------------------
  // Triggers: EN 455, EN 374, EN 420, Module B, Cat III, Nitrile Gloves
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

  // -------------------------------------------------------------------------
  // RULE A: GENERAL (!) - Factory/Management Certificates
  // -------------------------------------------------------------------------

  // A.1: ISO 14001 (Environmental)
  if (cert.includes("iso 14001") || cert.includes("14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
    return data;
  }

  // A.2: ISO 45001 (Occupational Health & Safety)
  if (cert.includes("iso 45001") || cert.includes("45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
    return data;
  }

  // A.3: ISO 27001 (Information Security)
  if (cert.includes("iso 27001") || cert.includes("27001")) {
    data.scope = "!";
    data.measure = "EU GDPR";
    return data;
  }

  // A.4: FSC (Forestry)
  if (cert.includes("fsc") && !cert.includes("fssc")) {
    data.scope = "!";
    data.measure = "FSC";
    return data;
  }

  // A.5: ISO 9001 / BRC / BRCGS / ISO 22000 / FSSC 22000 / GMP
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

  // -------------------------------------------------------------------------
  // RULE B: SPECIFIC (+) - Product Compliance Certificates
  // -------------------------------------------------------------------------

  // B.1: EU 10/2011 (Plastics Regulation)
  if (cert.includes("10/2011") || cert.includes("eu 10/2011") || cert.includes("plastic")) {
    data.scope = "+";
    data.measure = "(EC) No 10/2011";
    return data;
  }

  // B.2: EN 13432 (Compostable)
  if (cert.includes("en 13432") || cert.includes("13432") || cert.includes("compostable") || cert.includes("din certco")) {
    data.scope = "+";
    data.measure = "EN 13432";
    return data;
  }

  // B.3: EN 13430 / ISO 14021 (Recyclable)
  if (cert.includes("en 13430") || cert.includes("13430") || cert.includes("iso 14021") || cert.includes("recyclable")) {
    data.scope = "+";
    data.measure = "EN 13430";
    return data;
  }

  // B.4: DoC / Migration Reports / Heavy Metal / Food Contact Tests
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

  // -------------------------------------------------------------------------
  // DEFAULT: Business License / Unknown
  // -------------------------------------------------------------------------
  if (cert.includes("business license") || cert.includes("operating permit")) {
    data.scope = "!";
    data.measure = "National Regulation";
    return data;
  }

  // If nothing matched, default to General
  if (!data.scope) data.scope = "!";
  if (!data.measure) data.measure = "National Regulation";

  return data;
}

// =============================================================================
// 3. CALCULATE EXPIRY DATE (3-Year Rule for Test Reports)
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

  // Clean markdown code blocks
  let cleanText = text.replace(/```json/gi, "").replace(/```/g, "").trim();

  let data: any = {};

  // -------------------------------------------------------------------------
  // Attempt 1: Standard JSON Parse
  // -------------------------------------------------------------------------
  try {
    const firstBrace = cleanText.indexOf("{");
    const lastBrace = cleanText.lastIndexOf("}");
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      const jsonString = cleanText.substring(firstBrace, lastBrace + 1);
      data = JSON.parse(jsonString);
    }
  } catch (e) {
    console.warn("JSON.parse failed, attempting regex extraction...");
  }

  // -------------------------------------------------------------------------
  // Attempt 2: Regex Extraction (Scavenge Fields)
  // -------------------------------------------------------------------------
  const scavenge = (key: string): string | null => {
    // Match "key": "value" or "key": null
    const regex = new RegExp(`"${key}"\\s*:\\s*(?:"([^"]*)"|null)`, "i");
    const match = cleanText.match(regex);
    return match && match[1] ? match[1] : null;
  };

  // Fill in missing fields
  if (!data.supplier_name) data.supplier_name = scavenge("supplier_name");
  if (!data.country) data.country = scavenge("country");
  if (!data.scope) data.scope = scavenge("scope");
  if (!data.measure) data.measure = scavenge("measure");
  if (!data.certification) data.certification = scavenge("certification");
  if (!data.product_category) data.product_category = scavenge("product_category");
  if (!data.date_issued) data.date_issued = scavenge("date_issued");
  if (!data.date_expired) data.date_expired = scavenge("date_expired");

  // -------------------------------------------------------------------------
  // Apply Business Logic (Overwrite bad AI guesses)
  // -------------------------------------------------------------------------
  data = applyBusinessLogic(data);

  // -------------------------------------------------------------------------
  // Apply 3-Year Rule for Test Reports without Expiry
  // -------------------------------------------------------------------------
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

  // -------------------------------------------------------------------------
  // Final Sanitization (All fields become strings, never null)
  // -------------------------------------------------------------------------
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
// 5. SYSTEM PROMPT (First Layer of Defense)
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
// 6. CONFIGURATION
// =============================================================================
const MAX_INPUT_LENGTH = 50000; // Prevent timeouts

// =============================================================================
// 7. MAIN HANDLER
// =============================================================================
serve(async (req) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    // -------------------------------------------------------------------------
    // Parse Input (Support multiple payload formats)
    // -------------------------------------------------------------------------
    const input = await req.json();

    // Support: { text, filename } or { fileContent, fileName } or { image, filename }
    const rawText = input.text || input.fileContent || "";
    const rawImage = input.image || "";
    const fileName = input.filename || input.fileName || "Unknown File";

    const isTextMode = !!rawText && !rawImage;
    const isImageMode = !!rawImage;

    if (!rawText && !rawImage) {
      return new Response(
        JSON.stringify({ error: "Missing 'text', 'fileContent', or 'image' field" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    console.log(`Processing: ${fileName} | Mode: ${isTextMode ? "TEXT" : "IMAGE"} | Length: ${rawText.length || "N/A"}`);

    // -------------------------------------------------------------------------
    // Initialize OpenAI
    // -------------------------------------------------------------------------
    const openaiApiKey = Deno.env.get("OPENAI_API_KEY");
    if (!openaiApiKey) {
      return new Response(
        JSON.stringify({ error: "OpenAI API key not configured" }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const openai = new OpenAI({ apiKey: openaiApiKey });

    // -------------------------------------------------------------------------
    // Build Request
    // -------------------------------------------------------------------------
    let userContent: any[];

    if (isTextMode) {
      // Truncate to prevent timeouts
      const truncatedText =
        rawText.length > MAX_INPUT_LENGTH
          ? rawText.slice(0, MAX_INPUT_LENGTH) + "\n\n[TRUNCATED - File exceeded 50,000 characters]"
          : rawText;

      userContent = [
        {
          type: "text",
          text: `Extract compliance data from this file (${fileName}):\n\n${truncatedText}`,
        },
      ];
    } else {
      // Image mode
      const imageUrl = rawImage.startsWith("data:") ? rawImage : `data:image/jpeg;base64,${rawImage}`;

      userContent = [
        { type: "image_url", image_url: { url: imageUrl } },
        { type: "text", text: `Extract compliance data from this certificate (${fileName}).` },
      ];
    }

    // -------------------------------------------------------------------------
    // Call OpenAI
    // -------------------------------------------------------------------------
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
      console.error("OpenAI returned empty response");
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

    // -------------------------------------------------------------------------
    // Extract, Apply Logic, Sanitize
    // -------------------------------------------------------------------------
    const result = extractJSON(rawContent);

    return new Response(JSON.stringify(result), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });

    // -------------------------------------------------------------------------
  } catch (error) {
    console.error("Processing Error:", error);

    // FAILSAFE: Return sanitized error object (never crashes frontend)
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
      {
        status: 200, // Return 200 so frontend doesn't crash
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      }
    );
  }
});
