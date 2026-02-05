// Supabase Edge Function - FINAL PRODUCTION (Sanitization + Business Logic)
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
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
};

// 1. DATA SANITIZER (Prevents Excel Drops)
// Converts null/undefined to "" so the row NEVER disappears.
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined || val === "null") return defaultVal;
  return String(val).trim();
}

// 2. HARD-CODED BUSINESS LOGIC (Fixes Empty Measures)
function applyBusinessLogic(data: any): any {
  const cert = (data.certification || "").toLowerCase();
  const prod = (data.product_category || "").toLowerCase();
  const measure = (data.measure || "").toLowerCase();

  // RULE A: GLOVES (The "Force" Rule)
  if (cert.includes("en 455") || cert.includes("en 374") || cert.includes("en 420") || cert.includes("glove") || prod.includes("glove")) {
    data.scope = "+";
    // Always overwrite Measure for gloves, even if AI guessed something else
    data.measure = "EU Regulation 2016/425";
    if (!data.product_category || data.product_category === "Goods") data.product_category = "Gloves";
  }

  // RULE B: FACTORY CERTS (ISO 9001, BRC, FSSC)
  else if (cert.includes("iso 9001") || cert.includes("brc") || cert.includes("iso 22000") || cert.includes("fssc")) {
    data.scope = "!";
    // If measure is blank or generic, force the specific Regulation
    if (!data.measure || measure === "national regulation" || measure === "") {
      data.measure = "(EC) No 2023/2006";
    }
  }

  // RULE C: ISO 14001
  else if (cert.includes("iso 14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
  }

  // RULE D: ISO 45001
  else if (cert.includes("iso 45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
  }

  // RULE E: DoC (Declaration of Compliance) -> Specific
  else if (cert.includes("declaration of conformity") || cert.includes("doc")) {
    data.scope = "+";
    if (!data.measure) data.measure = "(EC) No 1935/2004";
  }

  return data;
}

// 3. ROBUST EXTRACTION (Self-Healing)
function extractJSON(text: string): any {
  console.log("Raw AI Response:", text);
  let cleanText = text.replace(/```json/g, "").replace(/```/g, "").trim();

  let data: any = {};

  // Attempt 1: Standard JSON Parse
  try {
    const first = cleanText.indexOf('{');
    const last = cleanText.lastIndexOf('}');
    if (first !== -1 && last !== -1) {
      data = JSON.parse(cleanText.substring(first, last + 1));
    }
  } catch (e) {
    console.warn("JSON Parse failed, trying regex...");
  }

  // Attempt 2: Regex Fallback (Scavenge missing fields)
  const scavenge = (key: string) => {
    const match = cleanText.match(new RegExp(`"${key}"\\s*:\\s*"([^"]*)"`, "i"));
    return match ? match[1] : null;
  };

  if (!data.supplier_name) data.supplier_name = scavenge("supplier_name");
  if (!data.country) data.country = scavenge("country");
  if (!data.certification) data.certification = scavenge("certification");
  if (!data.measure) data.measure = scavenge("measure");
  if (!data.scope) data.scope = scavenge("scope");
  if (!data.product_category) data.product_category = scavenge("product_category");
  if (!data.date_issued) data.date_issued = scavenge("date_issued");
  if (!data.date_expired) data.date_expired = scavenge("date_expired");

  // APPLY THE HARD LOGIC (Overwrite bad AI guesses)
  data = applyBusinessLogic(data);

  // 4. FINAL SANITIZATION (Critical for Excel Export)
  return {
    supplier_name: sanitize(data.supplier_name, "Unknown Supplier"),
    country: sanitize(data.country, "Unknown"),
    scope: sanitize(data.scope, "!"),
    measure: sanitize(data.measure, "National Regulation"),
    certification: sanitize(data.certification, "Standard"),
    product_category: sanitize(data.product_category, "Goods"),
    date_issued: sanitize(data.date_issued, ""),
    date_expired: sanitize(data.date_expired, ""),
    status: "Success"
  };
}

// 5. MASTER PROMPT
const systemPrompt = `
You are a Compliance Data Extraction Engine.
CRITICAL: OUTPUT RAW JSON ONLY.

### LOGIC RULES
- **Gloves:** Treat EN 455, EN 374, EN 420 as VALID. Supplier = Applicant.
- **Dates:** Format YYYY-MM-DD. If no expiry, calculate **Issue Date + 3 Years**.
- **European Dates:** "09.10.2024" = October 9th. NEVER September.

### OUTPUT SCHEME
{
  "supplier_name": "string",
  "country": "string",
  "scope": "string",
  "measure": "string",
  "certification": "string",
  "product_category": "string",
  "date_issued": "YYYY-MM-DD",
  "date_expired": "YYYY-MM-DD"
}
`;

// 6. MAX INPUT SIZE
const MAX_TEXT_LENGTH = 50000;

serve(async (req) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { image, text, filename } = await req.json();

    // Require either image or text
    if (!image && !text) {
      return new Response(
        JSON.stringify({ error: "Missing 'image' or 'text' field" }),
        { status: 400, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const isTextMode = !!text;
    console.log("Processing:", filename || "unknown", isTextMode ? "(text)" : "(image)");

    const openaiApiKey = Deno.env.get("OPENAI_API_KEY");
    if (!openaiApiKey) {
      return new Response(
        JSON.stringify({ error: "OpenAI API key not configured" }),
        { status: 500, headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    const openai = new OpenAI({ apiKey: openaiApiKey });

    // Build user message content
    let userContent: any[];

    if (isTextMode) {
      // TRUNCATE to prevent timeouts
      const truncatedText = text.length > MAX_TEXT_LENGTH
        ? text.slice(0, MAX_TEXT_LENGTH) + "\n\n[TRUNCATED]"
        : text;

      console.log(`Text: ${text.length} chars, truncated: ${truncatedText.length}`);

      userContent = [
        { type: "text", text: `Extract data from (${filename || "unknown"}):\n\n${truncatedText}` },
      ];
    } else {
      // Image mode
      const imageContent = image.startsWith("data:") ? image : `data:image/jpeg;base64,${image}`;

      userContent = [
        { type: "image_url", image_url: { url: imageContent } },
        { type: "text", text: `Extract data from (${filename || "unknown"}).` },
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
      console.error("No response from OpenAI");
      return new Response(
        JSON.stringify({
          supplier_name: "Error: No AI Response",
          country: "Unknown",
          scope: "!",
          measure: "Manual Review",
          certification: "AI returned empty",
          product_category: "Unknown",
          date_issued: "",
          date_expired: "",
          status: "Error"
        }),
        { headers: { ...corsHeaders, "Content-Type": "application/json" } }
      );
    }

    // Extract + Sanitize + Apply Business Logic
    const result = extractJSON(rawContent);

    return new Response(JSON.stringify(result), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });

  } catch (error) {
    console.error("Processing Error:", error);

    // FAILSAFE RETURN (Fully Sanitized)
    return new Response(
      JSON.stringify({
        supplier_name: "Error: Processing Failed",
        country: "Unknown",
        scope: "!",
        measure: "Manual Review",
        certification: "File too complex",
        product_category: "Unknown",
        date_issued: "",
        date_expired: "",
        status: "Error"
      }),
      { status: 200, headers: { ...corsHeaders, "Content-Type": "application/json" } }
    );
  }
});
