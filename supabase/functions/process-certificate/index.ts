// Supabase Edge Function - runs in Deno runtime
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

serve(async (req) => {
  // Handle CORS preflight requests
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { image, filename } = await req.json();

    if (!image) {
      return new Response(
        JSON.stringify({ error: "Missing 'image' field in request body" }),
        {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    console.log("Processing certificate with filename:", filename || "not provided");

    const openaiApiKey = Deno.env.get("OPENAI_API_KEY");
    if (!openaiApiKey) {
      return new Response(
        JSON.stringify({ error: "OpenAI API key not configured" }),
        {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    const openai = new OpenAI({ apiKey: openaiApiKey });

    const systemPrompt = `
You are a Compliance Data Extraction Engine.
Goal: Extract structured data for the Client Master File.

### 1. EXTRACTION RULES

**Supplier Name:** Normalize company names.
- "Safira Amb.", "SAFİRA AMBALAJ", "Safira Ambalaj San. Ve Tic." -> "Safira Ambalaj"
- "Huhtamaki Turkey", "Huhtamaki" -> "Huhtamaki"
- Remove legal suffixes: "San. Ve Tic. A.Ş.", "Co., Ltd", "Ltd. Şti.", "Pvt. Ltd"

**Country:** Detect from address block.
- "Istanbul", "Turkey", "Türkiye" -> "Turkey"
- "China", "Changsha", "Hunan" -> "China"
- "Dublin", "Ireland" -> "Ireland"
- "Germany", "Deutschland" -> "Germany"
- "Poland", "Polska" -> "Poland"

**Scope:** Short summary of the product covered.
- Examples: "Aqueous Coated Paper Cup", "PET Bottles", "Food Contact Materials", "Single Wall Cup"

**Measure (CRITICAL - DO NOT DEFAULT TO "General Compliance"):**
- ALWAYS look for specific standards. "General Compliance" is ONLY acceptable if NO standard is found.
- If "EN 13432" or "DIN CERTCO" found -> "EN 13432 (Compostable)"
- If "10/2011" found -> "Commission Regulation (EU) No 10/2011"
- If "2023/2006" found -> "Commission Regulation (EC) No 2023/2006"
- If "1935/2004" found -> "Regulation (EC) No 1935/2004"
- If "94/62/EC" found -> "Directive 94/62/EC"
- If "ISO 14021" found -> "ISO 14021 (Recyclable)"
- If "BRC" or "BRCGS" found -> "BRCGS Global Standard"
- If "FSC" found -> "FSC Standard"
- If migration/food contact test -> "Migration Test"
- ONLY use "General Compliance" as absolute last resort when NO standard numbers exist.

**Certification:** The document type or certifying body.
- Examples: "BRCGS", "ISO 9001", "ISO 22000", "FSSC 22000", "DIN CERTCO", "Migration Test Report", "Declaration of Compliance", "Recyclability Certificate", "FSC Cert", "Halal", "Kosher"

**Product Category:** Material classification.
- Options: "Paper", "Rigid Plastics", "Flexible Plastics", "Chemicals", "Metal", "Glass", "Wood"

**Certificate Number:** Extract from "Report No", "Rapor No", "Certificate No", "Registration No".

**Dates (Format: YYYY-MM-DD):**
- Extract "Issue Date" (or "Tarih") and "Expiry Date" (or "Valid until")
- Note: "11.03.2019" = March 11, 2019 (DD.MM.YYYY format)

### 2. OUTPUT JSON FORMAT
{
  "supplier_name": "string",
  "certificate_number": "string",
  "country": "string",
  "scope": "string",
  "measure": "string",
  "certification": "string",
  "product_category": "string",
  "date_issued": "YYYY-MM-DD",
  "date_expired": "YYYY-MM-DD or null"
}
`;

    // Determine if the image is a data URL or raw base64
    const imageContent = image.startsWith("data:")
      ? image
      : `data:image/jpeg;base64,${image}`;

    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "system",
          content: systemPrompt,
        },
        {
          role: "user",
          content: [
            {
              type: "image_url",
              image_url: {
                url: imageContent,
              },
            },
            {
              type: "text",
              text: `Analyze this certificate. Filename: "${filename || "unknown.pdf"}". Extract all certificate information.`,
            },
          ],
        },
      ],
      max_tokens: 1000,
    });

    const content = response.choices[0]?.message?.content;

    if (!content) {
      return new Response(
        JSON.stringify({ error: "No response from OpenAI" }),
        {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    // Parse the JSON response from GPT-4o
    // Remove potential markdown code blocks if present
    const cleanedContent = content
      .replace(/```json\n?/g, "")
      .replace(/```\n?/g, "")
      .trim();

    const extractedData = JSON.parse(cleanedContent);

    return new Response(JSON.stringify(extractedData), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (error) {
    console.error("Error processing certificate:", error);

    return new Response(
      JSON.stringify({
        error: "Failed to process certificate",
        details: error instanceof Error ? error.message : "Unknown error",
      }),
      {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      }
    );
  }
});
