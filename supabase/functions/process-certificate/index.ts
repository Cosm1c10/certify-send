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
Goal: Extract structured data for the Client Master File with strict "General" vs "Specific" classification.

### 1. CLASSIFICATION RULES (CRITICAL)

**Scope Column (The Symbol):**
- Output "!" (Exclamation Mark) IF:
  - Factory/Management System Certs: BRC, BRCGS, ISO 9001, ISO 22000, ISO 45001, ISO 14001, GMP, FSC, GRS, FSSC 22000.
  - **Declaration of Compliance (DoC)** - Per client instruction, this is GENERAL.
  - **Migration Test Reports** - Per client instruction, this is GENERAL.
- Output "+" (Plus Sign) IF:
  - Specific Product Performance Certs: Compostable (EN 13432), Recyclable (ISO 14021), Food Grade with 10/2011.

**Measure Column (The Standard Mapping):**
- IF "Declaration of Compliance" OR "Migration Report" OR "1935/2004" -> Output: "Regulation (EC) No 1935/2004"
- IF ISO 22000 OR BRC OR BRCGS OR ISO 9001 OR FSSC 22000 -> Output: "(EC) No 2023/2006"
- IF Compostable (DIN CERTCO / TUV / EN 13432) -> Output: "EN 13432 (Compostable)"
- IF Recyclable (ISO 14021) -> Output: "ISO 14021 (Recyclable)"
- IF FSC -> Output: "FSC"
- IF 10/2011 -> Output: "Commission Regulation (EU) No 10/2011"

**Product Category Column (The Description):**
- EXTRACT the detailed product description here.
- Examples: "Aqueous Coated Paper Cup", "Blanks of bio coated paper", "Disposable kraft paper cups", "PET Bottles".
- Do NOT put generic terms like "Paper" or "Plastic". Put the full product string from the certificate.

**Certification Column:** The document type or certifying body.
- Examples: "BRCGS", "ISO 9001", "ISO 22000", "FSSC 22000", "DIN CERTCO", "TUV", "Migration Test Report", "Declaration of Compliance", "FSC Cert"

**Supplier Name:** Normalize to simple name.
- "Safira Amb.", "SAFİRA AMBALAJ" -> "Safira Ambalaj"
- "Huhtamaki Turkey" -> "Huhtamaki"
- Remove legal suffixes: "San. Ve Tic. A.Ş.", "Co., Ltd", "Ltd. Şti."

**Country:** Detect from address.
- "Istanbul", "Türkiye" -> "Turkey"
- "China", "Changsha" -> "China"

**Certificate Number:** Extract from "Report No", "Certificate No", "Registration No".

**Dates (YYYY-MM-DD):** "11.03.2019" = March 11, 2019 (DD.MM.YYYY format)

### 2. OUTPUT JSON FORMAT
{
  "supplier_name": "string",
  "certificate_number": "string",
  "country": "string",
  "scope": "! or +",
  "measure": "string (Mapped Regulation)",
  "certification": "string (Document Type)",
  "product_category": "string (Product Description)",
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
