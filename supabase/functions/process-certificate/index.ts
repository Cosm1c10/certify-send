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

/**
 * Robust JSON extraction from AI response
 * Handles cases where AI outputs conversational text before/after JSON
 */
function extractJSON(text: string): any {
  // First, clean markdown code blocks
  const cleaned = text
    .replace(/```json\n?/g, "")
    .replace(/```\n?/g, "")
    .trim();

  try {
    // Attempt 1: Direct parse of cleaned text
    return JSON.parse(cleaned);
  } catch (e) {
    // Attempt 2: Extract JSON object from wrapper text
    const jsonMatch = cleaned.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      try {
        return JSON.parse(jsonMatch[0]);
      } catch (e2) {
        console.error("Failed to parse extracted JSON:", e2);
        throw new Error("Invalid JSON format in AI response");
      }
    }
    console.error("No JSON found in response:", cleaned.substring(0, 200));
    throw new Error("No JSON found in AI response");
  }
}

serve(async (req) => {
  // Handle CORS preflight requests
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const { image, text, filename } = await req.json();

    // Require either image or text
    if (!image && !text) {
      return new Response(
        JSON.stringify({ error: "Missing 'image' or 'text' field in request body" }),
        {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    const isTextMode = !!text;
    console.log("Processing certificate with filename:", filename || "not provided", isTextMode ? "(text mode)" : "(image mode)");

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
CRITICAL: OUTPUT ONLY RAW JSON. NO MARKDOWN. NO CONVERSATIONAL TEXT. NO EXPLANATIONS.

### 1. DOCUMENT HANDLING RULES
- **Standard Certificates:** Extract normally.
- **Business Licenses / Operating Permits:** Treat as valid documents.
  - Certification: "Business License" or "Operating Permit"
  - Measure: "National Regulation"
  - Scope: "!"
- **Multi-Language Documents:** If same certificate appears in multiple languages (e.g., Chinese + English), merge into ONE record.

### 2. CLASSIFICATION RULES

**Scope Column:**
- "!" (General) for: Factory Certs, Business Licenses, DoC, Migration Reports, ISO, BRC, FSC.
- "+" (Specific) for: Product Certs (Compostable EN 13432, Recyclable ISO 14021, Food Grade 10/2011).

**Measure Column:**
- DoC / Migration / 1935/2004 -> "Regulation (EC) No 1935/2004"
- ISO 22000 / BRC / BRCGS / ISO 9001 / FSSC 22000 -> "(EC) No 2023/2006"
- Compostable (DIN CERTCO / TUV / EN 13432) -> "EN 13432 (Compostable)"
- Recyclable (ISO 14021) -> "ISO 14021 (Recyclable)"
- FSC -> "FSC"
- 10/2011 -> "Commission Regulation (EU) No 10/2011"
- Business License -> "National Regulation"

**Product Category:** Full description from cert (e.g., "Aqueous Coated Paper Cup", "PET Bottles").

**Supplier Name:** Normalize (remove "San. Ve Tic. A.Ş.", "Co., Ltd", etc.).

**Country:** "Istanbul"/"Türkiye" -> "Turkey", "Changsha" -> "China".

**Dates:** DD.MM.YYYY format -> YYYY-MM-DD (e.g., "11.03.2019" = "2019-03-11").

### 3. OUTPUT JSON (RAW, NO WRAPPER)
{
  "supplier_name": "string",
  "certificate_number": "string",
  "country": "string",
  "scope": "! or +",
  "measure": "string",
  "certification": "string",
  "product_category": "string",
  "date_issued": "YYYY-MM-DD",
  "date_expired": "YYYY-MM-DD or null"
}
`;

    // Build user message content based on input type
    let userContent: any[];

    if (isTextMode) {
      // Text mode for DOCX files
      userContent = [
        {
          type: "text",
          text: `Analyze this certificate document. Filename: "${filename || "unknown.docx"}". Extract all certificate information from the following text:\n\n${text}`,
        },
      ];
    } else {
      // Image mode for PDF/Image files
      const imageContent = image.startsWith("data:")
        ? image
        : `data:image/jpeg;base64,${image}`;

      userContent = [
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
      ];
    }

    const response = await openai.chat.completions.create({
      model: "gpt-4o",
      messages: [
        {
          role: "system",
          content: systemPrompt,
        },
        {
          role: "user",
          content: userContent,
        },
      ],
      max_tokens: 1000,
    });

    const rawContent = response.choices[0]?.message?.content;

    if (!rawContent) {
      return new Response(
        JSON.stringify({ error: "No response from OpenAI" }),
        {
          status: 500,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

    // Debug log for troubleshooting
    console.log("Raw AI Response (first 500 chars):", rawContent.substring(0, 500));

    // Parse the JSON response using robust extraction
    const extractedData = extractJSON(rawContent);

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
