import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import OpenAI from "https://deno.land/x/openai@v4.20.1/mod.ts";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers":
    "authorization, x-client-info, apikey, content-type",
};

interface CertificateExtractionResult {
  supplier_name: string;
  country: string;
  product_category: string;
  ec_regulation: string;
  certification: string;
  date_issued: string;
  date_expired: string;
}

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

    console.log('Processing certificate with filename:', filename || 'not provided');

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
You are a Compliance Classification Engine.
YOUR GOAL: Classify the certificate into the Client's specific "Measure" buckets.

### 1. LOGIC MAPPING RULES (CRITICAL)

**FIELD: supplier_name (PRIORITY RULES)**
- PRIORITY 1 (Filename Rule): Look at the provided 'Filename' in the user message.
  - If the filename follows the pattern 'Name - ...' (e.g., 'Ahcof - Compostable cert.pdf'), EXTRACT 'Ahcof' as the supplier.
  - If the filename starts with a Company Name before a dash or separator, use that name.
- PRIORITY 2 (Document Rule): Only if the filename is generic (like 'scan.pdf', 'document.pdf', or just numbers), extract the 'Holder' or 'Manufacturer' from the certificate image.

**FIELD: ec_regulation (The "Measure" Bucket)**
Look at the document text and type. You MUST output one of the exact strings below. Do not output the text on the page.

* IF Cert is **BRC**, **BRCGS**, **ISO 22000**, or **GMP**
    -> OUTPUT: "Commission Regulation (EC) No 2023/2006"

* IF Cert is **Compostable**, **EN 13432**, **TUV Austria**, or **DIN CERTCO**
    -> OUTPUT: "Compostable Certification"

* IF Cert is **Recyclable**, **Cyclos**, **CHI**, or **EN 13430**
    -> OUTPUT: "Recyclable Certification"

* IF Cert is **ISO 9001**, **ISO 14001**, **FSC**
    -> OUTPUT: "General Measure"

* IF Cert mentions **1935/2004** explicitly (and is not BRC/Compostable)
    -> OUTPUT: "Regulation (EC) No 1935/2004"

* IF Cert mentions **10/2011** (Plastics)
    -> OUTPUT: "Commission Regulation (EU) No 10/2011"

* ELSE (Fallback)
    -> OUTPUT: "Unknown Measure"

**FIELD: certification**
* Extract the specific standard name (e.g., "BRCGS", "DIN CERTCO", "Cyclos-HTP").

**FIELD: country**
* Extract ONLY the Country name from the site address.

**FIELD: product_category**
* Brief description of the product.

**FIELD: date_issued**
* Format: YYYY-MM-DD.

**FIELD: date_expired**
* Format: YYYY-MM-DD.

### 2. RETURN JSON
{
  "supplier_name": "string",
  "country": "string",
  "product_category": "string",
  "ec_regulation": "string",
  "certification": "string",
  "date_issued": "YYYY-MM-DD",
  "date_expired": "YYYY-MM-DD"
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
              text: `Analyze this certificate. Filename: "${filename || 'unknown.pdf'}". Extract all certificate information.`,
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

    const extractedData: CertificateExtractionResult = JSON.parse(cleanedContent);

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
