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
    const { image } = await req.json();

    if (!image) {
      return new Response(
        JSON.stringify({ error: "Missing 'image' field in request body" }),
        {
          status: 400,
          headers: { ...corsHeaders, "Content-Type": "application/json" },
        }
      );
    }

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
You are a Compliance Officer for a Packaging Company.
YOUR GOAL: Extract data to match the "Certificate Management Master Sheet".

### 1. EXTRACTION RULES (STRICT)

**FIELD: supplier_name**
- Extract the Legal Manufacturer/Holder.
- If the certificate lists a "Trading Company" (like AHCOF) AND a "Site" (like Zhongyin), extract the SITE Name as the Supplier.

**FIELD: ec_regulation (The "Measure")**
- You must classify the document into one of the Client's Standard Measures.
- IF text contains "1935/2004" -> Output: "Regulation (EC) No 1935/2004"
- IF text contains "2023/2006" or "GMP" -> Output: "Commission Regulation (EC) No 2023/2006"
- IF text contains "10/2011" (Plastics) -> Output: "Commission Regulation (EU) No 10/2011"
- IF text contains "13432" (Compostable) -> Output: "EN 13432 (Compostable OK)"
- IF text contains "14287" (Foil) -> Output: "EN 14287 (Foil)"
- IF text contains "FSC" -> Output: "FSC (Forest Stewardship Council)"
- ELSE -> Output strictly what is written (e.g., "ISO 9001").

**FIELD: certification (The "Standard")**
- Extract the Certification Body or Type.
- Valid Examples: "BRCGS", "DIN CERTCO", "TUV Austria", "ISO 9001", "ISO 45001", "FSSC 22000".

**FIELD: product_category**
- Brief description (e.g., "Paper Cup", "PE Coated Board").

**FIELD: country**
- CRITICAL: Look at the *Address* of the manufacturing site. Extract ONLY the Country (e.g., "China").

**FIELD: date_issued**
- Format: YYYY-MM-DD.

**FIELD: date_expired**
- Format: YYYY-MM-DD.
- Logic: If "Valid until 31 Jan 2027" -> "2027-01-31".

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
              text: "Extract all certificate information from this image.",
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
