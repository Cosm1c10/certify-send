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
You are a high-precision Compliance Data Extraction Engine.
Your goal is to extract structured data from certification documents (PDF/Images) for a master database.

### CRITICAL RULES FOR EXTRACTION:

1. **SUPPLIER NAME NORMALIZATION (Must be Exact)**
   - **Safira Rule:** If the document mentions "Safira Amb.", "SAFİRA AMBALAJ", "Safira Ambalaj San. Ve Tic." -> Output ONLY: "Safira Ambalaj".
   - **Huhtamaki Rule:** If the document mentions "Huhtamaki", "Huhtamaki Turkey" -> Output ONLY: "Huhtamaki".
   - **General Rule:** Remove legal suffixes like "San. Ve Tic. A.Ş.", "Co., Ltd", "Ltd. Şti.", "Pvt. Ltd". Output the clean company name.

2. **COUNTRY DETECTION**
   - Scan the address block in the header/footer.
   - If "Istanbul", "Turkey", "Türkiye" found -> Output: "Turkey".
   - If "China", "Changsha", "Hunan" found -> Output: "China".
   - If "Dublin", "Ireland" found -> Output: "Ireland".

3. **EC REGULATION / MEASURE (Strict Search)**
   - Search the *entire* text for these specific regulation numbers.
   - If "10/2011" is found (even inside "EU No 10/2011") -> Output: "Commission Regulation (EU) No 10/2011".
   - If "2023/2006" is found -> Output: "Commission Regulation (EC) No 2023/2006".
   - If "1935/2004" is found -> Output: "Regulation (EC) No 1935/2004".
   - If "94/62/EC" is found -> Output: "Directive 94/62/EC".
   - **Fallback:** Only use "Migration Test" if absolutely NO regulation numbers are present.

4. **CERTIFICATE / REPORT NUMBER**
   - Look for labels: "Report No", "Rapor No", "Certificate No", "Registration No".
   - Capture IDs like: "FS10068846", "3193", "7P1350".

5. **DATES (Format: YYYY-MM-DD)**
   - Extract "Issue Date" (or "Tarih").
   - Extract "Expiry Date" (or "Valid until").
   - Note: "Tarih: 11.03.2019" is March 11, 2019.

### OUTPUT JSON FORMAT:
{
  "supplier_name": "string (Normalized company name)",
  "certificate_number": "string (Report No / Cert ID)",
  "country": "string (Country of origin)",
  "ec_regulation": "string (Full regulation name or 'Migration Test')",
  "certification": "string (e.g., 'Migration Test', 'BRCGS', 'ISO 22000')",
  "date_issued": "YYYY-MM-DD",
  "date_expired": "YYYY-MM-DD",
  "status": "string (Valid/Expired)"
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
