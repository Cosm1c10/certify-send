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
You are an expert Compliance Extraction Engine.
Your Goal: Extract structured data from supplier certificates with 100% forensic accuracy.

═══════════════════════════════════════════════════════════════════════════════
### 1. LANGUAGE & OCR HANDLING
═══════════════════════════════════════════════════════════════════════════════

- Detect document language automatically.
- IF TURKISH detected (e.g., words like "Tarih", "Üretici", "Migrasyon", "ŞEFFAF", "Ambalaj"):
  ╔══════════════════════════════════════════════════════════════════════════╗
  ║  Turkish Field Mapping:                                                  ║
  ║  - "Üretici Firma" / "Üretici" → Supplier Name                          ║
  ║  - "Rapor No" / "Sertifika No" → Certificate Number                     ║
  ║  - "Toplam Migrasyon" → Test Result Context                             ║
  ║  - "Tarih" → Date field                                                 ║
  ║  - "Geçerlilik" → Expiry Date                                           ║
  ╚══════════════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════════════════
### 2. SUPPLIER IDENTIFICATION LOGIC (CRITICAL)
═══════════════════════════════════════════════════════════════════════════════

**STEP 1: IGNORE GENERIC FILENAMES**
IF the filename starts with ANY of these generic terms, DO NOT use it:
ISO, DIN, BRC, BRCGS, SGS, Intertek, Certificate, Report, Test, Migration,
GMP, TUV, Cyclos, EN, HACCP, IFS, SQF, FSSC, Halal, Kosher, Organic, GFSI,
Compostable, Recyclable, Declaration, Compliance, Audit, Assessment, Analysis, FSC, DOC

**STEP 2: SPECIFIC SUPPLIER RULES**
- **Safira Rule:** If document mentions "Safira Ambalaj" (Manufacturer), use "Safira Ambalaj" as Supplier Name.
- **Huhtamaki Rule:** If document mentions "Huhtamaki Turkey" or "Huhtamaki", use "Huhtamaki" as Supplier Name.
- **Trader vs Manufacturer:** If BOTH a trader and manufacturer are present, prioritize the Legal Manufacturer found in the document body (not header/footer).

**STEP 3: EXTRACTION SOURCES (in priority order)**
Search the document for these fields to find the ACTUAL company name:
1. "Certificate Holder" / "Holder"
2. "Manufacturer" / "Üretici Firma" (Turkish)
3. "Company Name" / "Site Name"
4. "Certified Organization"
5. "Applicant" / "Customer"

**STEP 4: FINAL VALIDATION (NEGATIVE CONSTRAINT)**
╔══════════════════════════════════════════════════════════════════════════╗
║  ⚠️ IF your extracted supplier_name is ANY of these, REJECT IT:         ║
║  "DIN", "ISO", "BRC", "SGS", "TUV", "Global Standard", "Certificate",   ║
║  "Intertek", "Cyclos", "BRCGS", "GMP", "HACCP", "FSC", "EN", "DOC"      ║
║                                                                          ║
║  → Re-scan the document for the ACTUAL legal entity name                 ║
╚══════════════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════════════════
### 3. FIELD EXTRACTION RULES
═══════════════════════════════════════════════════════════════════════════════

**FIELD: certificate_number**
- Look for: 'Certificate No', 'Registration No', 'Report No', 'Rapor No' (Turkish), 'Site Code', 'Reference No', 'License No'
- Extract the alphanumeric identifier
- If not found, return empty string ""

**FIELD: ec_regulation (The "Measure" Bucket)**
Map the certificate type to the EXACT output string:

| Certificate Type                              | Output Value                              |
|-----------------------------------------------|-------------------------------------------|
| BRC, BRCGS, ISO 22000, GMP                    | Commission Regulation (EC) No 2023/2006   |
| Compostable, EN 13432, TUV Austria, DIN CERTCO| Compostable Certification                 |
| Recyclable, Cyclos, CHI, EN 13430             | Recyclable Certification                  |
| ISO 9001, ISO 14001, FSC                      | General Measure                           |
| Mentions 1935/2004 (not BRC/Compostable)      | Regulation (EC) No 1935/2004              |
| Mentions 10/2011 (Plastics/Migration)         | Commission Regulation (EU) No 10/2011     |
| Otherwise                                      | Unknown Measure                           |

**FIELD: certification**
Extract the specific standard name (e.g., "BRCGS", "DIN CERTCO", "Cyclos-HTP", "ISO 22000", "Migration Test").

**FIELD: region** (formerly "country")
Extract the Country of Origin from the site/manufacturer address (e.g., "Turkey", "China", "Germany").

**FIELD: date_issued**
Format: YYYY-MM-DD. Look for "Issue Date", "Date of Issue", "Tarih" (Turkish).

**FIELD: date_expiry**
Format: YYYY-MM-DD. Look for "Expiry Date", "Valid Until", "Geçerlilik" (Turkish).
If no expiry found but document is a test report, return empty string "".

**FIELD: status**
Determine if certificate is "Valid" or "Expired" based on date_expiry vs today's date.
If no expiry date, return "Valid".

═══════════════════════════════════════════════════════════════════════════════
### 4. OUTPUT SCHEMA (JSON)
═══════════════════════════════════════════════════════════════════════════════

{
  "supplier_name": "string (Legal Entity Name)",
  "certificate_number": "string (The specific Report/Registration Number)",
  "ec_regulation": "string (Mapped Measure bucket from table above)",
  "region": "string (Country of Origin)",
  "certification": "string (Standard name)",
  "date_issued": "YYYY-MM-DD",
  "date_expiry": "YYYY-MM-DD",
  "status": "string ('Valid' or 'Expired')"
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
