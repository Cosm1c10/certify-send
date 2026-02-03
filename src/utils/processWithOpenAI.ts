const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;

interface CertificateExtractionResult {
  supplier_name: string;
  certificate_number: string;
  country?: string;       // Legacy field
  region?: string;        // New field (v2)
  product_category?: string;
  ec_regulation: string;
  certification: string;
  date_issued: string;
  date_expired?: string;  // Legacy field
  date_expiry?: string;   // New field (v2)
  status?: string;        // New field (v2)
}

export async function processWithOpenAI(base64Image: string, filename?: string): Promise<CertificateExtractionResult> {
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local');
  }

  console.log('Processing certificate with filename:', filename || 'not provided');

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

  // Ensure proper data URL format
  const imageContent = base64Image.startsWith('data:')
    ? base64Image
    : `data:image/jpeg;base64,${base64Image}`;

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
    },
    body: JSON.stringify({
      model: 'gpt-4o',
      messages: [
        {
          role: 'system',
          content: systemPrompt,
        },
        {
          role: 'user',
          content: [
            {
              type: 'image_url',
              image_url: {
                url: imageContent,
              },
            },
            {
              type: 'text',
              text: `Analyze this certificate. Filename: "${filename || 'unknown.pdf'}". Extract all certificate information.`,
            },
          ],
        },
      ],
      max_tokens: 1000,
    }),
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'OpenAI API request failed');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;

  if (!content) {
    throw new Error('No response from OpenAI');
  }

  // Parse the JSON response, removing potential markdown code blocks
  const cleanedContent = content
    .replace(/```json\n?/g, '')
    .replace(/```\n?/g, '')
    .trim();

  return JSON.parse(cleanedContent);
}
