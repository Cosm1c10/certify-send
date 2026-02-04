const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;

interface CertificateExtractionResult {
  supplier_name: string;
  certificate_number: string;
  country: string;
  scope: string;           // Symbol: "!" (Factory/System) or "+" (Product-specific)
  measure: string;         // Mapped regulation reference
  certification: string;   // Document type or certifying body
  product_category: string; // Product description (detailed)
  date_issued: string;
  date_expired: string | null;
}

export async function processWithOpenAI(base64Image: string, filename?: string): Promise<CertificateExtractionResult> {
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local');
  }

  console.log('Processing certificate with filename:', filename || 'not provided');

  const systemPrompt = `
You are a Compliance Data Extraction Engine.
Goal: Extract structured data for the Client Master File with strict "General" vs "Specific" classification.

### 1. CLASSIFICATION RULES (CRITICAL)

**Scope Column (The Symbol):**
- Output "!" (Exclamation Mark) IF the certificate is for the **Factory/Management System**.
  - Examples: BRC, BRCGS, ISO 9001, ISO 22000, ISO 45001, ISO 14001, GMP, FSC, GRS, FSSC 22000.
- Output "+" (Plus Sign) IF the certificate is for a **Specific Product Performance**.
  - Examples: Compostable (EN 13432), Recyclable (ISO 14021), Migration Test, Food Grade, 10/2011, Declaration of Compliance.

**Product Category Column (The Description):**
- EXTRACT the detailed product description here.
- Examples: "Aqueous Coated Paper Cup", "Blanks of bio coated paper", "Disposable kraft paper cups", "PET Bottles".
- Do NOT put generic terms like "Paper" or "Plastic". Put the full product string from the certificate.

**Measure Column (The Standard Mapping):**
- IF ISO 22000 OR BRC OR BRCGS OR ISO 9001 OR FSSC 22000 -> Output: "(EC) No 2023/2006"
- IF Compostable (DIN CERTCO / TUV / EN 13432) -> Output: "EN 13432 (Compostable)"
- IF Recyclable (ISO 14021) -> Output: "ISO 14021 (Recyclable)"
- IF FSC -> Output: "FSC"
- IF 10/2011 -> Output: "Commission Regulation (EU) No 10/2011"
- IF 1935/2004 -> Output: "Regulation (EC) No 1935/2004"
- IF Migration Test (no specific reg) -> Output: "Migration Test"

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
