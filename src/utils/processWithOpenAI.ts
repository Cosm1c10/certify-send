const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;

interface CertificateExtractionResult {
  supplier_name: string;
  certificate_number: string;
  country: string;
  scope: string;           // Product description
  measure: string;         // Regulation reference
  certification: string;
  product_category: string;
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
