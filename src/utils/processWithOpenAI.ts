const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;

/**
 * Robust JSON extraction from AI response
 * Handles cases where AI outputs conversational text before/after JSON
 */
function extractJSON(text: string): any {
  // First, clean markdown code blocks
  const cleaned = text
    .replace(/```json\n?/g, '')
    .replace(/```\n?/g, '')
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
        console.error('Failed to parse extracted JSON:', e2);
        throw new Error('Invalid JSON format in AI response');
      }
    }
    console.error('No JSON found in response:', cleaned.substring(0, 200));
    throw new Error('No JSON found in AI response');
  }
}

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

export async function processWithOpenAI(base64Image: string, filename?: string, textContent?: string): Promise<CertificateExtractionResult> {
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local');
  }

  const isTextOnly = !!textContent && !base64Image;
  console.log('Processing certificate with filename:', filename || 'not provided', isTextOnly ? '(text mode)' : '(image mode)');

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
  let userContent: Array<{ type: string; text?: string; image_url?: { url: string } }>;

  if (isTextOnly && textContent) {
    // Text-only mode for DOCX files
    userContent = [
      {
        type: 'text',
        text: `Analyze this certificate document. Filename: "${filename || 'unknown.docx'}". Extract all certificate information from the following text:\n\n${textContent}`,
      },
    ];
  } else {
    // Image mode for PDF/Image files
    const imageContent = base64Image.startsWith('data:')
      ? base64Image
      : `data:image/jpeg;base64,${base64Image}`;

    userContent = [
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
    ];
  }

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
          content: userContent,
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
  const rawContent = data.choices[0]?.message?.content;

  if (!rawContent) {
    throw new Error('No response from OpenAI');
  }

  // Debug log for troubleshooting
  console.log('Raw AI Response (first 500 chars):', rawContent.substring(0, 500));

  // Parse the JSON response using robust extraction
  return extractJSON(rawContent);
}
