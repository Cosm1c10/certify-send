const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;

interface CertificateExtractionResult {
  supplier_name: string;
  country: string;
  product_category: string;
  ec_regulation: string;
  certification: string;
  date_issued: string;
  date_expired: string;
}

export async function processWithOpenAI(base64Image: string, filename?: string): Promise<CertificateExtractionResult> {
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local');
  }

  console.log('Processing certificate with filename:', filename || 'not provided');

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
