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

export async function processWithOpenAI(base64Image: string): Promise<CertificateExtractionResult> {
  if (!OPENAI_API_KEY) {
    throw new Error('OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local');
  }

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
              text: 'Extract all certificate information from this image.',
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
