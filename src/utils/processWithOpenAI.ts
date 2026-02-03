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
