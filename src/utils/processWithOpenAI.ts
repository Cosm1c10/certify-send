const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;

// MAX INPUT SIZE (prevents timeouts)
const MAX_TEXT_LENGTH = 50000; // Increased buffer

// INDESTRUCTIBLE JSON PARSER (Self-Healing)
function extractJSON(text: string): any {
  console.log('Raw AI Response:', text);
  let cleanText = text.replace(/```json/g, '').replace(/```/g, '').trim();

  // Attempt 1: Standard JSON Parse
  const firstOpen = cleanText.indexOf('{');
  const lastClose = cleanText.lastIndexOf('}');

  if (firstOpen !== -1 && lastClose !== -1) {
    const jsonString = cleanText.substring(firstOpen, lastClose + 1);
    try {
      return JSON.parse(jsonString);
    } catch (e) {
      console.warn('Standard JSON Parse Failed. Trying Regex Repair...', e);
    }
  }

  // Attempt 2: Regex Extraction (Fallback)
  // This extracts fields even if the JSON syntax is broken
  const extractField = (key: string) => {
    const regex = new RegExp(`"${key}"\\s*:\\s*"([^"]*)"`, 'i');
    const match = cleanText.match(regex);
    return match ? match[1] : null;
  };

  const supplier = extractField('supplier_name');

  // If we found a supplier, we assume we have partial data and return it
  if (supplier) {
    return {
      supplier_name: supplier,
      country: extractField('country') || '',
      scope: extractField('scope') || '+',
      measure: extractField('measure') || 'EU Regulation 2016/425', // Safe fallback
      certification: extractField('certification') || 'Standard',
      product_category: extractField('product_category') || 'Gloves',
      date_issued: extractField('date_issued'),
      date_expired: extractField('date_expired'),
      status: 'Extracted (Repair)'
    };
  }

  // Attempt 3: Total Failure (Return Error Object)
  return {
    supplier_name: 'Error: Could not read file',
    country: 'Unknown',
    scope: '!',
    measure: 'Manual Review',
    certification: 'Unknown',
    product_category: 'Unknown',
    date_issued: null,
    date_expired: null,
    status: 'Error'
  };
}

interface CertificateExtractionResult {
  supplier_name: string;
  country: string;
  scope: string;
  measure: string;
  certification: string;
  product_category: string;
  date_issued: string | null;
  date_expired: string | null;
  status?: string;
}

// THE MASTER PROMPT (Precision Tuned)
const systemPrompt = `
You are a Compliance Data Extraction Engine.
Goal: Extract structured data.
CRITICAL: OUTPUT ONLY RAW JSON.

### 1. EXTRACTION LOGIC
- **Test Reports (Gloves):** Treat "EN 455", "EN 374", "EN 420" as VALID CERTIFICATES.
  - **Supplier:** Find "Applicant" or "Manufacturer".
  - **Expiry:** If missing, calculate **Issue Date + 3 Years**.

### 2. CLASSIFICATION & MAPPING (STRICT)
#### **A. GENERAL (!)**
- **Scope:** "!"
- **Rules:** BRC, ISO 9001, ISO 22000, ISO 14001, ISO 45001, FSC.

#### **B. GLOVES & PPE (+)** (HIGHEST PRIORITY)
- **Triggers:** "EN 455", "EN 374", "EN 420", "Module B", "Cat III", "Nitrile Gloves".
- **CRITICAL:** If ANY of these triggers are found:
  - **Measure** MUST BE "EU Regulation 2016/425".
  - **Scope** MUST BE "+".
  - **NEVER** leave Measure blank.

#### **C. SPECIFIC (+)**
- **Scope:** "+"
- **Rules:** DoC, Migration Reports, EN 13432.

### 3. DATE RULES
- **Format:** YYYY-MM-DD.
- **Parsing:** "09.10.2024" = **October 9th** (European). NEVER September.

### 4. OUTPUT JSON SCHEME
{
  "supplier_name": "string",
  "country": "string",
  "scope": "string",
  "measure": "string",
  "certification": "string",
  "product_category": "string",
  "date_issued": "YYYY-MM-DD",
  "date_expired": "YYYY-MM-DD"
}
`;

export async function processWithOpenAI(
  base64Image: string,
  filename?: string,
  textContent?: string
): Promise<CertificateExtractionResult> {
  // GLOBAL TRY/CATCH - Never let the function crash
  try {
    if (!OPENAI_API_KEY) {
      throw new Error('OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local');
    }

    const isTextOnly = !!textContent && !base64Image;
    console.log('Processing:', filename || 'unknown', isTextOnly ? '(text)' : '(image)');

    // Build user message content based on input type
    let userContent: Array<{ type: string; text?: string; image_url?: { url: string } }>;

    if (isTextOnly && textContent) {
      // CRASH PROTECTION: Truncate massive text files
      const truncatedText = textContent.length > MAX_TEXT_LENGTH
        ? textContent.slice(0, MAX_TEXT_LENGTH) + '\n\n[TRUNCATED]'
        : textContent;

      console.log(`Text length: ${textContent.length}, truncated: ${truncatedText.length}`);

      userContent = [
        {
          type: 'text',
          text: `Extract certificate data from ("${filename || 'unknown'}"):\n\n${truncatedText}`,
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
          text: `Extract certificate data from ("${filename || 'unknown'}").`,
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
        max_tokens: 1500, // Increased for complex files
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      console.error('OpenAI API Error:', error);
      throw new Error(error.error?.message || 'OpenAI API request failed');
    }

    const data = await response.json();
    const rawContent = data.choices[0]?.message?.content;

    if (!rawContent) {
      console.error('No response from OpenAI');
      return {
        supplier_name: 'Error: No AI Response',
        country: 'Unknown',
        scope: '!',
        measure: 'Manual Review',
        certification: 'AI returned empty',
        product_category: 'Check File',
        date_issued: null,
        date_expired: null,
        status: 'Error'
      };
    }

    // Parse with self-healing extraction
    return extractJSON(rawContent);

  } catch (error) {
    // INVINCIBLE FALLBACK
    console.error('Processing Error:', error);
    return {
      supplier_name: 'Error: System Crash',
      country: 'Unknown',
      scope: '!',
      measure: 'Check File',
      certification: 'System Error',
      product_category: 'Unknown',
      date_issued: null,
      date_expired: null,
      status: 'Error'
    };
  }
}
