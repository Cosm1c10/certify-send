const OPENAI_API_KEY = import.meta.env.VITE_OPENAI_API_KEY;

// MAX INPUT SIZE
const MAX_TEXT_LENGTH = 50000;

// 1. DATA SANITIZER (Prevents Excel Drops)
function sanitize(val: any, defaultVal: string = ""): string {
  if (val === null || val === undefined || val === "null") return defaultVal;
  return String(val).trim();
}

// 2. HARD-CODED BUSINESS LOGIC (Fixes Empty Measures)
function applyBusinessLogic(data: any): any {
  const cert = (data.certification || "").toLowerCase();
  const prod = (data.product_category || "").toLowerCase();
  const measure = (data.measure || "").toLowerCase();

  // RULE A: GLOVES (The "Force" Rule)
  if (cert.includes("en 455") || cert.includes("en 374") || cert.includes("en 420") || cert.includes("glove") || prod.includes("glove")) {
    data.scope = "+";
    data.measure = "EU Regulation 2016/425";
    if (!data.product_category || data.product_category === "Goods") data.product_category = "Gloves";
  }

  // RULE B: FACTORY CERTS (ISO 9001, BRC, FSSC)
  else if (cert.includes("iso 9001") || cert.includes("brc") || cert.includes("iso 22000") || cert.includes("fssc")) {
    data.scope = "!";
    if (!data.measure || measure === "national regulation" || measure === "") {
      data.measure = "(EC) No 2023/2006";
    }
  }

  // RULE C: ISO 14001
  else if (cert.includes("iso 14001")) {
    data.scope = "!";
    data.measure = "EU Waste Framework Directive (2008/98/EC)";
  }

  // RULE D: ISO 45001
  else if (cert.includes("iso 45001")) {
    data.scope = "!";
    data.measure = "EU Directive 89/391/EEC";
  }

  // RULE E: DoC (Declaration of Compliance) -> Specific
  else if (cert.includes("declaration of conformity") || cert.includes("doc")) {
    data.scope = "+";
    if (!data.measure) data.measure = "(EC) No 1935/2004";
  }

  return data;
}

// 3. ROBUST EXTRACTION (Self-Healing)
function extractJSON(text: string): any {
  console.log("Raw AI Response:", text);
  let cleanText = text.replace(/```json/g, "").replace(/```/g, "").trim();

  let data: any = {};

  // Attempt 1: Standard JSON Parse
  try {
    const first = cleanText.indexOf('{');
    const last = cleanText.lastIndexOf('}');
    if (first !== -1 && last !== -1) {
      data = JSON.parse(cleanText.substring(first, last + 1));
    }
  } catch (e) {
    console.warn("JSON Parse failed, trying regex...");
  }

  // Attempt 2: Regex Fallback
  const scavenge = (key: string) => {
    const match = cleanText.match(new RegExp(`"${key}"\\s*:\\s*"([^"]*)"`, "i"));
    return match ? match[1] : null;
  };

  if (!data.supplier_name) data.supplier_name = scavenge("supplier_name");
  if (!data.country) data.country = scavenge("country");
  if (!data.certification) data.certification = scavenge("certification");
  if (!data.measure) data.measure = scavenge("measure");
  if (!data.scope) data.scope = scavenge("scope");
  if (!data.product_category) data.product_category = scavenge("product_category");
  if (!data.date_issued) data.date_issued = scavenge("date_issued");
  if (!data.date_expired) data.date_expired = scavenge("date_expired");

  // APPLY THE HARD LOGIC
  data = applyBusinessLogic(data);

  // 4. FINAL SANITIZATION
  return {
    supplier_name: sanitize(data.supplier_name, "Unknown Supplier"),
    country: sanitize(data.country, "Unknown"),
    scope: sanitize(data.scope, "!"),
    measure: sanitize(data.measure, "National Regulation"),
    certification: sanitize(data.certification, "Standard"),
    product_category: sanitize(data.product_category, "Goods"),
    date_issued: sanitize(data.date_issued, ""),
    date_expired: sanitize(data.date_expired, ""),
    status: "Success"
  };
}

interface CertificateExtractionResult {
  supplier_name: string;
  country: string;
  scope: string;
  measure: string;
  certification: string;
  product_category: string;
  date_issued: string;
  date_expired: string;
  status?: string;
}

// 5. MASTER PROMPT
const systemPrompt = `
You are a Compliance Data Extraction Engine.
CRITICAL: OUTPUT RAW JSON ONLY.

### LOGIC RULES
- **Gloves:** Treat EN 455, EN 374, EN 420 as VALID. Supplier = Applicant.
- **Dates:** Format YYYY-MM-DD. If no expiry, calculate **Issue Date + 3 Years**.
- **European Dates:** "09.10.2024" = October 9th. NEVER September.

### OUTPUT SCHEME
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
  try {
    if (!OPENAI_API_KEY) {
      throw new Error("OpenAI API key not configured. Add VITE_OPENAI_API_KEY to .env.local");
    }

    const isTextOnly = !!textContent && !base64Image;
    console.log("Processing:", filename || "unknown", isTextOnly ? "(text)" : "(image)");

    let userContent: Array<{ type: string; text?: string; image_url?: { url: string } }>;

    if (isTextOnly && textContent) {
      // TRUNCATE to prevent timeouts
      const truncatedText = textContent.length > MAX_TEXT_LENGTH
        ? textContent.slice(0, MAX_TEXT_LENGTH) + "\n\n[TRUNCATED]"
        : textContent;

      console.log(`Text: ${textContent.length} chars, truncated: ${truncatedText.length}`);

      userContent = [
        { type: "text", text: `Extract data from (${filename || "unknown"}):\n\n${truncatedText}` },
      ];
    } else {
      const imageContent = base64Image.startsWith("data:")
        ? base64Image
        : `data:image/jpeg;base64,${base64Image}`;

      userContent = [
        { type: "image_url", image_url: { url: imageContent } },
        { type: "text", text: `Extract data from (${filename || "unknown"}).` },
      ];
    }

    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: "gpt-4o",
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userContent },
        ],
        max_tokens: 1500,
      }),
    });

    if (!response.ok) {
      const error = await response.json();
      console.error("OpenAI API Error:", error);
      throw new Error(error.error?.message || "OpenAI API request failed");
    }

    const data = await response.json();
    const rawContent = data.choices[0]?.message?.content;

    if (!rawContent) {
      console.error("No response from OpenAI");
      return {
        supplier_name: "Error: No AI Response",
        country: "Unknown",
        scope: "!",
        measure: "Manual Review",
        certification: "AI returned empty",
        product_category: "Unknown",
        date_issued: "",
        date_expired: "",
        status: "Error"
      };
    }

    // Extract + Sanitize + Apply Business Logic
    return extractJSON(rawContent);

  } catch (error) {
    console.error("Processing Error:", error);
    return {
      supplier_name: "Error: Processing Failed",
      country: "Unknown",
      scope: "!",
      measure: "Manual Review",
      certification: "File too complex",
      product_category: "Unknown",
      date_issued: "",
      date_expired: "",
      status: "Error"
    };
  }
}
