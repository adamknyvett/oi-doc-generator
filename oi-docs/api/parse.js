const Anthropic = require("@anthropic-ai/sdk");

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });

  try {
    const { text, fileBase64, fileType, overrides = {} } = req.body;
    if (!text && !fileBase64) return res.status(400).json({ error: "No content provided" });

    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    const today = new Date().toLocaleDateString("en-AU", { day: "2-digit", month: "long", year: "numeric" });

    const systemPrompt = `You extract data from supplier quotes to produce Purchase Orders and Tax Invoices. Return ONLY valid JSON — no markdown fences, no explanation, nothing else.

MOST IMPORTANT RULES:
1. Find the PRICING TABLE / QUOTE section — it usually appears near the end under headings like "Quote", "Commercial Offer", "Pricing", or "Cost Breakdown"
2. Extract EVERY line item with its exact dollar amount — never return 0 for amounts if prices exist anywhere in the document
3. amounts must be plain numbers, no $ sign (e.g. 2640 not $2,640.00)
4. Extract bank/payment details: BSB, account number, account name, bank name
5. Extract ABN and ACN
6. PO = buyer issues to supplier | Invoice = supplier bills buyer
7. If no invoice number given, create one like INV-2026-001
8. due date = invoice date + 30 days if not stated

Return this exact JSON structure:
{
  "po": {
    "number": "", "date": "", "delivery": "", "reference": "", "description": "",
    "buyer_name": "", "buyer_addr": "", "buyer_contact": "", "buyer_abn": "",
    "supplier_name": "", "supplier_addr": "", "supplier_abn": "", "supplier_acn": "",
    "bank_name": "", "bank_bsb": "", "bank_acct": "", "bank_branch": "",
    "items": [{ "desc": "", "effort": "", "amount": 0 }],
    "subtotal": 0, "gst": 0, "total": 0, "notes": ""
  },
  "inv": {
    "number": "", "date": "", "due": "", "po_ref": "", "reference": "",
    "supplier_name": "", "supplier_addr": "", "supplier_abn": "", "supplier_acn": "",
    "bank_name": "", "bank_bsb": "", "bank_acct": "", "bank_branch": "",
    "buyer_name": "", "buyer_addr": "", "buyer_contact": "",
    "items": [{ "desc": "", "effort": "", "amount": 0 }],
    "subtotal": 0, "gst": 0, "total": 0, "notes": ""
  }
}`;

    let messages;

    if (fileBase64 && fileType === "application/pdf") {
      // Send PDF natively — claude-sonnet-4-6 can read PDFs
      messages = [{
        role: "user",
        content: [
          {
            type: "document",
            source: { type: "base64", media_type: "application/pdf", data: fileBase64 }
          },
          {
            type: "text",
            text: `Read this entire PDF carefully, especially the Quote/Pricing section near the end. Extract ALL line items with their exact dollar amounts.

Apply these overrides if provided: ${JSON.stringify(overrides)}
Today: ${today}

Return ONLY the JSON object.`
          }
        ]
      }];
    } else {
      // Text input
      messages = [{
        role: "user",
        content: `Extract all data from this quote document. Look carefully for the pricing table with line items and dollar amounts.

Apply overrides: ${JSON.stringify(overrides)}
Today: ${today}

DOCUMENT TEXT:
${(text || "").slice(0, 15000)}`
      }];
    }

    const msg = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 3000,
      system: systemPrompt,
      messages
    });

    let raw = msg.content[0].text.trim();
    raw = raw.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
    // Strip any preamble before the first { in case model adds commentary
    const firstBrace = raw.indexOf('{');
    if (firstBrace > 0) raw = raw.slice(firstBrace);
    // Strip anything after the last }
    const lastBrace = raw.lastIndexOf('}');
    if (lastBrace >= 0) raw = raw.slice(0, lastBrace + 1);
    const parsed = JSON.parse(raw);

    // Sanity check — if all amounts are 0 and we have text, something went wrong
    const allZero = (parsed.po?.items || []).every(i => !i.amount || i.amount === 0);
    if (allZero && parsed.po?.items?.length > 0) {
      return res.status(200).json({
        ok: true,
        data: parsed,
        warning: "Amounts could not be extracted automatically. Please check the document contains a pricing table, or use the paste text tab and include the pricing section."
      });
    }

    res.status(200).json({ ok: true, data: parsed });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
