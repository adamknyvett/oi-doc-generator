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

    const systemPrompt = `You extract data from supplier quotes to create Purchase Orders and Tax Invoices. Return ONLY valid JSON, no markdown, no explanation.

CRITICAL: Find ALL line items with dollar amounts — especially in pricing tables near the end. Never return 0 for amounts if prices exist in the document.

Return this exact structure:
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
}
Rules: amounts are plain numbers no $ sign. PO = buyer issues to supplier. Invoice = supplier bills buyer. due = invoice date + 30 days if not specified. Return ONLY JSON.`;

    // Build message content — use native PDF if uploaded, otherwise text
    let userContent;
    if (fileBase64 && fileType === "application/pdf") {
      userContent = [
        {
          type: "document",
          source: { type: "base64", media_type: "application/pdf", data: fileBase64 }
        },
        {
          type: "text",
          text: `Extract ALL data including every line item and dollar amount from this PDF quote. Apply overrides: ${JSON.stringify(overrides)}. Today: ${today}. Return ONLY JSON.`
        }
      ];
    } else {
      // Text fallback
      userContent = `Extract all data from this quote. Apply overrides: ${JSON.stringify(overrides)}. Today: ${today}.\n\nDOCUMENT:\n${(text || "").slice(0, 15000)}`;
    }

    const msg = await client.messages.create({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 3000,
      system: systemPrompt,
      messages: [{ role: "user", content: userContent }]
    });

    let raw = msg.content[0].text.trim();
    raw = raw.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();
    const parsed = JSON.parse(raw);
    res.status(200).json({ ok: true, data: parsed });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
