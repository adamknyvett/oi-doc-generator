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

    const apiKey = overrides.apiKey || process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(401).json({ error: "No API key — paste your Anthropic key in the field at the top of the page." });

    const client = new Anthropic({ apiKey });
    const today = new Date().toLocaleDateString("en-AU", { day: "2-digit", month: "long", year: "numeric" });

    const system = `Extract data from a supplier quote. Return ONLY valid JSON, no markdown, no explanation.

Extract everything you can find: supplier name/address/ABN/ACN, contact name/email, bank BSB/account, buyer details, all line items with amounts, subtotal, GST, total, PO number, dates, reference numbers.

Amounts must be plain numbers (2640 not $2,640.00). If you cannot find an amount, use 0.

Return this exact structure:
{
  "po": {
    "number": "", "date": "", "delivery": "", "reference": "", "description": "",
    "supplier_name": "", "supplier_addr": "", "supplier_abn": "", "supplier_acn": "",
    "supplier_contact": "", "supplier_email": "",
    "bank_name": "", "bank_bsb": "", "bank_acct": "", "bank_branch": "",
    "buyer_name": "", "buyer_addr": "", "buyer_contact": "",
    "items": [{ "desc": "", "effort": "", "amount": 0 }],
    "subtotal": 0, "gst": 0, "total": 0, "notes": "", "ref": ""
  }
}`;

    let messages;
    if (fileBase64 && fileType === "application/pdf") {
      messages = [{
        role: "user",
        content: [
          { type: "document", source: { type: "base64", media_type: "application/pdf", data: fileBase64 } },
          { type: "text", text: `Extract all data from this quote PDF. Look carefully for dollar amounts in any pricing or quote table. Today: ${today}. Return ONLY JSON.` }
        ]
      }];
    } else {
      messages = [{
        role: "user",
        content: `Extract all data from this quote. Today: ${today}.\n\nDOCUMENT:\n${(text||"").slice(0,15000)}`
      }];
    }

    const msg = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 2000,
      system,
      messages
    });

    let raw = msg.content[0].text.trim()
      .replace(/```json\n?/g,"").replace(/```\n?/g,"").trim();
    const fb = raw.indexOf("{"); if (fb > 0) raw = raw.slice(fb);
    const lb = raw.lastIndexOf("}"); if (lb >= 0) raw = raw.slice(0, lb+1);
    const parsed = JSON.parse(raw);

    const allZero = (parsed.po?.items||[]).every(i => !i.amount || i.amount === 0);
    res.status(200).json({ ok: true, data: parsed, needsAmounts: allZero });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
