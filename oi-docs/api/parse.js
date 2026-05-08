const Anthropic = require("@anthropic-ai/sdk");

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "POST") return res.status(405).json({ error: "POST only" });

  try {
    const { text, overrides = {} } = req.body;
    if (!text || !text.trim()) return res.status(400).json({ error: "No text provided" });

    const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
    const today = new Date().toLocaleDateString("en-AU", { day: "2-digit", month: "long", year: "numeric" });

    const msg = await client.messages.create({
model: "claude-haiku-4-5-20251001",
      max_tokens: 3000,
      system: `Extract data from a supplier quote or invoice. Return ONLY a valid JSON object — no markdown fences, no explanation. Use this exact structure:
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
Rules: amounts are plain numbers. PO = buyer issues to supplier. Invoice = supplier bills buyer. If no due date, set 30 days after invoice date. Return ONLY JSON.`,
      messages: [{
        role: "user",
        content: `Extract from this document. Apply overrides if provided: ${JSON.stringify(overrides)}\nToday: ${today}\n\nDOCUMENT:\n${text.slice(0, 8000)}`
      }]
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
