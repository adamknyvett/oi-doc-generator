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
    if (!apiKey) return res.status(401).json({ error: "No API key — paste your Anthropic key at the top of the page." });

    const client = new Anthropic({ apiKey });
    const today = new Date().toLocaleDateString("en-AU", { day: "2-digit", month: "long", year: "numeric" });

    const system = `You are a document data extraction specialist. Extract ALL financial and business data from this document.

CRITICAL — AMOUNTS: You MUST find every dollar amount in the document. Look in:
- Tables with price/amount/cost columns
- Line items with dollar figures  
- Totals, subtotals, GST rows
- Any text containing $ symbols followed by numbers
- Image-based tables — read the numbers visually

Return ONLY valid JSON, no markdown, no explanation:
{
  "po": {
    "number": "", "date": "", "delivery": "", "reference": "", "description": "",
    "supplier_name": "", "supplier_addr": "", "supplier_abn": "", "supplier_acn": "",
    "supplier_contact": "", "supplier_email": "", "supplier_phone": "",
    "bank_name": "", "bank_bsb": "", "bank_acct": "", "bank_branch": "",
    "buyer_name": "", "buyer_addr": "", "buyer_contact": "",
    "items": [{ "desc": "", "effort": "", "amount": 0, "bullet_points": [] }],
    "subtotal": 0, "gst": 0, "total": 0, "gst_rate": 0.10, "notes": "", "ref": ""
  },
  "inv": {
    "number": "", "date": "", "po_ref": "", "reference": "",
    "supplier_name": "", "supplier_addr": "", "supplier_abn": "", "supplier_acn": "",
    "supplier_contact": "", "supplier_email": "", "supplier_phone": "", "signed_by": "",
    "bank_name": "", "bank_bsb": "", "bank_acct": "", "bank_branch": "",
    "buyer_name": "", "buyer_addr": "", "buyer_contact": "",
    "items": [{ "desc": "", "effort": "", "amount": 0, "bullet_points": [] }],
    "subtotal": 0, "gst": 0, "total": 0, "gst_rate": 0.10
  }
}

Rules:
- amounts are plain numbers no $ sign (2640 not $2,640.00)
- For delivery/required date: only extract if it is explicitly a delivery or required-by date. Do NOT extract quote expiry dates or validity dates. Leave delivery blank if unclear.
- Detect GST rate: NZ documents use 0.15 (15%), AU documents use 0.10 (10%). Set gst_rate accordingly.
- If document is an invoice, populate inv. If it is a quote/PO, populate po. If both, populate both.
- For each line item, if there are sub-points, bullet points, or dash-separated details beneath the main description, extract them into bullet_points array
- Keep bullet_points as short clear strings without leading bullets or dashes
- Never return 0 for amounts if you can see numbers anywhere in the document
- Return ONLY the JSON object`;

    // Build content — send PDF natively AND as explicit vision request
    let content;

    if (fileBase64 && (fileType === "application/pdf" || fileType?.includes("pdf"))) {
      // Send as native PDF document — Claude can read both text and image layers
      content = [
        {
          type: "document",
          source: { type: "base64", media_type: "application/pdf", data: fileBase64 }
        },
        {
          type: "text",
          text: `Read this entire document carefully as if you were a human reading it visually. 
Find ALL dollar amounts — they may be in tables, images, or formatted as text.
Extract every line item with its exact dollar amount.
Today: ${today}.
Return ONLY the JSON object.`
        }
      ];
    } else if (fileBase64 && fileType?.includes("image")) {
      // Direct image upload (PNG, JPG etc)
      content = [
        {
          type: "image",
          source: { type: "base64", media_type: fileType, data: fileBase64 }
        },
        {
          type: "text",
          text: `Read all text and numbers from this invoice/quote image. Extract every line item and dollar amount. Today: ${today}. Return ONLY JSON.`
        }
      ];
    } else {
      // Plain text
      content = `Extract all data from this document. Today: ${today}.\n\nDOCUMENT:\n${(text||"").slice(0,15000)}`;
    }

    // Use opus for best vision/reading capability
    const msg = await client.messages.create({
      model: "claude-opus-4-5-20251101",
      max_tokens: 2500,
      system,
      messages: [{ role: "user", content }]
    });

    let raw = msg.content[0].text.trim()
      .replace(/```json\n?/g,"").replace(/```\n?/g,"").trim();
    const fb = raw.indexOf("{"); if (fb > 0) raw = raw.slice(fb);
    const lb = raw.lastIndexOf("}"); if (lb >= 0) raw = raw.slice(0, lb+1);
    let parsed = JSON.parse(raw);

    // If amounts still zero after opus — try again with explicit instruction
    const poZero  = (parsed.po?.items||[]).every(i => !i.amount || i.amount === 0);
    const invZero = (parsed.inv?.items||[]).every(i => !i.amount || i.amount === 0);

    if ((poZero || invZero) && fileBase64) {
      // Second pass — ask specifically about amounts
      const retry = await client.messages.create({
        model: "claude-opus-4-5-20251101",
        max_tokens: 1000,
        messages: [{
          role: "user",
          content: fileType?.includes("image")
            ? [
                { type: "image", source: { type: "base64", media_type: fileType, data: fileBase64 } },
                { type: "text", text: `Look carefully at this document image. List every dollar amount you can see, including line items. Format as JSON array: [{"desc":"item name","amount":1234.56}]. Also give subtotal, gst, total as numbers.` }
              ]
            : [
                { type: "document", source: { type: "base64", media_type: "application/pdf", data: fileBase64 } },
                { type: "text", text: `This PDF may have tables rendered as images. Look at every page carefully and find ALL dollar amounts. List them as JSON: [{"desc":"item name","amount":1234.56}] plus subtotal, gst, total.` }
              ]
        }]
      });

      let retryRaw = retry.content[0].text.trim()
        .replace(/```json\n?/g,"").replace(/```\n?/g,"").trim();
      const rb = retryRaw.indexOf("["); 
      if (rb >= 0) {
        try {
          const re = retryRaw.lastIndexOf("]");
          const items = JSON.parse(retryRaw.slice(rb, re+1));
          if (items.length > 0 && items[0].amount > 0) {
            // Patch amounts into parsed result
            const target = invZero ? parsed.inv : parsed.po;
            if (target) {
              target.items = items.map(i => ({desc: i.desc||'', effort: '', amount: i.amount||0}));
              const sub = items.reduce((s,i) => s + (i.amount||0), 0);
              target.subtotal = Math.round(sub * 100) / 100;
              target.gst      = Math.round(sub * 0.1 * 100) / 100;
              target.total    = Math.round((sub + sub*0.1) * 100) / 100;
            }
          }
        } catch(e) { /* retry parse failed, use what we have */ }
      }
    }

    const stillZero = (parsed.po?.items||[]).every(i => !i.amount || i.amount === 0) &&
                      (parsed.inv?.items||[]).every(i => !i.amount || i.amount === 0);

    res.status(200).json({ ok: true, data: parsed, needsAmounts: stillZero });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
