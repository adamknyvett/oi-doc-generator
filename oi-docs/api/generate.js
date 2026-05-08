const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
} = require("docx");

const NAVY="0A1628", TEAL="00B4D8", ACCENT="0077B6", LTGRAY="F4F7FA", MID="8CA0B5", WHITE="FFFFFF";

function money(v) {
  return "$" + parseFloat(v||0).toLocaleString("en-AU",{minimumFractionDigits:2,maximumFractionDigits:2});
}

const none = {style:BorderStyle.NONE,size:0,color:"FFFFFF"};
const btm  = {style:BorderStyle.SINGLE,size:3,color:"E8ECF0"};
const tealB= {style:BorderStyle.SINGLE,size:8,color:TEAL};

function tc(text, {bold=false,bg="FFFFFF",w=4680,align="left",color="111111",sz=18}={}) {
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:70,bottom:70,left:120,right:120},
    borders:{top:none,bottom:none,left:none,right:none},
    children:[new Paragraph({
      alignment:align==="right"?AlignmentType.RIGHT:AlignmentType.LEFT,
      spacing:{before:0,after:0},
      children:[new TextRun({text:String(text||""),bold,font:"Calibri",size:sz,color})]
    })]
  });
}

function tcB(text, {bold=false,bg="FFFFFF",w=4680,align="left",color="111111",sz=18}={}) {
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:80,bottom:80,left:120,right:120},
    borders:{top:none,bottom:btm,left:none,right:none},
    children:[new Paragraph({
      alignment:align==="right"?AlignmentType.RIGHT:AlignmentType.LEFT,
      spacing:{before:0,after:0},
      children:[new TextRun({text:String(text||""),bold,font:"Calibri",size:sz,color})]
    })]
  });
}

function gap(sz=120) {
  return new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:"",size:sz})]});
}

function sectionLabel(text) {
  return new Paragraph({
    spacing:{before:160,after:60},
    children:[new TextRun({text:text.toUpperCase(),bold:true,font:"Calibri",size:14,color:ACCENT})]
  });
}

function hr() {
  return new Paragraph({
    spacing:{before:80,after:80},
    border:{bottom:{style:BorderStyle.SINGLE,size:6,color:TEAL,space:1}},
    children:[]
  });
}

function buildDoc(d, docType) {
  const isInv   = docType==="INVOICE";
  const num     = d.number||"";
  const date    = d.date||"";
  const fromName= isInv?d.supplier_name:d.buyer_name;
  const fromAddr= (isInv?d.supplier_addr:d.buyer_addr)||"";
  const fromAbn = isInv?d.supplier_abn:d.buyer_abn;
  const fromAcn = isInv?(d.supplier_acn||""):"";
  const toName  = isInv?d.buyer_name:d.supplier_name;
  const toAddr  = (isInv?d.buyer_addr:d.supplier_addr)||"";
  const toContact=isInv?(d.buyer_contact||""):"";
  const fromLabel=isInv?"FROM / SUPPLIER":"BUYER / ISSUED BY";
  const toLabel  =isInv?"BILL TO":"VENDOR / SUPPLIER";
  const ref3lbl  =isInv?"DUE DATE":"DELIVERY";
  const ref3val  =isInv?(d.due||"30 days from invoice date"):(d.delivery||"As per quote");
  const ref4lbl  =isInv?"PO REFERENCE":"REFERENCE";
  const ref4val  =isInv?(d.po_ref||""):(d.reference||"");
  const totalLbl =isInv?"TOTAL DUE (incl. GST)":"TOTAL PAYABLE (incl. GST)";

  const ch = [];

  // Title
  ch.push(new Paragraph({
    spacing:{before:0,after:120},
    children:[new TextRun({text:isInv?"TAX INVOICE":"PURCHASE ORDER",bold:true,font:"Calibri",size:56,color:NAVY})]
  }));
  ch.push(new Paragraph({
    spacing:{before:0,after:80},
    children:[
      new TextRun({text:num+"   ",bold:true,font:"Calibri",size:22,color:ACCENT}),
      new TextRun({text:date,font:"Calibri",size:18,color:MID})
    ]
  }));
  ch.push(hr());
  ch.push(gap(100));

  // Parties
  const fA = fromAddr.replace(/\\n/g," | ");
  const tA = toAddr.replace(/\\n/g," | ");
  ch.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[4560,4800],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[
      new TableRow({children:[tc(fromLabel,{bold:true,w:4560,color:MID,sz:14}),tc(toLabel,{bold:true,w:4800,color:MID,sz:14})]}),
      new TableRow({children:[tc(fromName,{bold:true,w:4560,color:NAVY,sz:22}),tc(toName,{bold:true,w:4800,color:NAVY,sz:22})]}),
      new TableRow({children:[tc(fA,{w:4560,color:"444444",sz:16}),tc(tA+(toContact?" | Attn: "+toContact:""),{w:4800,color:"444444",sz:16})]}),
      new TableRow({children:[tc((fromAbn?"ABN: "+fromAbn:"")+(fromAcn?"  ·  ACN: "+fromAcn:""),{w:4560,color:MID,sz:14}),tc("",{w:4800})]})
    ]
  }));
  ch.push(gap(140));

  // Details bar
  ch.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[2340,2340,2340,2340],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[
      new TableRow({children:[
        tc(isInv?"INVOICE NUMBER":"PO NUMBER",{bold:true,bg:LTGRAY,w:2340,color:MID,sz:14}),
        tc("DATE",{bold:true,bg:LTGRAY,w:2340,color:MID,sz:14}),
        tc(ref3lbl,{bold:true,bg:LTGRAY,w:2340,color:MID,sz:14}),
        tc(ref4lbl,{bold:true,bg:LTGRAY,w:2340,color:MID,sz:14})
      ]}),
      new TableRow({children:[
        tc(num,{bold:true,bg:LTGRAY,w:2340,color:NAVY,sz:18}),
        tc(date,{bg:LTGRAY,w:2340,color:"333333",sz:18}),
        tc(ref3val,{bg:LTGRAY,w:2340,color:"333333",sz:18}),
        tc(ref4val,{bg:LTGRAY,w:2340,color:"333333",sz:18})
      ]})
    ]
  }));
  ch.push(gap(160));

  // Description
  if (d.description) {
    ch.push(sectionLabel("Project Description"));
    ch.push(new Paragraph({spacing:{before:0,after:160},children:[new TextRun({text:d.description,font:"Calibri",size:17,color:"333333"})]}));
  }

  // Line items
  ch.push(sectionLabel(isInv?"Services Rendered":"Line Items"));
  const headerRow = new TableRow({children:[
    new TableCell({width:{size:5040,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:none,bottom:none,left:none,right:none},
      children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:"DESCRIPTION",bold:true,font:"Calibri",size:16,color:WHITE})]})]}),
    new TableCell({width:{size:2520,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:none,bottom:none,left:none,right:none},
      children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:"EFFORT / DAYS",bold:true,font:"Calibri",size:16,color:WHITE})]})]}),
    new TableCell({width:{size:1800,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:none,bottom:none,left:none,right:none},
      children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[new TextRun({text:"AMOUNT (EXCL. GST)",bold:true,font:"Calibri",size:16,color:WHITE})]})]}),
  ]});

  const itemRows = (d.items||[]).map((item,i) => {
    const bg = i%2===0?"FFFFFF":LTGRAY;
    return new TableRow({children:[
      new TableCell({width:{size:5040,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:none,bottom:btm,left:none,right:none},
        children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:item.desc||"",font:"Calibri",size:17,color:"222222"})]})]}),
      new TableCell({width:{size:2520,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:none,bottom:btm,left:none,right:none},
        children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:item.effort||"",font:"Calibri",size:16,color:"555555"})]})]}),
      new TableCell({width:{size:1800,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:none,bottom:btm,left:none,right:none},
        children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[new TextRun({text:money(item.amount),font:"Calibri",size:17,color:"111111"})]})]})
    ]});
  });

  ch.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[5040,2520,1800],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[headerRow,...itemRows]
  }));
  ch.push(gap(100));

  // Totals
  ch.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[5040,2700,1620],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[
      new TableRow({children:[
        new TableCell({width:{size:5040,type:WidthType.DXA},borders:{top:none,bottom:btm,left:none,right:none},margins:{top:60,bottom:60,left:0,right:0},children:[new Paragraph({children:[]})]}),
        tcB("Sub-Total (excl. GST)",{w:2700,color:"555555",sz:17}),
        tcB(money(d.subtotal),{w:1620,align:"right",sz:17})
      ]}),
      new TableRow({children:[
        new TableCell({width:{size:5040,type:WidthType.DXA},borders:{top:none,bottom:btm,left:none,right:none},margins:{top:60,bottom:60,left:0,right:0},children:[new Paragraph({children:[]})]}),
        tcB("GST (10%)",{w:2700,color:"555555",sz:17}),
        tcB(money(d.gst),{w:1620,align:"right",sz:17})
      ]}),
      new TableRow({children:[
        new TableCell({width:{size:5040,type:WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},margins:{top:80,bottom:80,left:0,right:0},children:[new Paragraph({children:[]})]}),
        new TableCell({width:{size:2700,type:WidthType.DXA},borders:{top:tealB,bottom:none,left:none,right:none},margins:{top:80,bottom:80,left:120,right:120},
          children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:totalLbl,bold:true,font:"Calibri",size:20,color:ACCENT})]})]}),
        new TableCell({width:{size:1620,type:WidthType.DXA},borders:{top:tealB,bottom:none,left:none,right:none},margins:{top:80,bottom:80,left:120,right:120},
          children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[new TextRun({text:money(d.total),bold:true,font:"Calibri",size:22,color:ACCENT})]})]}),
      ]})
    ]
  }));
  ch.push(gap(180));

  // Bank
  if (d.bank_bsb||d.bank_acct||d.bank_name) {
    ch.push(sectionLabel("Payment Details"));
    const bankRows = [["Account Name",d.bank_name],["BSB",d.bank_bsb],["Account Number",d.bank_acct],["Bank",d.bank_branch]].filter(r=>r[1]);
    ch.push(new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[2200,7160],
      borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
      rows:bankRows.map(([k,v])=>new TableRow({children:[
        new TableCell({width:{size:2200,type:WidthType.DXA},shading:{fill:LTGRAY,type:ShadingType.CLEAR},margins:{top:70,bottom:70,left:120,right:120},borders:{top:none,bottom:btm,left:none,right:none},
          children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:k,bold:true,font:"Calibri",size:16,color:MID})]})]}),
        new TableCell({width:{size:7160,type:WidthType.DXA},shading:{fill:"FFFFFF",type:ShadingType.CLEAR},margins:{top:70,bottom:70,left:120,right:120},borders:{top:none,bottom:btm,left:none,right:none},
          children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:v||"",font:"Calibri",size:17,color:"111111"})]})]})
      ]}))
    }));
    ch.push(gap(140));
  }

  // Notes
  if (d.notes) {
    ch.push(sectionLabel("Notes"));
    ch.push(new Paragraph({spacing:{before:0,after:160},children:[new TextRun({text:d.notes,font:"Calibri",size:16,color:"444444"})]}));
  }

  // Auth (PO only)
  if (!isInv) {
    ch.push(gap(240));
    ch.push(sectionLabel("Authorisation"));
    ch.push(new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[3120,3120,3120],
      borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
      rows:[new TableRow({children:["Authorised Signature","Name & Title","Date"].map(label=>
        new TableCell({
          width:{size:3120,type:WidthType.DXA},
          margins:{top:60,bottom:320,left:0,right:60},
          borders:{top:none,bottom:{style:BorderStyle.SINGLE,size:4,color:"CCCCCC"},left:none,right:none},
          children:[new Paragraph({spacing:{before:0,after:0},children:[new TextRun({text:label,font:"Calibri",size:14,color:MID})]})]
        })
      )})]
    }));
  }

  return new Document({
    styles:{default:{document:{run:{font:"Calibri",size:18}}}},
    sections:[{
      properties:{page:{size:{width:11906,height:16838},margin:{top:1134,right:1134,bottom:1134,left:1134}}},
      children:ch
    }]
  });
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin","*");
  res.setHeader("Access-Control-Allow-Methods","POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers","Content-Type");
  if (req.method==="OPTIONS") return res.status(204).end();
  if (req.method!=="POST") return res.status(405).json({error:"POST only"});

  try {
    const { data } = req.body;
    if (!data) return res.status(400).json({error:"No data provided"});

    const results = {};
    if (data.po) {
      const doc = buildDoc(data.po,"PO");
      const buf = await Packer.toBuffer(doc);
      results.po_docx = buf.toString("base64");
    }
    if (data.inv) {
      const doc = buildDoc(data.inv,"INVOICE");
      const buf = await Packer.toBuffer(doc);
      results.inv_docx = buf.toString("base64");
    }

    res.status(200).json({ok:true,files:results});
  } catch(err) {
    console.error(err);
    res.status(500).json({error:err.message});
  }
};
