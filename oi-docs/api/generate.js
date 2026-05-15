const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign, ImageRun,
} = require('docx');
const fs = require('fs');
const path = require('path');

const NAVY="1B2A4A", LTGRAY="F2F2F2", MID="666666", WHITE="FFFFFF", BLACK="000000";
const none={style:BorderStyle.NONE,size:0,color:WHITE};
const thin={style:BorderStyle.SINGLE,size:4,color:"CCCCCC"};
const thk ={style:BorderStyle.SINGLE,size:12,color:NAVY};

function run(text,{bold=false,sz=18,color=BLACK,italic=false}={}) {
  return new TextRun({text:String(text||''),font:'Arial',size:sz,bold,color,italics:italic});
}

function cell(content,{w=2000,bg=WHITE,bold=false,sz=18,color=BLACK,align='left',borders={},colspan=1}={}) {
  const brd = Object.assign({top:thin,bottom:thin,left:thin,right:thin},borders);
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:60,bottom:60,left:100,right:100},
    verticalAlign:VerticalAlign.CENTER,
    borders:brd, columnSpan:colspan,
    children:[new Paragraph({
      alignment:align==='right'?AlignmentType.RIGHT:align==='center'?AlignmentType.CENTER:AlignmentType.LEFT,
      spacing:{before:20,after:20},
      children:[run(content,{bold,sz,color})]
    })]
  });
}

function gap(sz=6) {
  return new Paragraph({spacing:{before:0,after:0},children:[run('',{sz})]});
}

function getLogo() {
  try { return fs.readFileSync(path.join(__dirname,'..','oi_logo.png')); } catch(e) { return null; }
}

const money = v => '$'+parseFloat(v||0).toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2});

function gstLabel(d) {
  const rate = parseFloat(d.gst_rate||0.10);
  return 'GST ('+(rate*100|0)+'%)';
}

// ════════════════════════════════════════════════════════════════════════════
// PURCHASE ORDER
// ════════════════════════════════════════════════════════════════════════════
function buildPO(d) {
  const logoData = getLogo();
  const children = [];

  // Header
  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[3500,5860],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[new TableRow({children:[
      new TableCell({width:{size:3500,type:WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},margins:{top:0,bottom:0,left:0,right:0},
        children:[logoData
          ? new Paragraph({spacing:{before:0,after:0},children:[new ImageRun({data:logoData,transformation:{width:150,height:72},type:'png'})]})
          : new Paragraph({spacing:{before:0,after:0},children:[run('OCEAN INFINITY',{bold:true,sz:24,color:NAVY})]})
        ]}),
      new TableCell({width:{size:5860,type:WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},margins:{top:0,bottom:0,left:0,right:0},
        children:[
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('Ocean Infinity (Australia) Pty Ltd',{bold:true,sz:17,color:NAVY})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('2/237 Kennedy Drive',{sz:16,color:MID})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('Cambridge TAS 7170',{sz:16,color:MID})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('AUSTRALIA',{sz:16,color:MID})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run('Invoices: jo.eagle@oceaninfinity.com',{sz:15,color:MID,italic:true})]}),
        ]}),
    ]})]
  }));
  children.push(gap(8));

  // Title bar
  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[5860,3500],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[new TableRow({children:[
      new TableCell({width:{size:5860,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},borders:{top:none,bottom:none,left:none,right:none},margins:{top:80,bottom:80,left:160,right:0},
        children:[
          new Paragraph({spacing:{before:0,after:0},children:[run('Purchase Order Form',{bold:true,sz:24,color:WHITE})]}),
          new Paragraph({spacing:{before:0,after:0},children:[run('Document No: COM-FOR-003',{sz:16,color:"AABBCC"})]}),
        ]}),
      new TableCell({width:{size:3500,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},borders:{top:none,bottom:none,left:none,right:none},margins:{top:80,bottom:80,left:0,right:160},
        children:[
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run('Purchase Order No:',{sz:16,color:"AABBCC"})]}),
          new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(d.number||'',{bold:true,sz:22,color:WHITE})]}),
        ]}),
    ]})]
  }));
  children.push(gap(10));

  // Parties meta table
  const rows2 = [
    [['Supplier address:',true],['',false],['Purchase Order No:',true],[d.number||'',false,true]],
    [[d.supplier_name||'',false],['',false],['Supplier Quote No:',true],[d.ref||'',false]],
    [[(d.supplier_addr||''),false],['',false],['Supplier Contact:',true],[d.supplier_contact||'',false]],
    [['',false],['',false],['Supplier Email:',true],[d.supplier_email||'',false]],
    [['',false],['',false],['Account No:',true],['',false]],
    [['Delivery Address:',true],['',false],['Date of Order:',true],[d.date||'',false]],
    [['Ocean Infinity (Australia) Pty Ltd',false],['',false],['Ordered By:',true],[d.buyer_contact||'Ian Coffey',false]],
    [['2/237 Kennedy Drive, Cambridge TAS 7170',false],['',false],['Date Delivery Required:',true],[d.delivery||'',false]],
    [['',false],['',false],['Freight Arrangements:',true],['N/A',false]],
    [['',false],['',false],['Currency:',true],['AUD',false]],
    [['',false],['',false],['Price Includes GST?',true],['Yes',false]],
  ];
  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[1700,2500,1700,3460],
    rows: rows2.map((row,ri)=>new TableRow({children:[
      cell(row[0][0],{w:1700,bold:row[0][1],sz:16,bg:row[0][1]?LTGRAY:WHITE,borders:{top:thin,bottom:ri===rows2.length-1?thk:thin,left:thk,right:thin}}),
      cell(row[1][0],{w:2500,sz:16,borders:{top:thin,bottom:ri===rows2.length-1?thk:thin,left:thin,right:thin}}),
      cell(row[2][0],{w:1700,bold:true,sz:16,bg:LTGRAY,borders:{top:ri===0?thk:thin,bottom:ri===rows2.length-1?thk:thin,left:thin,right:thin}}),
      cell(row[3][0],{w:3460,bold:!!row[3][2],sz:row[3][2]?17:16,borders:{top:ri===0?thk:thin,bottom:ri===rows2.length-1?thk:thin,left:thin,right:thk}}),
    ]}))
  }));
  children.push(gap(10));

  // Line items
  const iHdr = new TableRow({children:[
    cell('Item',{w:500,bold:true,sz:16,bg:NAVY,color:WHITE,align:'center',borders:{top:thk,bottom:thin,left:thk,right:thin}}),
    cell('Code',{w:800,bold:true,sz:16,bg:NAVY,color:WHITE,align:'center',borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Description',{w:4800,bold:true,sz:16,bg:NAVY,color:WHITE,borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Qty',{w:500,bold:true,sz:16,bg:NAVY,color:WHITE,align:'center',borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Unit Price',{w:1100,bold:true,sz:16,bg:NAVY,color:WHITE,align:'right',borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Amount',{w:1660,bold:true,sz:16,bg:NAVY,color:WHITE,align:'right',borders:{top:thk,bottom:thin,left:thin,right:thk}}),
  ]});

  const iRows = (d.items||[]).map((item,i)=>{
    const bg=i%2===0?WHITE:LTGRAY;
    return new TableRow({children:[
      cell(String(i+1),{w:500,sz:16,align:'center',bg,borders:{top:thin,bottom:thin,left:thk,right:thin}}),
      cell('',{w:800,sz:16,bg,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
      new TableCell({width:{size:4800,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:100,right:100},borders:{top:thin,bottom:thin,left:thin,right:thin},
        children:[
          new Paragraph({spacing:{before:0,after:item.bullet_points?.length?40:0},children:[run(item.desc||'',{sz:17,bold:true})]}),
          ...(item.bullet_points||[]).map(bp=>new Paragraph({spacing:{before:0,after:20},indent:{left:160},children:[run('• '+bp,{sz:15,color:MID})]})),
          ...(item.effort?[new Paragraph({spacing:{before:item.bullet_points?.length?40:0,after:0},children:[run(item.effort,{sz:15,color:MID})]})]:[]),
        ]}),
      cell('1',{w:500,sz:16,align:'center',bg,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
      cell(money(item.amount),{w:1100,sz:16,align:'right',bg,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
      cell(money(item.amount),{w:1660,sz:16,align:'right',bg,borders:{top:thin,bottom:thin,left:thin,right:thk}}),
    ]});
  });

  const eRows = Array(3).fill(null).map((_,i)=>new TableRow({children:[
    cell('',{w:500,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thk,right:thin}}),
    cell('',{w:800,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:4800,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:500,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:1100,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:1660,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thk}}),
  ]}));

  const tRows=[['Sub-Total',d.subtotal,false],[gstLabel(d),d.gst,false],['TOTAL (incl. GST)',d.total,true]].map(([lbl,val,big])=>new TableRow({children:[
    new TableCell({width:{size:6600,type:WidthType.DXA},columnSpan:4,shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thk,right:thin},margins:{top:60,bottom:60,left:100,right:100},children:[new Paragraph({children:[]})]}),
    new TableCell({width:{size:1100,type:WidthType.DXA},shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thin,right:thin},margins:{top:60,bottom:60,left:100,right:100},children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(lbl,{bold:true,sz:big?18:16,color:big?WHITE:BLACK})]})]}),
    new TableCell({width:{size:1660,type:WidthType.DXA},shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thin,right:thk},margins:{top:60,bottom:60,left:100,right:100},children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(money(val),{bold:true,sz:big?19:16,color:big?WHITE:BLACK})]})]}),
  ]}));

  children.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[500,800,4800,500,1100,1660],rows:[iHdr,...iRows,...eRows,...tRows]}));
  children.push(gap(10));

  // Notes / sign-off
  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},columnWidths:[4680,4680],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[new TableRow({children:[
      new TableCell({width:{size:4680,type:WidthType.DXA},borders:{top:thk,bottom:thk,left:thk,right:thin},margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({spacing:{before:0,after:40},children:[run('Special Requirements:',{bold:true,sz:16})]}),
          new Paragraph({spacing:{before:0,after:0},children:[run(d.notes||'The PO number should be clearly shown on all documents.',{sz:15,color:MID})]})]}),
      new TableCell({width:{size:4680,type:WidthType.DXA},borders:{top:thk,bottom:thk,left:thin,right:thk},margins:{top:80,bottom:80,left:120,right:120},
        children:[new Paragraph({spacing:{before:0,after:40},children:[run('Important:',{bold:true,sz:16})]}),
          new Paragraph({spacing:{before:0,after:60},children:[run('The PO number '+(d.number||'')+' should be clearly shown on all documents.',{sz:15,color:MID})]}),
          new Paragraph({spacing:{before:0,after:0},children:[run('Signed: ',{bold:true,sz:16}),run('Peter Locke',{sz:16})]}),
          new Paragraph({spacing:{before:0,after:0},children:[run('Position: ',{bold:true,sz:16}),run('Managing Director',{sz:16})]}),
        ]}),
    ]})]
  }));
  children.push(gap(6));
  children.push(new Paragraph({spacing:{before:0,after:0},children:[run('DOCUMENT NO: COM-FOR-003  |  REVISION: 1  |  This document is UNCONTROLLED when printed',{sz:13,color:"AAAAAA",italic:true})]}));

  return new Document({styles:{default:{document:{run:{font:'Arial',size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:720,bottom:720,left:720}}},children}]});
}

// ════════════════════════════════════════════════════════════════════════════
// TAX INVOICE
// ════════════════════════════════════════════════════════════════════════════
function buildInvoice(d) {
  const logoData = getLogo();
  const children = [];

  // Header: supplier left, logo right
  const suppLines = [
    {text:d.supplier_name||'ETM Perspectives Pty Ltd', bold:true, sz:20},
    {text:d.supplier_addr||'PO Box 66, Kettering Tasmania 7155', bold:false, sz:16},
    {text:'Mobile: '+(d.supplier_phone||'0429 999 314'), bold:false, sz:16},
    {text:'Email: '+(d.supplier_email||'duane@etmp.com.au'), bold:false, sz:16},
    {text:'ACN: '+(d.supplier_acn||'112 806 121'), bold:false, sz:16},
    {text:'ABN: '+(d.supplier_abn||'37 112 806 121'), bold:false, sz:16},
  ];

  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},columnWidths:[5000,4360],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[new TableRow({children:[
      new TableCell({width:{size:5000,type:WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},margins:{top:0,bottom:0,left:0,right:0},
        children:suppLines.map((l,i)=>new Paragraph({spacing:{before:0,after:i===0?40:30},children:[run(l.text,{bold:l.bold,sz:l.sz,color:i===0?NAVY:MID})]}))}),
      new TableCell({width:{size:4360,type:WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},margins:{top:0,bottom:0,left:0,right:0},
        children:[logoData
          ? new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[new ImageRun({data:logoData,transformation:{width:140,height:67},type:'png'})]})
          : new Paragraph({children:[]})
        ]}),
    ]})]
  }));

  children.push(gap(200));

  // TAX INVOICE title
  children.push(new Paragraph({alignment:AlignmentType.CENTER,spacing:{before:0,after:160},
    children:[run('TAX INVOICE:  '+(d.number||''),{bold:true,sz:28,color:NAVY})]}));

  children.push(gap(80));

  // To / Date / Reference / PO
  [
    ['To:', [[d.buyer_contact||'Ian Coffey'],[d.buyer_name||'Ocean Infinity (Australia) Pty Ltd'],[d.buyer_addr||'2/237 Kennedy Drive, Cambridge TAS 7170'],['Australia']]],
    ['Date:', [[d.date||'']]],
    ['Reference:', [[d.reference||'']]],
    ['Purchase Order:', [[d.po_ref||'']]],
  ].forEach(([label,lines])=>{
    children.push(new Table({
      width:{size:9360,type:WidthType.DXA},columnWidths:[1800,7560],
      borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
      rows:[new TableRow({children:[
        new TableCell({width:{size:1800,type:WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},margins:{top:40,bottom:40,left:0,right:0},
          children:[new Paragraph({spacing:{before:0,after:0},children:[run(label,{bold:true,sz:18})]})]
        }),
        new TableCell({width:{size:7560,type:WidthType.DXA},borders:{top:none,bottom:none,left:none,right:none},margins:{top:40,bottom:40,left:0,right:0},
          children:lines.map(l=>new Paragraph({spacing:{before:0,after:0},children:[run(l[0],{sz:18})]}))
        }),
      ]})]
    }));
  });

  children.push(gap(120));

  // Heading
  children.push(new Paragraph({spacing:{before:0,after:80},
    children:[run('TAX INVOICE FOR PROFESSIONAL SERVICES by ETM Perspectives as per quote',{bold:true,sz:17})]}));

  // Items table
  const iHdr = new TableRow({children:[
    new TableCell({width:{size:5400,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:thk,bottom:thin,left:thk,right:thin},children:[new Paragraph({spacing:{before:0,after:0},children:[run('Deliverables',{bold:true,sz:17,color:WHITE})]})]}),
    new TableCell({width:{size:2360,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:thk,bottom:thin,left:thin,right:thin},children:[new Paragraph({spacing:{before:0,after:0},children:[run('Days Effort',{bold:true,sz:17,color:WHITE})]})]}),
    new TableCell({width:{size:1600,type:WidthType.DXA},shading:{fill:NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:thk,bottom:thin,left:thin,right:thk},children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run('Amount (GST Excl.)',{bold:true,sz:17,color:WHITE})]})]}),
  ]});

  const iRows = (d.items||[]).map((item,i)=>{
    const bg=i%2===0?WHITE:LTGRAY;
    return new TableRow({children:[
      new TableCell({width:{size:5400,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:thin,bottom:thin,left:thk,right:thin},
        children:[
          new Paragraph({spacing:{before:0,after:20},children:[run(item.desc||'',{bold:true,sz:17})]}),
          ...(item.bullet_points||[]).map(bp=>new Paragraph({spacing:{before:0,after:20},indent:{left:160},children:[run('• '+bp,{sz:16,color:MID})]})),
        ]}),
      new TableCell({width:{size:2360,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:thin,bottom:thin,left:thin,right:thin},
        children:[new Paragraph({spacing:{before:0,after:0},children:[run(item.effort||'',{sz:16,color:MID})]})]}),
      new TableCell({width:{size:1600,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:120,right:120},borders:{top:thin,bottom:thin,left:thin,right:thk},
        children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(money(item.amount),{sz:17})]})]}),
    ]});
  });

  const tRows=[['Sub-Total',d.subtotal,false],[gstLabel(d),d.gst,false],['TOTAL',d.total,true]].map(([lbl,val,big])=>new TableRow({children:[
    new TableCell({width:{size:7760,type:WidthType.DXA},columnSpan:2,shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thk,right:thin},margins:{top:60,bottom:60,left:120,right:120},
      children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(lbl,{bold:true,sz:big?19:17,color:big?WHITE:BLACK})]})]}),
    new TableCell({width:{size:1600,type:WidthType.DXA},shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thin,right:thk},margins:{top:60,bottom:60,left:120,right:120},
      children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(money(val),{bold:true,sz:big?20:17,color:big?WHITE:BLACK})]})]}),
  ]}));

  children.push(new Table({width:{size:9360,type:WidthType.DXA},columnWidths:[5400,2360,1600],rows:[iHdr,...iRows,...tRows]}));
  children.push(gap(160));

  // Bank details
  [['Account Name:',d.bank_name||'ETM Perspectives Pty Ltd'],
   ['Account Type:','Business Account'],
   ['BSB Number:',d.bank_bsb||'034093'],
   ['Account Number:',d.bank_acct||'252236'],
   ['Branch:',d.bank_branch||'Westpac, Garden City Mt Gravatt, Logan & Kessels Roads, Upper Mt Gravatt, Qld 4122'],
  ].forEach(([k,v])=>{
    children.push(new Paragraph({spacing:{before:0,after:40},children:[run(k+'  ',{bold:true,sz:17}),run(v,{sz:17})]}));
  });

  children.push(gap(120));
  children.push(new Paragraph({spacing:{before:0,after:40},children:[run(d.signed_by||'Duane Vickery',{sz:17})]}));
  children.push(new Paragraph({spacing:{before:0,after:40},children:[run('Managing Director',{sz:17})]}));
  children.push(new Paragraph({spacing:{before:0,after:0},children:[run(d.supplier_name||'ETM Perspectives Pty Ltd',{sz:17})]}));

  return new Document({styles:{default:{document:{run:{font:'Arial',size:18}}}},sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1000,right:1000,bottom:1000,left:1000}}},children}]});
}

// ════════════════════════════════════════════════════════════════════════════
// Handler
// ════════════════════════════════════════════════════════════════════════════
module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin","*");
  res.setHeader("Access-Control-Allow-Methods","POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers","Content-Type");
  if (req.method==="OPTIONS") return res.status(204).end();
  if (req.method!=="POST") return res.status(405).json({error:"POST only"});
  try {
    const { data, docType } = req.body;
    if (!data) return res.status(400).json({error:"No data"});
    const results = {};
    if (docType==='invoice' && data.inv) {
      const buf = await Packer.toBuffer(buildInvoice(data.inv));
      results.inv_docx = buf.toString("base64");
    } else if (data.po) {
      const buf = await Packer.toBuffer(buildPO(data.po));
      results.po_docx = buf.toString("base64");
    }
    res.status(200).json({ok:true,files:results});
  } catch(err) {
    console.error(err);
    res.status(500).json({error:err.message});
  }
};
