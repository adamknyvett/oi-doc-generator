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
  const para = new Paragraph({
    alignment:align==='right'?AlignmentType.RIGHT:align==='center'?AlignmentType.CENTER:AlignmentType.LEFT,
    spacing:{before:20,after:20},
    children:[run(content,{bold,sz,color})]
  });
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    shading:{fill:bg,type:ShadingType.CLEAR},
    margins:{top:60,bottom:60,left:100,right:100},
    verticalAlign:VerticalAlign.CENTER,
    borders:brd, columnSpan:colspan,
    children:[para]
  });
}

function gap(sz=6) {
  return new Paragraph({spacing:{before:0,after:0},children:[run('',{sz})]});
}

function buildPO(d) {
  // Load OI logo from bundled file
  let logoData;
  try {
    logoData = fs.readFileSync(path.join(__dirname,'..','oi_logo.png'));
  } catch(e) {
    logoData = null;
  }

  const items = d.items || [];
  const children = [];

  // ── Header: logo + address ──────────────────────────────────────────────
  const logoCell = new TableCell({
    width:{size:3500,type:WidthType.DXA},
    borders:{top:none,bottom:none,left:none,right:none},
    margins:{top:0,bottom:0,left:0,right:0},
    children:[logoData
      ? new Paragraph({spacing:{before:0,after:0},children:[new ImageRun({data:logoData,transformation:{width:150,height:72},type:'png'})]})
      : new Paragraph({spacing:{before:0,after:0},children:[run('OCEAN INFINITY',{bold:true,sz:24,color:NAVY})]})
    ]
  });

  const addrCell = new TableCell({
    width:{size:5860,type:WidthType.DXA},
    borders:{top:none,bottom:none,left:none,right:none},
    margins:{top:0,bottom:0,left:0,right:0},
    children:[
      new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('Ocean Infinity (Australia) Pty Ltd',{bold:true,sz:17,color:NAVY})]}),
      new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('2/237 Kennedy Drive',{sz:16,color:MID})]}),
      new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('Cambridge TAS 7170',{sz:16,color:MID})]}),
      new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:4},children:[run('AUSTRALIA',{sz:16,color:MID})]}),
      new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run('Invoices: jo.eagle@oceaninfinity.com',{sz:15,color:MID,italic:true})]}),
    ]
  });

  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[3500,5860],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[new TableRow({children:[logoCell,addrCell]})]
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

  // Parties meta
  const suppAddr = (d.supplier_addr||'').replace(/\\n/g,' ');
  const delivAddr = (d.buyer_addr||'2/237 Kennedy Drive, Cambridge TAS 7170').replace(/\\n/g,' ');
  const rows2 = [
    [['Supplier address:',true],  ['',false],  ['Purchase Order No:',true],  [d.number||'',false,true]],
    [[d.supplier_name||'',false], ['',false],  ['Supplier Quote No:',true],  [d.ref||'ETM Quote to Ocean Infinity 16Dec',false]],
    [[suppAddr,false],            ['',false],  ['Supplier Contact:',true],   [d.supplier_contact||'Ashlen Foster-Britton',false]],
    [['',false],                  ['',false],  ['Supplier Email:',true],     [d.supplier_email||'Ashlen@etmp.com.au',false]],
    [['',false],                  ['',false],  ['Account No:',true],         ['',false]],
    [['Delivery Address:',true],  ['',false],  ['Date of Order:',true],      [d.date||'',false]],
    [[delivAddr,false],           ['',false],  ['Ordered By:',true],         [d.buyer_contact||'Ian Coffey',false]],
    [['',false],                  ['',false],  ['Date Delivery Required:',true],[d.delivery||'',false]],
    [['',false],                  ['',false],  ['Freight Arrangements:',true],['N/A',false]],
    [['',false],                  ['',false],  ['Currency:',true],           ['AUD',false]],
    [['',false],                  ['',false],  ['Price Includes GST?',true], ['Yes',false]],
  ];
  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[1700,2500,1700,3460],
    rows: rows2.map((row,ri) => new TableRow({children:[
      cell(row[0][0],{w:1700,bold:row[0][1],sz:16,bg:row[0][1]?LTGRAY:WHITE,borders:{top:thin,bottom:ri===rows2.length-1?thk:thin,left:thk,right:thin}}),
      cell(row[1][0],{w:2500,sz:16,bg:row[1][1]?LTGRAY:WHITE,borders:{top:thin,bottom:ri===rows2.length-1?thk:thin,left:thin,right:thin}}),
      cell(row[2][0],{w:1700,bold:row[2][1],sz:16,bg:LTGRAY,borders:{top:ri===0?thk:thin,bottom:ri===rows2.length-1?thk:thin,left:thin,right:thin}}),
      cell(row[3][0],{w:3460,bold:!!row[3][2],sz:row[3][2]?17:16,borders:{top:ri===0?thk:thin,bottom:ri===rows2.length-1?thk:thin,left:thin,right:thk}}),
    ]}))
  }));
  children.push(gap(10));

  // Line items header
  const iHdr = new TableRow({children:[
    cell('Item',{w:500,bold:true,sz:16,bg:NAVY,color:WHITE,align:'center',borders:{top:thk,bottom:thin,left:thk,right:thin}}),
    cell('Code',{w:800,bold:true,sz:16,bg:NAVY,color:WHITE,align:'center',borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Description',{w:4800,bold:true,sz:16,bg:NAVY,color:WHITE,borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Qty',{w:500,bold:true,sz:16,bg:NAVY,color:WHITE,align:'center',borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Unit Price',{w:1100,bold:true,sz:16,bg:NAVY,color:WHITE,align:'right',borders:{top:thk,bottom:thin,left:thin,right:thin}}),
    cell('Amount',{w:1660,bold:true,sz:16,bg:NAVY,color:WHITE,align:'right',borders:{top:thk,bottom:thin,left:thin,right:thk}}),
  ]});

  const money = v => '$'+parseFloat(v||0).toLocaleString('en-AU',{minimumFractionDigits:2,maximumFractionDigits:2});

  const iRows = items.map((item,i) => {
    const bg = i%2===0?WHITE:LTGRAY;
    return new TableRow({children:[
      cell(String(i+1),{w:500,sz:16,align:'center',bg,borders:{top:thin,bottom:thin,left:thk,right:thin}}),
      cell('',{w:800,sz:16,bg,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
      new TableCell({width:{size:4800,type:WidthType.DXA},shading:{fill:bg,type:ShadingType.CLEAR},margins:{top:60,bottom:60,left:100,right:100},borders:{top:thin,bottom:thin,left:thin,right:thin},
        children:[
          new Paragraph({spacing:{before:0,after:20},children:[run(item.desc||'',{sz:17,bold:true})]}),
          ...(item.effort?[new Paragraph({spacing:{before:0,after:0},children:[run(item.effort,{sz:15,color:MID})]})]:[]),
        ]}),
      cell('1',{w:500,sz:16,align:'center',bg,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
      cell(money(item.amount),{w:1100,sz:16,align:'right',bg,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
      cell(money(item.amount),{w:1660,sz:16,align:'right',bg,borders:{top:thin,bottom:thin,left:thin,right:thk}}),
    ]});
  });

  // Empty rows
  const eRows = Array(3).fill(null).map((_,i)=>new TableRow({children:[
    cell('',{w:500,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thk,right:thin}}),
    cell('',{w:800,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:4800,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:500,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:1100,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thin}}),
    cell('',{w:1660,bg:i%2===0?LTGRAY:WHITE,borders:{top:thin,bottom:thin,left:thin,right:thk}}),
  ]}));

  // Totals
  const tRows = [
    ['Sub-Total',money(d.subtotal),false],
    ['GST (10%)',money(d.gst),false],
    ['TOTAL (incl. GST)',money(d.total),true],
  ].map(([lbl,val,big],i)=>new TableRow({children:[
    new TableCell({width:{size:6600,type:WidthType.DXA},columnSpan:4,shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thk,right:thin},margins:{top:60,bottom:60,left:100,right:100},children:[new Paragraph({children:[]})]}),
    new TableCell({width:{size:1100,type:WidthType.DXA},shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thin,right:thin},margins:{top:60,bottom:60,left:100,right:100},children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(lbl,{bold:true,sz:big?18:16,color:big?WHITE:BLACK})]})]}),
    new TableCell({width:{size:1660,type:WidthType.DXA},shading:{fill:big?NAVY:LTGRAY,type:ShadingType.CLEAR},borders:{top:thin,bottom:big?thk:thin,left:thin,right:thk},margins:{top:60,bottom:60,left:100,right:100},children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{before:0,after:0},children:[run(val,{bold:true,sz:big?19:16,color:big?WHITE:BLACK})]})]}),
  ]}));

  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[500,800,4800,500,1100,1660],
    rows:[iHdr,...iRows,...eRows,...tRows],
  }));
  children.push(gap(10));

  // Notes / sign-off
  children.push(new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[4680,4680],
    borders:{top:none,bottom:none,left:none,right:none,insideH:none,insideV:none},
    rows:[new TableRow({children:[
      new TableCell({width:{size:4680,type:WidthType.DXA},borders:{top:thk,bottom:thk,left:thk,right:thin},margins:{top:80,bottom:80,left:120,right:120},
        children:[
          new Paragraph({spacing:{before:0,after:40},children:[run('Special Requirements:',{bold:true,sz:16})]}),
          new Paragraph({spacing:{before:0,after:0},children:[run(d.notes||'The PO number should be clearly shown on all documents. Please acknowledge with delivery date.',{sz:15,color:MID})]}),
        ]}),
      new TableCell({width:{size:4680,type:WidthType.DXA},borders:{top:thk,bottom:thk,left:thin,right:thk},margins:{top:80,bottom:80,left:120,right:120},
        children:[
          new Paragraph({spacing:{before:0,after:40},children:[run('Important:',{bold:true,sz:16})]}),
          new Paragraph({spacing:{before:0,after:60},children:[run('The PO number '+( d.number||'')+' should be clearly shown on all documents.',{sz:15,color:MID})]}),
          new Paragraph({spacing:{before:0,after:0},children:[run('Signed: ',{bold:true,sz:16}),run('Peter Locke',{sz:16})]}),
          new Paragraph({spacing:{before:0,after:0},children:[run('Position: ',{bold:true,sz:16}),run('General Manager',{sz:16})]}),
        ]}),
    ]})]
  }));
  children.push(gap(6));
  children.push(new Paragraph({spacing:{before:0,after:0},children:[run('DOCUMENT NO: COM-FOR-003  |  REVISION: 1  |  This document is UNCONTROLLED when printed',{sz:13,color:"AAAAAA",italic:true})]}));

  return new Document({
    styles:{default:{document:{run:{font:'Arial',size:18}}}},
    sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:720,bottom:720,left:720}}},children}]
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
    if (!data) return res.status(400).json({error:"No data"});

    const results = {};
    if (data.po) {
      const doc = buildPO(data.po);
      const buf = await Packer.toBuffer(doc);
      results.po_docx = buf.toString("base64");
    }
    res.status(200).json({ok:true,files:results});
  } catch(err) {
    console.error(err);
    res.status(500).json({error:err.message});
  }
};
