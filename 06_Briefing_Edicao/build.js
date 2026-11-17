const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,ShadingType,AlignmentType,BorderStyle,PageBreak}=require('docx');
const W=9360; // 6.5in
const F='Calibri';
const DARK='1F4E79', LIGHT='DCE6F1', PALE='F3F6FA';
const noB={style:BorderStyle.NONE,size:0,color:'FFFFFF'};
const thin={style:BorderStyle.SINGLE,size:6,color:'9FC1E8'};
function run(t,o={}){return new TextRun({text:t,font:F,size:o.size||21,bold:o.bold,italics:o.italics,color:o.color});}
function p(text,o={}){ // text may be string or array of runs
  const children=Array.isArray(text)?text:[run(text,o)];
  return new Paragraph({children,spacing:{before:o.before??60,after:o.after??60},alignment:o.align});
}
function richP(segs,o={}){ // segs: [[text,{bold}],...]
  return new Paragraph({children:segs.map(s=>run(s[0],Object.assign({size:21},s[1]||{}))),spacing:{before:o.before??60,after:o.after??60}});
}
function bandRow(label){
  return new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],rows:[new TableRow({children:[new TableCell({width:{size:W,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:DARK,color:'auto'},margins:{top:100,bottom:100,left:140,right:140},children:[p([run(label,{bold:true,color:'FFFFFF',size:24})],{before:0,after:0})]})]})]});
}
function subBand(label){
  return new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],rows:[new TableRow({children:[new TableCell({width:{size:W,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:LIGHT,color:'auto'},margins:{top:80,bottom:80,left:140,right:140},children:[p([run(label,{bold:true,color:DARK,size:22})],{before:0,after:0})]})]})]});
}
function label(t,hint){return richP([[t,{bold:true,size:22}],hint?[' '+hint,{italics:true,color:'70757F',size:19}]:['',{}]],{before:140,after:40});}
function box(paras){ // paras: array of Paragraph or strings
  const ch=paras.map(x=>typeof x==='string'?p(x,{before:40,after:40}):x);
  return new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],rows:[new TableRow({children:[new TableCell({width:{size:W,type:WidthType.DXA},borders:{top:thin,bottom:thin,left:thin,right:thin},margins:{top:100,bottom:100,left:160,right:160},children:ch})]})]});
}
function grid(headers,rows,widths){
  const mk=(txt,hdr)=>new TableCell({width:{size:0,type:WidthType.DXA},borders:{top:thin,bottom:thin,left:thin,right:thin},shading:hdr?{type:ShadingType.CLEAR,fill:LIGHT,color:'auto'}:undefined,margins:{top:60,bottom:60,left:90,right:90},children:(Array.isArray(txt)?txt:[txt]).map(t=>p([run(t,{bold:hdr,size:hdr?19:18})],{before:20,after:20}))});
  const hrow=new TableRow({tableHeader:true,children:headers.map((h,i)=>{const c=mk(h,true);c.options.width={size:widths[i],type:WidthType.DXA};return c;})});
  const body=rows.map(r=>new TableRow({children:r.map((t,i)=>{const c=mk(t,false);c.options.width={size:widths[i],type:WidthType.DXA};return c;})}));
  return new Table({width:{size:W,type:WidthType.DXA},columnWidths:widths,rows:[hrow,...body]});
}
const gap=()=>p('',{before:0,after:0});
const D=JSON.parse(fs.readFileSync('content.json','utf8'));
const kids=[];
// Title band
kids.push(new Table({width:{size:W,type:WidthType.DXA},columnWidths:[W],rows:[new TableRow({children:[new TableCell({width:{size:W,type:WidthType.DXA},shading:{type:ShadingType.CLEAR,fill:DARK,color:'auto'},margins:{top:220,bottom:220,left:200,right:200},children:[p([run('TEMPLATE DE ORIENTAÇÃO DE EDIÇÃO',{bold:true,color:'FFFFFF',size:30})],{before:0,after:60}),p([run('VSL — VIDEO SALES LETTER',{bold:true,color:'FFFFFF',size:26})],{before:0,after:60}),p([run('Preenchido pelo Copywriter responsável antes do briefing com o editor e o coordenador de edição.',{italics:true,color:'E4ECF7',size:19})],{before:0,after:0})]})]})]}));
kids.push(gap());
kids.push(richP([['Como usar este documento: ',{bold:true,italics:true}],[D.howto,{italics:true}]],{before:80,after:120}));
for(const sec of D.sections){
  kids.push(gap()); kids.push(bandRow(sec.title)); kids.push(gap());
  for(const it of sec.items){
    if(it.sub){kids.push(subBand(it.sub)); kids.push(gap()); continue;}
    if(it.label) kids.push(label(it.label,it.hint));
    if(it.box) kids.push(box(it.box.map(x=>Array.isArray(x)?richP(x,{before:40,after:40}):x)));
    if(it.grid) kids.push(grid(it.grid.headers,it.grid.rows,it.grid.widths));
    kids.push(gap());
  }
}
const doc=new Document({
  styles:{default:{document:{run:{font:F,size:21}}}},
  sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1440,bottom:1440,left:1440,right:1440}}},children:kids}]
});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('BRIEFING_Edicao_StallionSoda.docx',b);console.log('ok',b.length);});
