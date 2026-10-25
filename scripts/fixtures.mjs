// Original synthetic fixtures. Coordinates are known to this harness, not a blind AI evaluation.
import fs from 'node:fs/promises';
import PptxGenJS from 'pptxgenjs';
import * as ok from '@office-kit/pptx';
import { performance } from 'node:perf_hooks';
import sharp from 'sharp';
const W=1440,H=810,scale=108;const C={ink:'16324F',muted:'587088',blue:'245CEC',teal:'00A68C',paper:'F5F8FC'};
const texts=[];const pages=[];let ops=[];
function text(t,x,y,w,h,size=26,color=C.ink,bold=false){ops.push({kind:'text',t,x,y,w,h,size,color,bold});}
function shape(preset,x,y,w,h,fill,line=null,shadow=false,gradient=null){ops.push({kind:'shape',preset,x,y,w,h,fill,line,shadow,gradient});}
function line(x,y,x2,y2,color=C.muted,width=2,arrow=false){ops.push({kind:'line',x,y,x2,y2,color,width,arrow});}
function base(title,sub){ops=[];shape('rect',0,0,W,H,C.paper);text('SLIDERECAST / ORIGINAL TEST MATERIAL',64,42,1240,30,17,C.muted,true);text(title,64,91,1300,65,42,C.ink,true);text(sub,64,166,1270,45,22,C.muted);}
function finish(name){pages.push({name,ops});texts.push(ops.filter(o=>o.kind==='text').map(o=>o.t));}
base('Rebuild the visual. Keep the meaning.','Fidelity-first reconstruction / 高保真視覺重建');
shape('roundRect',64,250,420,425,C.blue,null,true,[C.blue,'5639C7']);
text('01',94,279,330,75,45,'FFFFFF',true);text('See the original',94,374,344,60,32,'FFFFFF',true);text('版面、字型、色彩',94,458,344,44,24,'FFFFFF');
shape('ellipse',345,285,90,90,'FFFFFF');shape('ellipse',365,305,50,50,C.blue);shape('rect',384,307,12,46,'FFFFFF');
shape('roundRect',525,250,399,425,'FFFFFF',null,true);text('02',557,279,300,75,45,C.teal,true);text('Recreate layers',557,374,325,60,30,C.ink,true);text('Native text and shapes',557,458,324,42,22,C.muted);text('主要元素可個別編輯',557,514,320,42,22,C.muted);
for(let i=0;i<3;i++)shape('roundRect',825-i*12,299+i*14,60,40,['DDE8FF','80A4F8',C.blue][i]);
shape('roundRect',963,250,413,425,'FFFFFF',null,true);text('03',995,279,300,75,45,C.blue,true);text('Compare & correct',995,374,354,60,29,C.ink,true);text('Render before delivery',995,458,344,42,22,C.muted);line(1002,575,1100,575,C.teal,5);line(1100,575,1080,557,C.teal,5);
text('Source designed independently for testing. No client or company documents.',64,738,1270,35,17,C.muted);finish('infographic');
base('From pixels to editable layers','Complex screenshot fixture / 原始外觀優先');
const nodes=[[90,320,'Inspect',C.blue],[435,320,'Rebuild',C.teal],[780,320,'Render','7953CF'],[1125,320,'Correct','D67A24']];
for(let i=0;i<3;i++)line(nodes[i][0]+235,420,nodes[i+1][0]-12,420,C.muted,3,true);
for(let i=0;i<nodes.length;i++){const [x,y,label,color]=nodes[i];shape('roundRect',x,y,235,205,'FFFFFF',null,true);shape('ellipse',x+82,y+25,70,70,color);if(i===0){shape('ellipse',x+99,y+42,30,30,'FFFFFF');line(x+123,y+64,x+139,y+80,'FFFFFF',4);}else if(i===1){for(let k=0;k<3;k++)shape('rect',x+103-k*7,y+41+k*9,30,20,k===2?'FFFFFF':'A8E5DC');}else if(i===2){shape('rect',x+97,y+46,40,25,'FFFFFF');line(x+117,y+72,x+117,y+81,'FFFFFF',3);}else{line(x+102,y+57,x+114,y+70,'FFFFFF',4);line(x+114,y+70,x+136,y+48,'FFFFFF',4);}text(label,x+15,y+121,205,51,26,C.ink,true);}
line(1243,540,1243,620,C.muted,2);line(1243,620,550,620,C.muted,2);line(550,620,550,547,C.muted,2,true);text('Targeted correction loop',636,650,560,46,22,C.muted);text('文字保持原樣',90,731,1100,45,20,C.ink);finish('complex-flow');
base('A chart is also a visual design','Default: editable shapes and labels. Native chart data only on request.');
shape('roundRect',70,255,920,450,'FFFFFF',null,true);text('Observed sample values',110,285,790,50,27,C.ink,true);
for(let i=0;i<4;i++){let y=390+i*78;line(120,y,930,y,'E4EAF2',1);}
for(const [i,v] of [32,58,76,92].entries()){let x=177+i*176;let h=v*3;shape('roundRect',x,650-h,82,h,[C.blue,C.teal,'7953CF','D67A24'][i]);text(String(v),x-10,610-h,110,36,24,C.ink,true);text(['Q1','Q2','Q3','Q4'][i],x-10,660,110,33,21,C.muted);}
text('Exact labels',1050,350,290,50,29,C.ink,true);text('32 / 58 / 76 / 92',1050,425,300,44,25,C.blue,true);text('No inferred data',1050,506,300,44,22,C.muted);text('No style changes',1050,561,300,44,22,C.muted);finish('visual-chart');
await fs.mkdir('examples/source',{recursive:true});await fs.mkdir('examples/results',{recursive:true});
function esc(v){return v.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');}
function svg(page){let s=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><defs><filter id="sh" x="-25%" y="-25%" width="150%" height="160%"><feDropShadow dx="0" dy="5" stdDeviation="6" flood-color="#16324F" flood-opacity=".13"/></filter><linearGradient id="grad"><stop stop-color="#245CEC"/><stop offset="1" stop-color="#5639C7"/></linearGradient></defs>`;for(const o of page.ops){if(o.kind==='text'){s+=`<text x="${o.x}" y="${o.y+o.size*1.12}" font-family="Noto Sans CJK TC" font-size="${o.size}" font-weight="${o.bold?700:400}" fill="#${o.color}">${esc(o.t)}</text>`;}else if(o.kind==='line'){s+=`<path d="M${o.x},${o.y} L${o.x2},${o.y2}" fill="none" stroke="#${o.color}" stroke-width="${o.width}"/>`;if(o.arrow)s+=`<path d="M${o.x2-9},${o.y2-6} L${o.x2},${o.y2} L${o.x2-9},${o.y2+6}" fill="none" stroke="#${o.color}" stroke-width="${o.width}"/>`;}else{let attr=`fill="${o.gradient?'url(#grad)':'#'+o.fill}" ${o.shadow?'filter="url(#sh)"':''}`;s+=o.preset==='ellipse'?`<ellipse cx="${o.x+o.w/2}" cy="${o.y+o.h/2}" rx="${o.w/2}" ry="${o.h/2}" ${attr}/>`:`<rect x="${o.x}" y="${o.y}" width="${o.w}" height="${o.h}" rx="${o.preset==='roundRect'?12:0}" ${attr}/>`;}}return s+'</svg>';}
for(const page of pages){const raw=svg(page);await fs.writeFile(`examples/source/${page.name}.svg`,raw);await sharp(Buffer.from(raw)).png().toFile(`examples/source/${page.name}.png`);}
await fs.writeFile('examples/source/expected-text.json',JSON.stringify(texts,null,2)+'\n');
const timing=[];
for(const engine of ['pptxgenjs','office-kit']){
 const start=performance.now();let pres=engine==='pptxgenjs'?new PptxGenJS():ok.createPresentation();
 if(engine==='pptxgenjs'){pres.defineLayout({name:'SOURCE',width:W/scale,height:H/scale});pres.layout='SOURCE';pres.author='SlideRecast';pres.subject='Original synthetic test fixtures';}else ok.setSlideSize(pres,{width:ok.inches(W/scale),height:ok.inches(H/scale)});
 for(const page of pages){const slide=engine==='pptxgenjs'?pres.addSlide():ok.addBlankSlide(pres);for(const [index,o] of page.ops.flatMap(o=>o.kind==='line'&&o.arrow?[{...o,arrow:false},{kind:'line',x:o.x2-9,y:o.y2-6,x2:o.x2,y2:o.y2,color:o.color,width:o.width},{kind:'line',x:o.x2,y:o.y2,x2:o.x2-9,y2:o.y2+6,color:o.color,width:o.width}]:[o]).entries()){
  const name=`${page.name}-${o.kind}-${index}`;
  if(engine==='pptxgenjs'){
   if(o.kind==='text')slide.addText(o.t,{x:o.x/scale,y:o.y/scale,w:o.w/scale,h:o.h/scale,fontFace:'Noto Sans CJK TC',fontSize:o.size*72/scale,color:o.color,bold:o.bold,margin:0,breakLine:false,valign:'top',objectName:name});
   else if(o.kind==='line')slide.addShape(pres.ShapeType.line,{x:o.x/scale,y:o.y/scale,w:(o.x2-o.x)/scale,h:(o.y2-o.y)/scale,line:{color:o.color,width:o.width*72/scale,...(o.arrow?{endArrowType:'triangle'}:{})},objectName:name});
   else {
    slide.addShape(pres.ShapeType[o.preset],{x:o.x/scale,y:o.y/scale,w:o.w/scale,h:o.h/scale,rectRadius:12/scale,radius:12/scale,fill:{color:o.fill},line:{color:o.fill,transparency:100},...(o.shadow?{shadow:{type:'outer',color:C.ink,opacity:.13,blur:4,offset:3,angle:90}}:{}),objectName:name});
    // PptxGenJS has no tested gradient-fill API here. Use separate native bands,
    // preserving editability and explicitly report the approximation.
    if(o.gradient){for(let j=0;j<80;j++){const t=j/79;const color=[0,2,4].map(k=>Math.round(parseInt(o.gradient[0].slice(k,k+2),16)*(1-t)+parseInt(o.gradient[1].slice(k,k+2),16)*t).toString(16).padStart(2,'0')).join('');slide.addShape(pres.ShapeType.rect,{x:(o.x+12+j*(o.w-24)/80)/scale,y:o.y/scale,w:(o.w-24)/80/scale,h:o.h/scale,fill:{color},line:{transparency:100,color},objectName:`gradient-band-${j}`});}}
   }
  } else {
   let sh;
   const box={x:ok.inches(o.x/scale),y:ok.inches(o.y/scale),w:ok.inches((o.w||1)/scale),h:ok.inches((o.h||1)/scale),name};
   if(o.kind==='text'){sh=ok.addSlideTextBox(slide,{...box,text:o.t});ok.setShapeTextFormat(sh,{font:'Noto Sans CJK TC',fontEastAsian:'Noto Sans CJK TC',size:o.size*72/scale,color:'#'+o.color,bold:o.bold});ok.setShapeTextMargins(sh,{left:0,right:0,top:0,bottom:0});ok.setShapeTextAnchor(sh,'top');}
   else if(o.kind==='line'){sh=ok.addSlideLine(slide,{from:{x:ok.inches(o.x/scale),y:ok.inches(o.y/scale)},to:{x:ok.inches(o.x2/scale),y:ok.inches(o.y2/scale)},color:'#'+o.color,widthEmu:ok.pt(o.width*72/scale),name});if(o.arrow)ok.setShapeStrokeArrow(sh,'tail',{type:'triangle'});}
   else {sh=ok.addSlideShape(slide,{...box,preset:o.preset});ok.setShapeFill(sh,'#'+o.fill);ok.clearShapeStroke(sh);if(o.preset==='roundRect')ok.setShapeAdjustValues(sh,{adj:Math.round(12/Math.min(o.w,o.h)*100000)});if(o.gradient)ok.setShapeGradientFill(sh,{angleDeg:0,stops:o.gradient.map((c,i)=>({offset:i,color:'#'+c}))});if(o.shadow)ok.setShapeShadow(sh,{color:'#'+C.ink,opacity:.13,blurEmu:ok.pt(4),offsetEmu:ok.pt(3),angleDeg:90});}
  }
 }}
 const file=`examples/results/${engine}.pptx`;if(engine==='pptxgenjs')await pres.writeFile({fileName:file});else await fs.writeFile(file,await ok.savePresentation(pres));
 timing.push({engine,generation_seconds:(performance.now()-start)/1000,bytes:(await fs.stat(file)).size});
}
// Optional native-chart capability probe, intentionally outside fidelity examples.
for(const engine of ['pptxgenjs','office-kit']){
 const pres=engine==='pptxgenjs'?new PptxGenJS():ok.createPresentation();const slide=engine==='pptxgenjs'?pres.addSlide():ok.addBlankSlide(pres);
 if(engine==='pptxgenjs'){pres.layout='LAYOUT_WIDE';slide.addText('Optional editable data probe',{x:.6,y:.4,w:11,h:.6,fontSize:26});slide.addChart(pres.ChartType.bar,[{name:'Sample',labels:['Q1','Q2','Q3','Q4'],values:[32,58,76,92]}],{x:1,y:1.5,w:10,h:4.7,barDir:'col',showValue:true,chartColors:[C.blue]});await pres.writeFile({fileName:`examples/results/${engine}-native-chart.pptx`});}
 else {const t=ok.addSlideTextBox(slide,{x:ok.inches(.6),y:ok.inches(.4),w:ok.inches(11),h:ok.inches(.6),text:'Optional editable data probe'});ok.setShapeTextFormat(t,{size:26});ok.addSlideChart(slide,{x:ok.inches(1),y:ok.inches(1.5),w:ok.inches(10),h:ok.inches(4.7),spec:{kind:'column',categories:['Q1','Q2','Q3','Q4'],series:[{name:'Sample',values:[32,58,76,92]}]}});await fs.writeFile(`examples/results/${engine}-native-chart.pptx`,await ok.savePresentation(pres));}
}
await fs.writeFile('examples/results/generation.json',JSON.stringify(timing,null,2)+'\n');console.log(timing);
