import fs from 'node:fs/promises';
import * as p from '../../skills/image-to-editable-ppt/vendor/office-kit.mjs';
import {fileURLToPath} from 'node:url';
const OUT=fileURLToPath(new URL('.',import.meta.url));
const d=p.createPresentation();p.setSlideSize(d,{width:p.pt(960),height:p.pt(540)});
const U=n=>p.pt(n/1.5),C={bg:'#F5F8FC',ink:'#16324F',muted:'#587088',blue:'#245CEC',green:'#00A68C',purple:'#7953CF',orange:'#D67A24',white:'#FFFFFF'};
const expected=[];let s,txts;
// Targeted adjustments after inspecting all initial renders against the rasters.
// Keys refer only to native text objects; major geometry remains in place.
const adjust={
 'test-material-label':{dx:1,dy:4,size:17},
 'page-title':{dx:1,dy:6,size:42}, 'page-subtitle':{dx:1,dy:6,size:22},
 'step-01-number':{dx:1,dy:3,size:46},'step-02-number':{dx:1,dy:3,size:46},'step-03-number':{dx:1,dy:3,size:46},
 'step-01-title':{dy:5,size:32},'step-02-title':{dy:5,size:30},'step-03-title':{dy:5,size:29},
 'step-01-details':{dx:2,dy:8},'step-02-english':{dy:8},'step-02-chinese':{dx:1,dy:9},'step-03-details':{dx:1,dy:7},
 'source-footer':{dx:1,dy:7,size:17},
 'stage-Inspect':{dy:6,size:26},'stage-Rebuild':{dy:6,size:26},'stage-Render':{dy:6,size:26},'stage-Correct':{dy:6,size:26},
 'targeted-loop-label':{dy:7},'preserve-wording-note':{dx:1,dy:6},
 'chart-heading':{dy:8,size:27},'chart-note-title':{dy:8,size:29},
 'chart-exact-values':{dx:1,dy:8,size:25},'chart-no-inference':{dx:1,dy:7,size:22},'chart-no-style-changes':{dx:1,dy:7,size:22},
 ...Object.fromEntries([1,2,3,4].flatMap(i=>[['value-Q'+i,{dx:1,dy:6,size:24}],['quarter-Q'+i,{dx:1,dy:6,size:21}]]))
};
function slide(){s=p.addBlankSlide(d);p.setSlideBackground(s,C.bg);txts=[];expected.push(txts);return s}
function shape(name,x,y,w,h,fill,preset='rect',radius=0,shadow=false){const a=p.addSlideShape(s,{preset,x:U(x),y:U(y),w:U(w),h:U(h),name});p.setShapeFill(a,fill);p.setShapeNoStroke(a);if(radius)p.setShapeAdjustValues(a,{adj:Math.round(radius/Math.min(w,h)*100000)});if(shadow)p.setShapeShadow(a,{color:C.ink,opacity:.13,blurEmu:U(12),offsetEmu:U(5),angleDeg:90});return a}
function text(name,str,x,y,w,h,size,color=C.ink,bold=false,align='left'){const fix=adjust[name]||{};x+=fix.dx||0;y+=fix.dy||0;size=fix.size||size;const a=p.addSlideTextBox(s,{x:U(x),y:U(y),w:U(w),h:U(h),text:str,name});p.setShapeTextFormat(a,{font:'Noto Sans CJK TC',fontEastAsian:'Noto Sans CJK TC',size:size/1.5,color,bold});p.setShapeTextMargins(a,{left:0,top:0,right:0,bottom:0});p.setShapeTextAnchor(a,'top');p.setShapeTextWrap(a,'none');p.setShapeTextAutoFit(a,'none');p.setShapeAlignment(a,align);txts.push(str);return a}
function line(name,x1,y1,x2,y2,color=C.muted,width=3){const a=p.addSlideLine(s,{from:{x:U(x1),y:U(y1)},to:{x:U(x2),y:U(y2)},color,widthEmu:U(width),name});return a}
function poly(name,pts,color,width){return pts.slice(1).map((q,i)=>line(name+'-'+i,...pts[i],...q,color,width));}
function group(name,objects){return p.groupShapes(objects,{name})}
function head(title,sub){text('page-title',title,64,89,1280,69,44,C.ink,true);text('page-subtitle',sub,64,163,1312,48,23,C.muted)}
// Page 1: independently transcribed from visible raster.
slide();text('test-material-label','SLIDERECAST / ORIGINAL TEST MATERIAL',64,39,1000,30,18,C.muted,true);
head('Rebuild the visual. Keep the meaning.','Fidelity-first reconstruction / 高保真視覺重建');
const card1=shape('step-01-gradient-card',64,250,420,425,C.blue,'roundRect',12,true);p.setShapeGradientFill(card1,{angleDeg:0,stops:[{offset:0,color:'#245CEC'},{offset:1,color:'#553BC8'}]});
shape('step-02-card',525,250,399,425,C.white,'roundRect',12,true);shape('step-03-card',963,250,413,425,C.white,'roundRect',12,true);
text('step-01-number','01',94,280,200,70,48,C.white,true);text('step-01-title','See the original',94,372,370,60,34,C.white,true);text('step-01-details','版面、字型、色彩',94,454,340,48,24,C.white);
group('original-inspection-symbol',[shape('inspection-ring',345,285,90,90,C.white,'ellipse'),shape('inspection-center',365,305,50,50,C.blue,'ellipse'),shape('inspection-center-divider',384,307,12,46,C.white)]);
text('step-02-number','02',557,280,200,70,48,C.green,true);text('step-02-title','Recreate layers',557,373,355,60,32,C.ink,true);text('step-02-english','Native text and shapes',557,454,355,45,22,C.muted);text('step-02-chinese','主要元素可個別編輯',557,509,355,45,22,C.muted);
group('layer-stack',[shape('layer-back',825,299,60,40,'#DCE6FF','roundRect',12),shape('layer-middle',813,313,60,40,'#7FA5F8','roundRect',12),shape('layer-front',801,327,60,40,C.blue,'roundRect',12)]);
text('step-03-number','03',995,280,200,70,48,C.blue,true);text('step-03-title','Compare & correct',995,373,367,60,32,C.ink,true);text('step-03-details','Render before delivery',995,454,367,45,22,C.muted);
group('correct-arrow',poly('correct-arrow-path',[[1002,575],[1099,575],[1080,557]],C.green,5));
text('source-footer','Source designed independently for testing. No client or company documents.',64,734,1300,45,18,C.muted);
// Page 2: native card diagram and native icons.
slide();text('test-material-label','SLIDERECAST / ORIGINAL TEST MATERIAL',64,39,1000,30,18,C.muted,true);head('From pixels to editable layers','Complex screenshot fixture / 原始外觀優先');
for(const [i,x] of [90,435,780,1125].entries())shape('process-card-'+i,x,320,235,205,C.white,'roundRect',12,true);
for(const [i,x] of [325,670,1015].entries())group('forward-connector-'+i,poly('forward-connector-'+i,[[x,420],[x+99,420],[x+89,414]],C.muted,3).concat(line('arrow-lower-'+i,x+99,420,x+89,426,C.muted,3)));
const icons=[];
icons.push(shape('inspect-disc',172,345,70,70,C.blue,'ellipse'));icons.push(shape('magnifier-white',189,362,30,30,C.white,'ellipse'));icons.push(line('magnifier-handle',213,387,229,400,C.white,5));group('inspect-icon',icons);
group('rebuild-icon',[shape('rebuild-disc',517,345,70,70,C.green,'ellipse'),shape('rebuild-layer-back',538,361,30,20,'#A8E5DC'),shape('rebuild-layer-middle',531,370,30,20,'#A8E5DC'),shape('rebuild-layer-front',524,379,30,20,C.white)]);
group('render-icon',[shape('render-disc',862,345,70,70,C.purple,'ellipse'),shape('monitor-screen',877,366,40,25,C.white),line('monitor-stand',897,392,897,401,C.white,3)]);
group('correct-icon',[shape('correct-disc',1207,345,70,70,C.orange,'ellipse'),...poly('check-mark',[[1227,377],[1239,389],[1261,368]],C.white,4)]);
['Inspect','Rebuild','Render','Correct'].forEach((label,i)=>text('stage-'+label,label,[105,450,795,1140][i],438,220,50,28,C.ink,true));
group('targeted-correction-loop',[line('loop-right',1243,540,1243,620,C.muted,2),line('loop-bottom',1243,620,550,620,C.muted,2),line('loop-left',550,620,550,547,C.muted,2),line('loop-head-upper',550,547,541,541,C.muted,2),line('loop-head-lower',550,547,541,553,C.muted,2)]);
text('targeted-loop-label','Targeted correction loop',636,647,600,45,22,C.muted);text('preserve-wording-note','文字保持原樣',90,728,650,44,20,C.ink);
// Page 3: editable shape-based chart, without an inferred workbook.
slide();text('test-material-label','SLIDERECAST / ORIGINAL TEST MATERIAL',64,39,1000,30,18,C.muted,true);head('A chart is also a visual design','Default: editable shapes and labels. Native chart data only on request.');
shape('chart-panel',70,255,920,450,C.white,'roundRect',12,true);text('chart-heading','Observed sample values',110,280,800,52,30,C.ink,true);
[390,468,546,624].forEach((y,i)=>line('chart-gridline-'+i,120,y,930,y,'#E6EDF5',1.5));
const vals=[32,58,76,92],xs=[177,353,529,705],tops=[554,476,422,374],cs=[C.blue,C.green,C.purple,C.orange];
for(let i=0;i<4;i++){shape('bar-Q'+(i+1),xs[i],tops[i],82,650-tops[i],cs[i],'roundRect',12);text('value-Q'+(i+1),''+vals[i],xs[i]-10,tops[i]-44,120,45,26,C.ink,true);text('quarter-Q'+(i+1),'Q'+(i+1),xs[i]-10,657,100,45,22,C.muted)}
text('chart-note-title','Exact labels',1050,345,350,55,32,C.ink,true);text('chart-exact-values','32 / 58 / 76 / 92',1050,421,360,50,26,C.blue,true);text('chart-no-inference','No inferred data',1050,502,360,50,23,C.muted);text('chart-no-style-changes','No style changes',1050,557,360,50,23,C.muted);
await fs.writeFile(OUT+'/editable.pptx',await p.savePresentation(d));await fs.writeFile(OUT+'/generator-text-inventory.json',JSON.stringify(expected,null,2));
console.log(JSON.stringify({slides:p.getSlideCount(d),nativeImages:p.getAllImages(d).length,output:OUT+'/editable.pptx'}));
