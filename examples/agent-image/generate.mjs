// Independent visual recreation from complex-flow.png only.
// Coordinates measured against the 1440 x 810 source raster.
import {createRequire} from 'node:module';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const require=createRequire(import.meta.url);
const PptxGenJS=require(fileURLToPath(new URL('../../skills/image-to-editable-ppt/vendor/pptxgenjs.cjs',import.meta.url)));
const here=path.dirname(fileURLToPath(import.meta.url));
const p=new PptxGenJS();
p.defineLayout({name:'SOURCE',width:12,height:6.75});p.layout='SOURCE';
p.author='SlideRecast independent conversion';p.subject='Editable recreation of the permitted source PNG';
p.title='From pixels to editable layers';p.company='';p.lang='zh-TW';
const s=p.addSlide();s.background={color:'F5F8FC'};
const S=p.ShapeType;
const u=n=>n/120;
const palette={ink:'16324F',muted:'587088',blue:'245CEC',teal:'00A68C',purple:'7953CF',orange:'D67A24'};
function shape(type,x,y,w,h,fill,name,extra={}){
 s.addShape(type,{x:u(x),y:u(y),w:u(w),h:u(h),fill:{color:fill},line:{color:fill,transparency:100},objectName:name,...extra});
}
function line(x1,y1,x2,y2,color,width,name){
 s.addShape(S.line,{x:u(Math.min(x1,x2)),y:u(Math.min(y1,y2)),w:u(Math.abs(x2-x1)),h:u(Math.abs(y2-y1)),flipH:x1>x2,flipV:y1>y2,line:{color,width:width*.6},objectName:name});
}
function pathShape(points,color,width,name){
 const xs=points.map(a=>a[0]),ys=points.map(a=>a[1]);const x=Math.min(...xs),y=Math.min(...ys),w=Math.max(...xs)-x,h=Math.max(...ys)-y;
 s.addShape(S.custGeom,{x:u(x),y:u(y),w:u(w),h:u(h),points:points.map(a=>({x:u(a[0]-x),y:u(a[1]-y)})),fill:{type:'none'},line:{color,width:width*.6},objectName:name});
}
function text(t,x,y,w,h,size,bold,color,name){
 s.addText(t,{x:u(x),y:u(y),w:u(w),h:u(h),fontFace:'Noto Sans CJK TC',fontSize:size*.6,bold,color,margin:0,breakLine:false,valign:'top',align:'left',lang:'zh-TW',paraSpaceAfter:0,objectName:name});
}
text('SLIDERECAST / ORIGINAL TEST MATERIAL',64,43,680,36,17,true,palette.muted,'eyebrow');
text('From pixels to editable layers',65,95,850,75,42,true,palette.ink,'headline');
text('Complex screenshot fixture / 原始外觀優先',64,168,700,46,21.7,false,palette.muted,'subtitle');

for(const [i,x] of [90,435,780,1125].entries()){
 shape(S.roundRect,x,320,235,205,'FFFFFF',`card-${i+1}`,{rectRadius:u(12),shadow:{type:'outer',color:'16324F',opacity:.13,blur:9.6,offset:3.6,angle:90,rotateWithShape:false}});
}
for(const [i,x] of [325,670,1015].entries()){
 line(x,420,x+98,420,palette.muted,3,`forward-connector-${i+1}`);
 pathShape([[x+89,414],[x+98,420],[x+89,426]],palette.muted,3,`forward-arrowhead-${i+1}`);
}
const centers=[207,552,897,1242],fills=[palette.blue,palette.teal,palette.purple,palette.orange];
centers.forEach((x,i)=>shape(S.ellipse,x-35,345,70,70,fills[i],`icon-disc-${i+1}`));
// Magnifying glass: editable circle and handle.
shape(S.ellipse,189,362,30,30,'FFFFFF','inspect-lens');
line(214,387,229,400,'FFFFFF',4.5,'inspect-handle');
// Rebuild: three independently editable overlapping tiles.
shape(S.rect,538,361,30,20,'A8E5DC','rebuild-tile-back');
shape(S.rect,531,370,30,20,'A8E5DC','rebuild-tile-middle');
shape(S.rect,524,379,30,20,'FFFFFF','rebuild-tile-front');
// Monitor: editable screen and stand.
shape(S.rect,877,366,40,25,'FFFFFF','render-monitor');
shape(S.rect,895.5,392,3,9,'FFFFFF','render-monitor-stand');
// Correct: native open polyline check.
pathShape([[1227,377],[1239,389],[1261,368]],'FFFFFF',4,'correct-check');
for(const [i,t] of ['Inspect','Rebuild','Render','Correct'].entries()){
 text(t,[106,451,796,1140][i],444,203,56,26,true,palette.ink,`step-label-${i+1}`);
}
pathShape([[1243,540],[1243,620],[550,620],[550,547]],palette.muted,2,'return-loop');
pathShape([[541,541],[550,547],[541,553]],palette.muted,2,'return-loop-arrowhead');
text('Targeted correction loop',636,652,450,46,22,false,palette.muted,'loop-caption');
text('文字保持原樣',90,734,280,44,20,false,palette.ink,'footer-note');
await p.writeFile({fileName:path.join(here,'editable-flow.pptx')});
