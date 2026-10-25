import fs from 'node:fs/promises';
import * as pptx from '@office-kit/pptx';
import {renderSlideToImage,buildFontkitMeasurer,auditTextLayout} from '@office-kit/pptx-preview/node';
import {performance} from 'node:perf_hooks';
const [source,out,fontFile]=process.argv.slice(2);
if(!source||!out)throw new Error('Usage: node scripts/preview.mjs input.pptx output-dir [CJK-font-file]');
const start=performance.now();const deck=await pptx.loadPresentation(await fs.readFile(source));
await fs.mkdir(out,{recursive:true});
const measureText=buildFontkitMeasurer(fontFile?{fonts:[{family:'Noto Sans CJK TC',source:fontFile}]}:{});
const images=[];
for(const [i,slide] of pptx.getSlides(deck).entries()){
 const file=`slide-${String(i+1).padStart(3,'0')}.png`;
 await fs.writeFile(`${out}/${file}`,renderSlideToImage(deck,slide,{width:1440,measureText,fontFiles:fontFile?[fontFile]:[],loadSystemFonts:false}));images.push(file);
}
const report={renderer:'@office-kit/pptx-preview 0.15.0',seconds:(performance.now()-start)/1000,fontFile:fontFile||null,images,
 layout_audit:auditTextLayout(deck,{measureText,reportSoftWraps:true}),package_validation:pptx.validatePresentation(deck),
 scope:'Approximate preview and font-dependent layout heuristic; not Microsoft validation.'};
await fs.writeFile(`${out}/preview.json`,JSON.stringify(report,null,2)+'\n');console.log(JSON.stringify(report,null,2));
