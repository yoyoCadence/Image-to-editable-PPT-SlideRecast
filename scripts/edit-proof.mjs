// Programmatic native edits with read-back. Does not substitute for an application UI test.
import fs from 'node:fs/promises';
import crypto from 'node:crypto';
import * as p from '@office-kit/pptx';
const hash=b=>crypto.createHash('sha256').update(b).digest('hex');const report=[];
for(const engine of ['office-kit','pptxgenjs']) {
 const input=`examples/results/${engine}.pptx`;const bytes=await fs.readFile(input);const deck=await p.loadPresentation(bytes);
 const slide=p.getSlides(deck)[0];const title=p.getSlideShapes(slide).find(s=>p.getShapeText(s)==='Rebuild the visual. Keep the meaning.');
 if(!title)throw new Error('Native text not found');
 p.setShapeText(title,'EDIT PROOF / 原生文字已修改');
 const card=p.getSlideShapes(slide).find(s=>p.getShapeName(s).startsWith('infographic-shape') && p.getShapeBounds(s)?.w<p.inches(5));
 if(!card)throw new Error('Native card not found');p.setShapeFill(card,'#0F8A75');
 const b=p.getShapeBounds(card);p.setShapePosition(card,b.x+p.inches(.08),b.y);
 const output=`examples/results/${engine}-edited.pptx`;await fs.writeFile(output,await p.savePresentation(deck));
 const reloaded=await p.loadPresentation(await fs.readFile(output));
 if(!p.getSlideShapes(p.getSlides(reloaded)[0]).some(s=>p.getShapeText(s)==='EDIT PROOF / 原生文字已修改'))throw new Error('Text edit read-back failed');
 const editedCard=p.getSlideShapes(p.getSlides(reloaded)[0]).find(s=>p.getShapeName(s)===p.getShapeName(card));
 if(p.getShapeFillColor(editedCard)!=='#0F8A75')throw new Error('Fill edit read-back failed');
 if(hash(await fs.readFile(input))!==hash(bytes))throw new Error('Original modified');
 const nativeInput=`examples/results/${engine}-native-chart.pptx`;const originalChart=await fs.readFile(nativeInput);const chartDeck=await p.loadPresentation(originalChart);
 const chart=p.getSlideCharts(p.getSlides(chartDeck)[0])[0];if(!chart)throw new Error('Native chart not found');
 p.setChartSpec(chart,{kind:'column',categories:['Q1','Q2','Q3','Q4'],series:[{name:'Sample',values:[33,59,77,93]}]});
 const chartOutput=`examples/results/${engine}-chart-edited.pptx`;await fs.writeFile(chartOutput,await p.savePresentation(chartDeck));
 const chartReload=await p.loadPresentation(await fs.readFile(chartOutput));const values=p.getShapeChartSeriesValues(p.getSlideShapes(p.getSlides(chartReload)[0]).find(s=>p.isChartShape(s)),'Sample');
 if(JSON.stringify(values)!=='[33,59,77,93]')throw new Error('Chart edit read-back failed');
 if(hash(await fs.readFile(nativeInput))!==hash(originalChart))throw new Error('Original chart modified');
 report.push({engine,text_changed:true,shape_fill_changed:true,shape_position_changed:true,chart_values_read_back:values,original_sha256:hash(bytes),original_retained:true,
              method:'Office Kit load/edit/save/reload on native objects. Embedded XLSX validated separately.',ui_selection_verified:false,microsoft_verified:false});
}
await fs.writeFile('examples/results/edit-proof.json',JSON.stringify(report,null,2)+'\n');console.log(report);
