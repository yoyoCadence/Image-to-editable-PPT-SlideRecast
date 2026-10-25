import { build } from 'esbuild';
import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';

const vendor='skills/image-to-editable-ppt/vendor';
await fs.rm(vendor,{recursive:true,force:true}); await fs.mkdir(vendor,{recursive:true});
const records=[]; const packages=new Map();
for (const [name,code,format] of [
  ['pptxgenjs.cjs',"module.exports = require('pptxgenjs');",'cjs'],
  ['office-kit.mjs',"export * from '@office-kit/pptx';",'esm'],
]) {
  const r=await build({stdin:{contents:code,resolveDir:process.cwd()},bundle:true,platform:'node',format,target:'node18',minify:true,legalComments:'eof',metafile:true,write:false});
  const bytes=r.outputFiles[0].contents;await fs.writeFile(path.join(vendor,name),bytes);
  records.push({file:name,bytes:bytes.length,sha256:crypto.createHash('sha256').update(bytes).digest('hex')});
  for (const input of Object.keys(r.metafile.inputs)) {
    if (!input.includes('node_modules/')) continue;
    let folder=path.dirname(path.resolve(input));
    while(folder!==path.dirname(folder)) {
      try {const p=JSON.parse(await fs.readFile(path.join(folder,'package.json'),'utf8'));if(p.name){packages.set(p.name,{folder,package:p});break;}}catch{}
      folder=path.dirname(folder);
    }
  }
}
let notice='# Bundled third-party notices\n\nGenerated from the actual bundle input packages. Project license remains pending.\n\n';
for(const [name,{folder,package:p}] of [...packages].sort()) {
  const files=(await fs.readdir(folder)).filter(f=>/^(license|licence|copying|third.party.notices)([._-]|$)/i.test(f));
  if (!files.length) {
    const readme=(await fs.readdir(folder)).find(f=>/^readme\.md$/i.test(f));
    const content=readme?await fs.readFile(path.join(folder,readme),'utf8'):'';
    const start=content.search(/^##?\s+License/im);
    if(start<0 || !content.slice(start).includes('Permission is hereby granted')) throw new Error(`Missing license text for ${name}`);
    notice+=`## ${name} ${p.version} (${p.license})\n\nLicense section from ${readme}:\n\n${content.slice(start)}\n\n`;
    continue;
  }
  notice+=`## ${name} ${p.version} (${p.license})\n\n`;
  for(const file of files)notice+=`### ${file}\n\n${await fs.readFile(path.join(folder,file),'utf8')}\n\n`;
  if(name==='pako') {
    const source=await fs.readFile(path.join(folder,'lib/zlib/deflate.js'),'utf8');
    notice+=`### zlib-derived source notice (unaltered copyright and terms)\n\n${source.slice(0,source.indexOf('var utils'))}\n\n`;
  }
  if(name==='pptxgenjs') {
    const source=await fs.readFile(path.join(folder,'dist/pptxgen.cjs.js'),'utf8');
    const start=source.indexOf('/********');const end=source.indexOf('*/',start);
    if(start>=0 && end>start)notice+=`### Upstream embedded Microsoft helper notice\n\n${source.slice(start,end+2)}\n\n`;
  }
}
await fs.writeFile(path.join(vendor,'THIRD-PARTY-NOTICES.txt'),notice);
const refs='skills/image-to-editable-ppt/references';
await fs.copyFile('node_modules/pptxgenjs/types/index.d.ts',path.join(refs,'pptxgenjs-api.d.ts'));
const officeFiles=await fs.readdir('node_modules/@office-kit/pptx/dist');
const officeTypes=officeFiles.find(f=>/^index-.+\.d\.ts$/.test(f));
await fs.copyFile(path.join('node_modules/@office-kit/pptx/dist',officeTypes),path.join(refs,'office-kit-api.d.ts'));
await fs.writeFile(path.join(vendor,'manifest.json'),JSON.stringify({generated_at:new Date().toISOString(),bundles:records,packages:[...packages].map(([name,{package:p}])=>({name,version:p.version,license:p.license})),office_kit_min_node:'22.18.0',notes:'Preview and external executables are not bundled. Bundling target node18 does not certify Office Kit below its declared engine.'},null,2)+'\n');
console.log(JSON.stringify(records,null,2));
