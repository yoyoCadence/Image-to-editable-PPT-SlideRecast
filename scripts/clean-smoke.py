#!/usr/bin/env python3
"""Load extracted ZIP bundles from a temporary folder, without npm installs."""
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'skills/image-to-editable-ppt/scripts'))
from inspect_pptx import inspect
from render import render

def main():
    results=[]
    for archive in ['image-to-editable-ppt-claude-alpha.zip','sliderecast-openai-plugin-alpha.zip']:
        start=time.perf_counter()
        with tempfile.TemporaryDirectory(prefix='sliderecast-clean-') as folder:
            temp=Path(folder)
            with ZipFile(ROOT/'dist'/archive) as z:
                for name in z.namelist():
                    if Path(name).is_absolute() or '..' in Path(name).parts:raise RuntimeError('Unsafe archive path')
                z.extractall(temp)
            skill=temp/'image-to-editable-ppt' if archive.startswith('image-to-') else temp/'skills/image-to-editable-ppt'
            skill_md=skill/('skill.md' if archive.startswith('image-to-') else 'SKILL.md')
            if not skill_md.exists():raise RuntimeError('Missing platform-specific skill entry')
            vendor=skill/'vendor';manifest=json.loads((vendor/'manifest.json').read_text())
            for bundle in manifest['bundles']:
                if hashlib.sha256((vendor/bundle['file']).read_bytes()).hexdigest()!=bundle['sha256']:raise RuntimeError('Extracted bundle hash mismatch')
            env=dict(os.environ);env.pop('NODE_PATH',None)
            entry=temp/'smoke.mjs'
            entry.write_text("import fs from 'node:fs/promises';\nimport {createRequire} from 'node:module';\n"+
                'import * as p from '+json.dumps((vendor/'office-kit.mjs').as_uri())+';\n'+
                'const require=createRequire(import.meta.url);const PptxGenJS=require('+json.dumps(str(vendor/'pptxgenjs.cjs'))+');\n'+
                "const deck=p.createPresentation();const s=p.addBlankSlide(deck);p.addSlideTextBox(s,{x:p.inches(1),y:p.inches(1),w:p.inches(8),h:p.inches(1),text:'Clean bundle Office Kit'});const r=p.addSlideShape(s,{preset:'rect',x:p.inches(1),y:p.inches(2),w:p.inches(4),h:p.inches(2)});p.setShapeGradientFill(r,{angleDeg:0,stops:[{offset:0,color:'#245CEC'},{offset:1,color:'#5639C7'}]});await fs.writeFile('office.pptx',await p.savePresentation(deck));\n"+
                "const other=new PptxGenJS();const sl=other.addSlide();sl.addText('Clean bundle PptxGenJS',{x:1,y:1,w:8,h:1,fontSize:28});sl.addShape(other.ShapeType.rect,{x:1,y:2,w:4,h:2,fill:{color:'245CEC'}});await other.writeFile({fileName:'gen.pptx'});\n")
            subprocess.run(['node',str(entry)],cwd=temp,env=env,check=True,timeout=30,capture_output=True)
            inventories=[]
            for name in ['office','gen']:
                report=inspect(temp/f'{name}.pptx')
                if not report['valid_package_checks'] or report['slides'][0]['native_text_objects']!=1:raise RuntimeError('Clean output invalid')
                rendered=render(temp/f'{name}.pptx',temp/f'{name}-render',width=720)
                inventories.append({'engine':name,'package_checks':True,'native_text_objects':1,'render_pages':len(rendered['pages'])})
            results.append({'archive':archive,'seconds':time.perf_counter()-start,'npm_install_invoked':False,'project_node_modules_used':False,'bundled_engines':inventories,'host_node':subprocess.run(['node','--version'],capture_output=True,text=True,check=True).stdout.strip(),'claude_ai_test':False})
    (ROOT/'examples/results/clean-smoke.json').write_text(json.dumps(results,indent=2)+'\n');print(json.dumps(results,indent=2))
if __name__=='__main__':main()
