#!/usr/bin/env python3
"""Known-layout synthetic benchmark. All measurements retain denominators."""
import json
import sys
import hashlib
from pathlib import Path
import subprocess
import time
import os
import shutil

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'skills/image-to-editable-ppt/scripts'))
from inspect_pptx import inspect
from compare import compare
from render import render
from prepare_pdf import prepare

def main():
    import fitz
    results=ROOT/'examples/results';source=ROOT/'examples/source'
    names=['infographic','complex-flow','visual-chart'];expected=json.loads((source/'expected-text.json').read_text())
    font=os.environ.get('SLIDERECAST_TEST_FONT')
    if not font and shutil.which('fc-match'):
        font=subprocess.run(['fc-match','-f','%{file}','Noto Sans CJK TC'],capture_output=True,text=True,check=True).stdout
    if not font or not Path(font).is_file():raise RuntimeError('Set SLIDERECAST_TEST_FONT to an installed Noto Sans CJK TC font file for this developer benchmark.')
    # Independently authored image-only PDF; it tests page handling, not PDF selectable-text extraction.
    with fitz.open() as doc:
        for name in names:
            page=doc.new_page(width=960,height=540);page.insert_image(page.rect,filename=str(source/f'{name}.png'))
        doc.save(source/'multipage.pdf',deflate=True)
    # Also test selectable text in a separate PDF with deterministic text ordering.
    with fitz.open() as doc:
        for i,label in enumerate(['First page','Second page']):
            page=doc.new_page();page.insert_text((72,72),label,fontsize=24)
        doc.save(source/'selectable-text.pdf')
    pdf_pages=prepare(source/'multipage.pdf',results/'pdf-pages')
    selectable=prepare(source/'selectable-text.pdf',results/'selectable-pages')
    report={'date':'2026-10-13','design':'3 independently authored synthetic known-layout fixtures, two engines. Shared coordinates intentionally measure authoring/rendering capability, not image understanding or agent generalization.',
            'sample_slides_per_engine':3,'engines':{},'pdf':{'raster_pages':len(pdf_pages),'selectable_pages':len(selectable),'text_extracted':[p['has_extractable_text'] for p in selectable]},
            'claude_ai':'BLOCKED: browser reached sign-in page; not executed','microsoft_powerpoint':'NOT EXECUTED',
            'manual_repair_minutes':None,'claude_tokens_or_cost':None}
    for engine in ['office-kit','pptxgenjs']:
        file=results/f'{engine}.pptx';inventory=inspect(file,expected);(results/f'{engine}-structure.json').write_text(json.dumps(inventory,ensure_ascii=False,indent=2)+'\n')
        rendering=render(file,results/f'{engine}-render')
        preview_dir=results/f'{engine}-preview'
        subprocess.run(['node',str(ROOT/'scripts/preview.mjs'),str(file),str(preview_dir),font],cwd=ROOT,check=True,stdout=subprocess.DEVNULL)
        comparisons=[]
        for i,name in enumerate(names):
            metrics=compare(source/f'{name}.png',results/f'{engine}-render/slide-{i+1:03}.png',results/f'{engine}-comparison-{i+1}')
            preview_delta=compare(results/f'{engine}-render/slide-{i+1:03}.png',preview_dir/f'slide-{i+1:03}.png',results/f'{engine}-renderer-delta-{i+1}')
            comparisons.append({'fixture':name,'rgb_mae':metrics['normalized_rgb_mae'],'changed_fraction':metrics['pixels_with_max_channel_delta_gt_24_fraction'],
                                'preview_vs_libreoffice_rgb_mae':preview_delta['normalized_rgb_mae']})
        native=inspect(results/f'{engine}-native-chart.pptx');(results/f'{engine}-chart-structure.json').write_text(json.dumps(native,indent=2)+'\n')
        render(results/f'{engine}-native-chart.pptx',results/f'{engine}-chart-render')
        preview=json.loads((preview_dir/'preview.json').read_text())
        report['engines'][engine]={'package_checks':inventory['valid_package_checks'],'text_object_exact_matches':sum(x['exact'] for x in inventory['expected_text']),
            'text_object_total':sum(len(p) for p in expected),'text_slide_denominator':3,'inventory':[{k:p[k] for k in ['native_text_objects','native_shapes','connectors','pictures','gradient_fills','shadows']} for p in inventory['slides']],
            'rendering':rendering,'comparison':comparisons,'preview_layout_audit':preview['layout_audit'],'office_kit_package_validation':preview['package_validation'],
            'optional_chart':native['charts'],'pptx_bytes':file.stat().st_size}
    (results/'benchmark.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    print(json.dumps({k:{'package_checks':v['package_checks'],'text_slides':v['text_object_exact_matches'],'comparisons':v['comparison']} for k,v in report['engines'].items()},indent=2))
if __name__=='__main__': main()
