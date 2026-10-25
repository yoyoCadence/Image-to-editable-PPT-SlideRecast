#!/usr/bin/env python3
"""Preserve PDF page order, extract available text, rasterize for agent inspection."""
import argparse
import json
from pathlib import Path

def prepare(source, output, width=1440):
    from pypdf import PdfReader
    from raster_pdf import raster_pdf
    output=Path(output);output.mkdir(parents=True,exist_ok=True);pages=[]
    doc=PdfReader(source)
    images=raster_pdf(source,output,width)
    if len(images)!=len(doc.pages):raise RuntimeError('PDF parser/raster page counts disagree')
    for i,page in enumerate(doc.pages):
        text=f'page-{i+1:03}.txt';content=page.extract_text() or ''
        (output/text).write_text(content,encoding='utf-8')
        pages.append({'page':i+1,'width_pt':float(page.mediabox.width),'height_pt':float(page.mediabox.height),
            'image':images[i],'text':text,'has_extractable_text':bool(content.strip()),'rotation_deg':page.rotation})
    (output/'pages.json').write_text(json.dumps(pages,indent=2)+'\n');return pages
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');a=p.parse_args();print(json.dumps(prepare(a.source,a.output),indent=2))
