#!/usr/bin/env python3
"""Render copies through LibreOffice, retaining the original PPTX bytes."""
import argparse
import hashlib
import json
import subprocess
import tempfile
import shutil
import time
from pathlib import Path
from environment import find_soffice

def render(source, output, width=1440):
    from raster_pdf import raster_pdf
    source=Path(source).resolve(); output=Path(output).resolve();output.mkdir(parents=True,exist_ok=True)
    original_hash=hashlib.sha256(source.read_bytes()).hexdigest();start=time.perf_counter()
    soffice=find_soffice()
    if not soffice: raise RuntimeError('LibreOffice unavailable. Use an available renderer and report its name; do not claim visual verification.')
    with tempfile.TemporaryDirectory(prefix='sliderecast-render-') as folder:
        tmp=Path(folder);copy=tmp/'input.pptx';copy.write_bytes(source.read_bytes())
        profile=(tmp/'profile').as_uri()
        result=subprocess.run([soffice,f'-env:UserInstallation={profile}','--headless','--convert-to','pdf','--outdir',str(tmp),str(copy)],capture_output=True,text=True,timeout=120)
        pdf=tmp/'input.pdf'
        if result.returncode or not pdf.exists(): raise RuntimeError(f'LibreOffice conversion failed: {result.stderr} {result.stdout}')
        shutil.copyfile(pdf,output/'render.pdf')
        pages=raster_pdf(pdf,output,width,prefix='slide')
    if hashlib.sha256(source.read_bytes()).hexdigest()!=original_hash: raise RuntimeError('Original PPTX changed')
    version=subprocess.run([soffice,'--version'],capture_output=True,text=True,timeout=10).stdout.strip()
    from raster_pdf import find_pdftoppm
    rasterizer=subprocess.run([find_pdftoppm(),'-v'],capture_output=True,text=True,timeout=10)
    rasterizer_version=(rasterizer.stdout or rasterizer.stderr).splitlines()[0]
    report={'renderer':version,'pdf_rasterizer':rasterizer_version,'seconds':time.perf_counter()-start,'pages':pages,'source_sha256':original_hash,'original_retained':True,'microsoft_powerpoint_verified':False}
    (output/'render.json').write_text(json.dumps(report,indent=2)+'\n');return report

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('output');p.add_argument('--width',type=int,default=1440);a=p.parse_args();print(json.dumps(render(a.source,a.output,a.width),indent=2))
