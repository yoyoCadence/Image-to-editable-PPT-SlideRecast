#!/usr/bin/env python3
"""Rasterize through an external Poppler process; no PDF native library bundled."""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

def find_pdftoppm():
    candidates=[os.environ.get('SLIDERECAST_PDFTOPPM'),shutil.which('pdftoppm')]
    if os.environ.get('CODEX_PRIMARY_RUNTIME_ROOT'):
        candidates.append(str(Path(os.environ['CODEX_PRIMARY_RUNTIME_ROOT'])/'dependencies/bin/override/pdftoppm'))
    return next((p for p in candidates if p and Path(p).is_file()),None)

def raster_pdf(source,output,width=1440,prefix='page'):
    tool=find_pdftoppm()
    if not tool:raise RuntimeError('Poppler pdftoppm unavailable. Use an existing host PDF rasterizer and record that fallback.')
    output=Path(output);output.mkdir(parents=True,exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='sliderecast-pdf-') as folder:
        temporary=Path(folder)/'raster'
        result=subprocess.run([tool,'-png','-scale-to-x',str(width),'-scale-to-y','-1',str(Path(source).resolve()),str(temporary)],capture_output=True,text=True,timeout=120)
        if result.returncode:raise RuntimeError(f'Poppler rasterization failed: {result.stderr}')
        pages=sorted(Path(folder).glob('raster-*.png'),key=lambda p:int(p.stem.rsplit('-',1)[1]))
        if not pages:raise RuntimeError('PDF rasterizer produced no pages')
        names=[]
        for i,page in enumerate(pages):
            name=f'{prefix}-{i+1:03}.png';shutil.copyfile(page,output/name);names.append(name)
    return names
