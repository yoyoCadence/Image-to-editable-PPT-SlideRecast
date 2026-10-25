#!/usr/bin/env python3
"""Report available capabilities; never install anything."""
import importlib.util
import json
import os
import shutil
import subprocess
from pathlib import Path

def find_soffice():
    candidates = [os.environ.get('SLIDERECAST_SOFFICE'), shutil.which('libreoffice'), shutil.which('soffice')]
    if os.environ.get('CODEX_PRIMARY_RUNTIME_ROOT'):
        candidates.append(str(Path(os.environ['CODEX_PRIMARY_RUNTIME_ROOT'])/'dependencies/bin/override/soffice'))
    candidates += [r'C:\Program Files\LibreOffice\program\soffice.exe', '/Applications/LibreOffice.app/Contents/MacOS/soffice']
    return next((x for x in candidates if x and Path(x).is_file()), None)

def probe():
    node=shutil.which('node');version=None
    if node:
        version=subprocess.run([node,'--version'],capture_output=True,text=True,timeout=10).stdout.strip()
    return {'node':version,'office_kit_minimum_node':'22.18.0','soffice':find_soffice(),
        'python_modules':{m:importlib.util.find_spec(m) is not None for m in ['PIL','pypdf','pptx']},
        'pdftoppm':shutil.which('pdftoppm'),'fc_match':shutil.which('fc-match'),
        'cloud_test':'This probe does not establish Claude.ai compatibility.'}
if __name__=='__main__': print(json.dumps(probe(),indent=2))
