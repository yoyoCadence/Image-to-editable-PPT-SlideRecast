#!/usr/bin/env python3
"""Create separate private-test candidates; never upload or publish."""
import hashlib
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED, ZipInfo

ROOT=Path(__file__).resolve().parents[1]
SKILL=ROOT/'skills/image-to-editable-ppt'

def files(folder):
    return sorted(p for p in folder.rglob('*') if p.is_file() and '__pycache__' not in p.parts and not p.name.endswith('.pyc'))

def archive(output, entries):
    temporary=output.with_suffix(output.suffix+'.tmp')
    with ZipFile(temporary,'w',compression=ZIP_DEFLATED,compresslevel=9) as z:
        for path, name in sorted(entries,key=lambda x:x[1]):
            info=ZipInfo(name,date_time=(2026,10,13,0,0,0));info.compress_type=ZIP_DEFLATED;info.external_attr=0o100644<<16
            z.writestr(info,path.read_bytes())
    with ZipFile(temporary) as z:
        if z.testzip():raise RuntimeError('Archive CRC failed')
    temporary.replace(output)
    return {'file':output.name,'bytes':output.stat().st_size,'sha256':hashlib.sha256(output.read_bytes()).hexdigest()}

def main():
    vendor=SKILL/'vendor'
    for name in ['pptxgenjs.cjs','office-kit.mjs','THIRD-PARTY-NOTICES.txt','manifest.json']:
        if not (vendor/name).is_file():raise RuntimeError('Run npm run build first: missing '+name)
    manifest=json.loads((vendor/'manifest.json').read_text())
    for item in manifest['bundles']:
        if hashlib.sha256((vendor/item['file']).read_bytes()).hexdigest()!=item['sha256']:raise RuntimeError('Bundle hash mismatch')
    # Agent Skills uses uppercase SKILL.md. Claude's Help Center also spells skill.md
    # in its uploader guide, so use that canonical spelling only in the Claude ZIP.
    skill_entries=[]
    for path in files(SKILL):
        rel=path.relative_to(SKILL).as_posix()
        if rel.startswith('agents/'):continue
        if rel=='SKILL.md':rel='skill.md'
        skill_entries.append((path,'image-to-editable-ppt/'+rel))
    output=ROOT/'dist';output.mkdir(exist_ok=True)
    records=[archive(output/'image-to-editable-ppt-claude-alpha.zip',skill_entries)]
    plugin_entries=[(ROOT/'plugin.json','plugin.json'),(ROOT/'assets/logo.svg','assets/logo.svg')]
    plugin_entries.extend((p,'skills/image-to-editable-ppt/'+p.relative_to(SKILL).as_posix()) for p in files(SKILL))
    for name in ['THIRD-PARTY-NOTICES.md','docs/LICENSING.md']:
        if (ROOT/name).exists():plugin_entries.append((ROOT/name,name))
    records.append(archive(output/'sliderecast-openai-plugin-alpha.zip',plugin_entries))
    (output/'manifest.json').write_text(json.dumps({'status':'private-test candidates; not formally released, uploaded, submitted or platform-accepted','archives':records},indent=2)+'\n')
    print(json.dumps(records,indent=2))
if __name__=='__main__':main()
