#!/usr/bin/env python3
"""Read-only PPTX package checks. Not an ECMA-376 schema validator."""
import argparse
import collections
import hashlib
import json
import posixpath
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}
R = '{http://schemas.openxmlformats.org/package/2006/relationships}'
def norm(text):
    return ' '.join(text.split())

def inspect(path, expected=None):
    errors, warnings, slides, charts = [], [], [], []
    with ZipFile(path) as z:
        if z.testzip():
            errors.append('ZIP CRC failure')
        names = set(z.namelist())
        if len(names) != len(z.namelist()):
            errors.append('Duplicate ZIP entries')
        roots = {}
        for name in names:
            if name.endswith(('.xml', '.rels')):
                try:
                    roots[name] = ET.fromstring(z.read(name))
                except ET.ParseError as exc:
                    errors.append(f'{name}: {exc}')
        for part in ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels']:
            if part not in roots:
                errors.append(f'Missing required XML part: {part}')
        rel_maps = {}
        for name, root in roots.items():
            if not name.endswith('.rels'):
                continue
            base = posixpath.dirname(posixpath.dirname(name)) if name != '_rels/.rels' else ''
            rel_maps[name] = {}
            for rel in root:
                target = rel.get('Target', '')
                if rel.get('TargetMode') == 'External':
                    warnings.append(f'External relationship: {name} -> {target}')
                    continue
                target = posixpath.normpath(posixpath.join(base, target)) if not target.startswith('/') else target[1:]
                if target not in names:
                    errors.append(f'Dangling relationship: {name} -> {target}')
                if rel.get('Id') in rel_maps[name]:
                    errors.append(f'Duplicate relationship ID: {name} {rel.get("Id")}')
                rel_maps[name][rel.get('Id')] = target
        pres = roots.get('ppt/presentation.xml')
        if pres is not None:
            size = pres.find('p:sldSz', NS)
            width, height = (int(size.get('cx')), int(size.get('cy'))) if size is not None else (0, 0)
            if width <= 0 or height <= 0:
                errors.append('Invalid slide dimensions')
            for index, sid in enumerate(pres.findall('p:sldIdLst/p:sldId', NS)):
                part = rel_maps.get('ppt/_rels/presentation.xml.rels', {}).get(sid.get('{'+NS['r']+'}id'))
                root = roots.get(part)
                if root is None:
                    errors.append(f'Unresolved slide {index + 1}')
                    continue
                texts = []
                for shape in root.findall('.//p:sp', NS):
                    paragraphs = [''.join(t.text or '' for t in p.findall('.//a:t', NS)) for p in shape.findall('p:txBody/a:p', NS)]
                    text = '\n'.join(paragraphs)
                    if text:
                        texts.append(text)
                ids = [e.get('id') for e in root.findall('.//p:cNvPr', NS)]
                if len(ids) != len(set(ids)):
                    errors.append(f'Slide {index+1}: duplicate shape IDs')
                pictures = root.findall('.//p:pic', NS)
                image_bounds = []
                for pic in pictures:
                    ext = pic.find('p:spPr/a:xfrm/a:ext', NS)
                    if ext is not None:
                        image_bounds.append([int(ext.get('cx')), int(ext.get('cy'))])
                fullpage = any(w >= width*.95 and h >= height*.95 for w, h in image_bounds)
                if fullpage:
                    warnings.append(f'Slide {index+1}: large picture, inspect for flattening')
                outside = []
                for xfrm in root.findall('p:cSld/p:spTree/*/p:spPr/a:xfrm', NS):
                    off, ext = xfrm.find('a:off', NS), xfrm.find('a:ext', NS)
                    if off is not None and ext is not None:
                        x,y,w,h = [int(e.get(k)) for e,k in [(off,'x'),(off,'y'),(ext,'cx'),(ext,'cy')]]
                        if x < -1000 or y < -1000 or x+w > width+1000 or y+h > height+1000:
                            outside.append([x,y,w,h])
                if outside:
                    warnings.append(f'Slide {index+1}: {len(outside)} top-level objects outside canvas')
                slides.append({'index': index+1, 'part':part, 'texts':texts,
                    'native_text_objects':len(texts), 'native_shapes':len(root.findall('.//p:sp', NS)),
                    'connectors':len(root.findall('.//p:cxnSp', NS)), 'groups':len(root.findall('.//p:grpSp', NS)),
                    'pictures':len(pictures), 'large_picture':fullpage, 'outside_canvas':outside,
                    'gradient_fills':len(root.findall('.//a:gradFill', NS)),
                    'shadows':len(root.findall('.//a:outerShdw', NS))})
        for name in sorted(roots):
            if re.fullmatch(r'ppt/charts/chart\d+\.xml', name):
                root = roots[name]
                charts.append({'part':name, 'series':len(root.findall('.//c:ser',NS)),
                    'cached_values':[v.text for v in root.findall('.//c:val/c:numRef/c:numCache/c:pt/c:v',NS)],
                    'external_data_relationships':len(root.findall('c:externalData',NS))})
        expected_checks = []
        if expected is not None:
            if len(expected) != len(slides):
                errors.append(f'Expected {len(expected)} slides, found {len(slides)}')
            for i, wanted in enumerate(expected):
                actual = collections.Counter(norm(t) for t in slides[i]['texts']) if i < len(slides) else collections.Counter()
                desired = collections.Counter(norm(t) for t in wanted)
                missing, extra = list((desired-actual).elements()), list((actual-desired).elements())
                expected_checks.append({'slide':i+1,'missing':missing,'extra':extra,'exact':not missing and not extra})
                if missing or extra:
                    errors.append(f'Slide {i+1}: text mismatch')
    return {'file':Path(path).name, 'sha256':hashlib.sha256(Path(path).read_bytes()).hexdigest(),
        'scope':'ZIP/XML/relationships/order/object inventory only; not full OOXML schema or UI editability validation',
        'valid_package_checks':not errors, 'errors':errors, 'warnings':warnings,
        'slides':slides, 'charts':charts, 'expected_text':expected_checks}

def main():
    p=argparse.ArgumentParser();p.add_argument('pptx');p.add_argument('--expected');p.add_argument('--out')
    a=p.parse_args();report=inspect(a.pptx, json.loads(Path(a.expected).read_text()) if a.expected else None)
    out=json.dumps(report,ensure_ascii=False,indent=2)+'\n'
    if a.out: Path(a.out).write_text(out)
    else: print(out)
    raise SystemExit(0 if report['valid_package_checks'] else 1)
if __name__ == '__main__': main()
