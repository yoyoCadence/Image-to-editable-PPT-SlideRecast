#!/usr/bin/env python3
"""Pixel evidence and ranked difference regions, not a fidelity percentage."""
import argparse
import json
from pathlib import Path

def compare(source, rendered, output, grid=6):
    from PIL import Image, ImageChops, ImageEnhance, ImageDraw
    a=Image.open(source).convert('RGB');b=Image.open(rendered).convert('RGB')
    if abs(a.width/a.height-b.width/b.height)>.005: raise ValueError('Aspect ratio differs; do not stretch away layout errors')
    original_size=b.size;b=b.resize(a.size,Image.Resampling.LANCZOS)
    diff=ImageChops.difference(a,b);hist=diff.histogram()
    mae=sum((i%256)*count for i,count in enumerate(hist))/(a.width*a.height*3*255)
    pix=diff.load();changed=sum(max(pix[x,y])>24 for y in range(a.height) for x in range(a.width))/(a.width*a.height)
    regions=[]
    for iy in range(grid):
        for ix in range(grid):
            box=(ix*a.width//grid,iy*a.height//grid,(ix+1)*a.width//grid,(iy+1)*a.height//grid)
            crop=diff.crop(box);h=crop.histogram();m=sum((i%256)*n for i,n in enumerate(h))/(crop.width*crop.height*3*255)
            regions.append({'box_px':box,'normalized_mae':m})
    regions.sort(key=lambda r:r['normalized_mae'],reverse=True)
    output=Path(output);output.mkdir(parents=True,exist_ok=True)
    overlay=Image.blend(a,b,.5);ImageEnhance.Brightness(diff).enhance(3).save(output/'difference.png');overlay.save(output/'overlay.png')
    pair=Image.new('RGB',(a.width*2,a.height+36),'#e8edf4');pair.paste(a,(0,36));pair.paste(b,(a.width,36));draw=ImageDraw.Draw(pair);draw.text((16,10),'SOURCE',fill='black');draw.text((a.width+16,10),'RENDER',fill='black');pair.save(output/'comparison.png')
    report={'source_size':a.size,'render_size_before_resampling':original_size,'normalized_rgb_mae':mae,'pixels_with_max_channel_delta_gt_24_fraction':changed,'threshold':24,
        'method':'RGB MAE divided by 255, rendered resampled to source size, no alignment correction; difference may include font/antialiasing/renderer effects',
        'ranked_regions':regions,'semantic_or_text_correctness':'not measured by pixels'}
    (output/'metrics.json').write_text(json.dumps(report,indent=2)+'\n');return report
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source');p.add_argument('rendered');p.add_argument('output');a=p.parse_args();print(json.dumps(compare(a.source,a.rendered,a.output),indent=2))
