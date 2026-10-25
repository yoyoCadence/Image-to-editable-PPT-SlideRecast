# Tested authoring tools

Pinned at 2026-10-13: `@office-kit/pptx` 0.24.1, `pptxgenjs` 4.0.1. Bundles include their dependencies and notices. Preview 0.15.0 is a developer test dependency, **not** part of the skill bundle, because its Node path depends on native resvg and font files.

## Office Kit

Use Node >=22.18.0. The exported bundle is the core ESM entry, not `/node`; use standard `node:fs/promises` yourself. Units matter: geometry uses EMU (English Metric Unit), font sizes use points. The same API can author, read, and modify PPTX. Pin this 0.x release; minor updates may break code.

```js
import fs from 'node:fs/promises';
import * as pptx from '/resolved/skill-dir/vendor/office-kit.mjs';
const deck = pptx.createPresentation();
pptx.setSlideSize(deck, {width:pptx.inches(12),height:pptx.inches(7)});
const slide = pptx.addBlankSlide(deck);
const card = pptx.addSlideShape(slide, {
  preset:'roundRect', x:pptx.inches(1), y:pptx.inches(1),
  w:pptx.inches(4), h:pptx.inches(2), name:'source-panel'
});
pptx.clearShapeStroke(card);
pptx.setShapeGradientFill(card, {
  angleDeg:0, stops:[{offset:0,color:'#245CEC'},{offset:1,color:'#5639C7'}]
});
pptx.setShapeShadow(card,{color:'#16324F',opacity:.13,
  blurEmu:pptx.pt(4),offsetEmu:pptx.pt(3),angleDeg:90});
const title = pptx.addSlideTextBox(slide,{
  x:pptx.inches(1.2),y:pptx.inches(1.2),w:pptx.inches(3.6),
  h:pptx.inches(.6),text:'保留原始文字',name:'source-title'
});
pptx.setShapeTextFormat(title,{font:'Noto Sans CJK TC',
  fontEastAsian:'Noto Sans CJK TC',size:24,color:'#FFFFFF'});
pptx.setShapeTextMargins(title,{left:0,right:0,top:0,bottom:0});
pptx.setShapeTextAnchor(title,'top'); // setter uses top/center/bottom, not t/ctr/b
await fs.writeFile('/work/output.pptx',await pptx.savePresentation(deck));
```

For edits: `loadPresentation(bytes)`, `getSlides(deck)`, `getSlideShapes(slide)`, `setShapeText(shape,text)`, `setShapePosition(shape,x,y)`, then `savePresentation(deck)`. Preserve source bytes. `setShapeAdjustValues(shape,{adj:...})` tunes rounded-corner geometry. `addSlideLine()` authors lines with `from` and `to` EMU coordinates. `setShapeStrokeArrow()` uses `'head'`/`'tail'`; inspect the rendered arrow direction. For other geometry APIs read the publicly licensed `.d.ts` excerpts in `office-kit-api.d.ts`, or public documentation. Do not guess option names.

## PptxGenJS

```js
import {createRequire} from 'node:module';
const require=createRequire(import.meta.url);
const PptxGenJS=require('/resolved/skill-dir/vendor/pptxgenjs.cjs');
const deck=new PptxGenJS();
deck.defineLayout({name:'SOURCE',width:12,height:7});deck.layout='SOURCE';
const slide=deck.addSlide();
slide.addText('保留原始文字',{x:1,y:1,w:8,h:.7,
  fontFace:'Noto Sans CJK TC',fontSize:24,margin:0,valign:'top',objectName:'source-title'});
slide.addShape(deck.ShapeType.roundRect,{x:1,y:2,w:4,h:2,
  fill:{color:'245CEC'},line:{color:'245CEC',transparency:100},
  shadow:{type:'outer',color:'16324F',opacity:.13,blur:4,offset:3,angle:90},objectName:'source-panel'});
await deck.writeFile({fileName:'/work/output.pptx'});
```

Use inches for geometry and points for text. PptxGenJS creates new decks; it cannot load an existing PPTX. No direct native gradient-fill API was verified in 4.0.1. Prefer Office Kit for designs that require that effect; native band approximations have visible boundaries and excessive object counts. Do not use an untested `gradient` option and assume it works. CJK fonts must exist on the rendering/viewing host; do not assume embedding.

For extended API options consult the bundled upstream `pptxgenjs-api.d.ts` and official docs. These third-party declarations retain MIT notices and are not SlideRecast's independently authored logic.

## Rendering and fallback

For mixed PDF page sizes/orientations, inspect each rendered page's actual dimensions. A PPTX deck has a shared slide size. Preserve proportions on a consistent canvas with disclosed padding, or produce separate decks when requested; do not stretch pages or silently crop their content. Crop boxes/rotations can differ from the PDF media box, so use the visible raster as the geometry reference.

LibreOffice headless can export a **copy** to PDF, then the external Poppler process can rasterize it. A dedicated user profile avoids interference between concurrent processes. `render.py` retains hashes and reports versions. If no LibreOffice exists, inspect available host tools; do not download a desktop installer into a cloud sandbox blindly.

If Preview is installed: import `loadPresentation/getSlides` from `@office-kit/pptx` and `renderSlideToImage/buildFontkitMeasurer` from `@office-kit/pptx-preview/node`. Pass a CJK file in `fontFiles` and register the same font with `buildFontkitMeasurer`. `auditTextLayout` is a font-dependent overflow heuristic, not a full rendering guarantee. Register CJK fonts or report its result as approximate. Renderers disagree on fonts, effects, and geometry. Compare discrepancies before changing content, retain both renders, and reserve Microsoft confirmation for an actual Microsoft application test.

Official references:

- https://github.com/office-kit/pptx
- https://office-kit.github.io/pptx/
- https://github.com/gitbrent/PptxGenJS
- https://gitbrent.github.io/PptxGenJS/docs/api-shapes/
- https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html
