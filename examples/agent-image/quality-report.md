# Independent PNG-to-editable-PPTX conversion

The one-page source `complex-flow.png` was reconstructed as one editable 12 × 6.75 inch slide, preserving its 1440 × 810 aspect ratio, visible wording, four-step layout, icon colors, card shadows, arrow directions, and correction loop.

## Verified

- Read only the named PNG, the requested skill package, runtime/font information, and this output folder. No neighboring source SVG, fixture script, expected-text file, prior deck, or other private project was used. No network uploads or GitHub writes occurred.
- Independently transcribed all nine text strings into `expected-text.json`; final XML text matches exactly, with no missing/extra strings. The visible Chinese wording was reviewed against the PNG.
- `inspection.json` reports one ordered slide, valid ZIP/XML/relationship checks, no errors/warnings, and no objects outside the canvas.
- The slide contains 33 native shape/text objects: nine text objects and 24 nontext shapes/paths. There are four native shadows and zero pictures, charts, gradients, or groups. Cards, discs, magnifier, tiles, monitor, checkmark, lines and arrowheads are native editable geometry. There is no source screenshot background.
- Rendered every slide through **LibreOfficeDev 26.8.0.0.alpha0**, build `2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`. Original PPTX bytes were retained during rendering. Final render is 1440 × 810, so comparison required no size adjustment.
- Inspected the complete source/render pair, the overlay, and enlarged header, labels, icons, caption, and footer details. All major content is visible; no missing words, unintended wrapping, clipping, overlapping labels, displaced cards, or reversed arrows remain.

## Corrections made

| Region | Observed issue | Correction | Observed result |
| --- | --- | --- | --- |
| Text throughout | Initial text ink was 5–9 source pixels too high because native text-box top differs from glyph top | Moved the headline, subtitle, step labels, loop caption and footer independently | Baselines now closely follow the source; all text remains readable |
| Eyebrow, subtitle, step labels, caption | Smaller text initially exceeded source widths | Reduced only affected font sizes; preserved wording | Labels and caption closely match source widths |
| Eyebrow | Regular weight was visibly lighter than the source | Applied bold | Weight now follows the source |
| Rebuild icon | Initial teal and back-tile fills differed from the PNG | Sampled the permitted raster; changed fills to `00A68C` and `A8E5DC` | Disc and tiles now use the measured source colors |

Before/after evidence is retained in `render-before/`, `render-after/`, `render-final/` and corresponding comparison folders. The final comparison's normalized RGB MAE is **0.0039009869**; this is a pixel-difference diagnostic, not a fidelity/success percentage. Changes above the diagnostic threshold include small glyph and shadow/rasterizer differences.

## Limits and remaining differences

- The PNG contains no reliable font metadata. Installed **Noto Sans CJK TC** was used and named for Latin/East Asian runs; fonts are not embedded. The exact source font identity is unverified. Minor subtitle/caption glyph spacing, antialiasing and shadow softness differ on close inspection.
- Arrow geometry is editable native line/path shapes, not automatically attached PowerPoint connectors. Icon pieces are individually editable and named, but not grouped.
- Microsoft PowerPoint rendering, manual selection/editing in a desktop UI, and full Open XML schema validation were not performed. Structural native-object evidence does not replace those checks.
- The bundled API declaration files referenced by `tools.md` were initially absent; the parent build populated them before the authoring API details were consulted. No package installation or unverified API workaround was needed.

## Deliverables

- `editable-flow.pptx` — final editable reconstruction.
- `generate.mjs` — source-specific generation code using bundled PptxGenJS 4.0.1.
- `expected-text.json`, `inspection.json` — independent text and package-check evidence.
- `render-final/slide-001.png`, `render-final/render.pdf`, `render-final/render.json` — all-page render evidence.
- `compare-final/comparison.png`, `overlay.png`, `difference.png`, `metrics.json` — final comparison evidence.
- `final-*-detail.png` — enlarged source-above/render-below review regions.

Archived code import path normalized to a repository-relative URL; reconstruction logic unchanged. Run `npm run build` before reproducing. Original evaluation retained the raw authored code in the development checkpoint.
