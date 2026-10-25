# PDF forward reconstruction quality report

Converted `multipage.pdf` to `editable.pptx` using source-specific native Office Kit authoring code in `generate.mjs`. The source is three raster-only PDF pages, each 960 × 540 pt; the output preserves that 16:9 size, all three pages, and their order. Source and final slide rasters are 1440 × 810 px. No full-page screenshot, picture background, or hidden raster fidelity layer is included.

## Verified evidence

- Inspected every source page, every initial and final rendered slide, all final overlays and comparison images, plus enlarged text, icon and arrow regions. Major geometry, layering, and all visible wording are retained. No unresolved clipping or major overlaps were observed.
- `source-transcription.json` is a separate visual transcription. `inspection.json` matches all 39 native text objects across the three slides, with no missing/extra text. This is an agent transcription check, not third-party OCR certification.
- Native elements include all text; cards and rounded bars; the gradient; shadows; grouped icons; forward connectors and the correction loop. Native XML inventories are 23 / 22 / 21 shapes, 2 / 18 / 4 connectors, and 3 / 8 / 0 groups. Shape counts include text-bearing shapes. There are zero picture layers and no native chart/workbook. The chart deliberately uses editable bars, labels and gridlines and preserves the explicitly visible values 32, 58, 76, 92.
- ZIP/XML/relationship/order checks passed, with no errors, warnings or outside-canvas objects. These are structural checks, not complete Open XML schema validation or application-level editability proof.
- All slides were rendered from a copy by **LibreOfficeDev 26.8.0.0.alpha0**, build `2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`; final PDF rasterization used **pdftoppm 26.05.0**. The final PPTX hash is recorded in `render-final/render.json`; the renderer retained the PPTX bytes unchanged.

## Targeted corrections

| Region/object | Initial cause | Targeted change | Observed result |
| --- | --- | --- | --- |
| Titles, numbers and other text | Estimated type sizes made several labels too wide and roughly 4–9 px too high | Adjusted individual sizes and positions in the `adjust` map, keeping wording and major geometry | Text placement and extents now closely follow source; no wrap/clipping seen |
| Small material label on pages 2–3 | First full-page view missed the repeated header | Enlarged source header crops; restored the exact native text on both pages | All three headers appear and all 39 text objects match transcription |
| Page 1 circular inspection symbol | Thin divider estimate missed the visible central white rectangle | Changed only the divider to a 12 × 46 px native rectangle at source position | Source/initial/final crop shows restored central bar |
| Page 2 layer icon and monitor stand | Estimated icon colors/stand position differed locally | Sampled only source-raster color and adjusted native fills/stand | Detail comparison more closely follows source |

`initial.pptx`, `generate-initial.mjs`, and `render-initial/` preserve the starting result. `ring-source-initial-final.png` and `correction-detail-page-*.png` preserve visual correction evidence. Final all-page source/render, overlay, difference, and ranked-region evidence is under `compare-final/`.

| Page | Initial normalized RGB MAE | Final normalized RGB MAE |
| --- | ---: | ---: |
| 1 | 0.023209 | 0.006137 |
| 2 | 0.013685 | 0.003142 |
| 3 | 0.017266 | 0.003616 |

These are measured pixel errors, **not fidelity percentages** or semantic/editability scores. Images were compared at equal dimensions, with no shift alignment. Different text antialiasing, gradient rendering and shadows contribute to error.

## Remaining limitations

The source's font family cannot be identified reliably from raster-only pages. Installed **Noto Sans CJK TC** is used for Latin and East Asian text, with visually tuned sizes; fonts are not embedded. Small differences remain in mixed-language subtitle spacing, some glyph metrics, shadows/gradient rasterization, and icon stroke antialiasing. Underlying numeric chart data is not an editable Excel workbook; the individually editable shapes and exact labels preserve the requested default scope.

**Microsoft PowerPoint verification was not performed.** No actual Microsoft application selection, edit, save/reopen or repair-prompt test was available. Native OOXML structure and successful LibreOffice rendering support a useful editable reconstruction but do not establish Microsoft compatibility, perfect visual equivalence, or application-level editability.

## Scope and reproducibility

Used only the installed skill directory, named PDF, generated output files, available runtime libraries/fonts and renderer. No neighboring source images/SVG, fixture code, benchmark data, expected-text fixtures, private repository files, prior PPTX, network uploads, GitHub writes or additional agents were used.

Authoring runtime: Node 24.19.0 and bundled Office Kit 0.24.1. Regenerate with `node generate.mjs` in the same installed-skill environment, then use the skill's inspection/render/compare helpers. The generator intentionally has an absolute bundle import and output path for this authorized test directory.

Source SHA-256: `31fd13c6691e4371664268feb056977f1e5f60b255b64b0d272d848a2577d487`.

Archived generator paths are normalized to the repository-relative bundle/output folder. Run `npm run build` first. Reconstruction logic is unchanged. The raw evaluation checkpoint retains the original absolute-path generator and initial evidence.
