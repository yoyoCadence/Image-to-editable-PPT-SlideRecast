# Test log — 2026-10-13

Status: local development preview, no formal release. All results refer to the pinned environment in PROVENANCE.md. Source PPTX bytes were retained. No company/client documents were used.

## Paired generator tests (known layout)

Three original synthetic designs (infographic, complex flow screenshot, styled bar visual) were reconstructed by each library: **2 decks, 6 rendered slides**. Coordinates and wording are shared with the fixture harness, so this is a generator/rendering capability test, not a blind image-conversion success rate.

Both decks passed ZIP/XML/relationship/page-order checks and exactly matched all **39 text objects per deck**, across **3/3 pages** each. Both contain zero source-page pictures. Office Kit uses one native gradient and 8 native shadows across the deck. PptxGenJS preserves gradient editability using many color bands, but the edge approximation is visible and is not recommended over a native gradient when fidelity matters.

The optional native-column-chart probe creates one data series with 32/58/76/92 for each engine. Office Kit then changes text, fill/position and chart values; reloading recovers 33/59/77/93 from both chart caches and the embedded XLSX workbook. This is programmatic object editing, **not** application UI selection or Microsoft compatibility.

## Pixel method

Normalized RGB MAE is sum absolute RGB differences / (width × height × 3 × 255). Changed fraction counts pixels with a maximum channel delta >24. Render dimensions are resampled only for a matching aspect ratio; no shift-alignment is applied. No fidelity percentage is defined. Preview-vs-LibreOffice measures renderer disagreement and cannot identify the authoritative Microsoft display.

Final paired benchmark uses LibreOffice PDF output rasterized by **Poppler 26.05.0**. Earlier initial evidence used PyMuPDF rasterization; do not compare mixed-backend scores as a correction improvement.

| Engine | Fixture | Source-vs-render RGB MAE | Changed-pixel fraction | Preview-vs-LibreOffice RGB MAE |
| --- | --- | ---: | ---: | ---: |
| office-kit | infographic | 0.018089 | 0.038235 | 0.017673 |
| office-kit | complex-flow | 0.010536 | 0.021332 | 0.010447 |
| office-kit | visual-chart | 0.012835 | 0.026983 | 0.013156 |
| pptxgenjs | infographic | 0.019044 | 0.042882 | 0.023898 |
| pptxgenjs | complex-flow | 0.010493 | 0.020865 | 0.017209 |
| pptxgenjs | visual-chart | 0.012722 | 0.026987 | 0.018211 |

Preview's font-aware audit initially identified three 4.89 px vertical-overflow warnings in numeric headings. Targeted text-box height changes removed those warnings without altering wording. Initial connector heads used a filled arrow preset; they were changed to native line segments to match the source's open arrow appearance, including the deliberately preserved source loop arrow geometry. Final Preview audit has zero reported issues on both decks; rendered visual inspection remains separate.

## Timing and file size

- pptxgenjs: generation 0.0169 s, LibreOffice plus rasterization 1.2262 s, PPTX 148,503 B.
- office-kit: generation 0.0833 s, LibreOffice plus rasterization 1.2981 s, PPTX 16,702 B.

One observed run per engine, with three slides each; these are not latency distributions. Library defaults/package compression differ, so file size is not an isolated measure of intrinsic efficiency. Timing excludes model reasoning, installation and manual review.

## PDF checks

Original image-only PDF: three pages rendered/prepared in order. A separate two-page selectable-text PDF verifies extraction/order (`First page`, `Second page`). Original inputs remain unchanged. Complex PDF fonts, crop boxes and mixed page sizes are not covered by those unit fixtures.

## Independent agent evidence

A fresh agent received only the canonical skill and the raw complex-flow PNG, with no SVG/fixture coordinates/expected answers/prior decks. It generated task-specific PptxGenJS code, rendered the result, and corrected local text positions/sizes/weight and icon colors. Parent review inspected its comparison image and reran its portable-path generator.

Result: 1/1 requested PNG converted, one slide, nine independently transcribed text objects exactly matched, 33 total native shape/text objects, four native shadows, zero pictures. Final same-stack normalized RGB MAE: 0.0039009869. Remaining font/shadow differences and absent icon grouping are documented in examples/agent-image/quality-report.md. This is **one** agent task, not a population success rate or Claude.ai evaluation.

A second fresh agent received only the skill and the three-page raster PDF. It authored an Office Kit deck with grouped native icons, native gradients/shadows, lines and shape-based chart bars. All **3/3 pages** were rendered and reviewed. The correction loop restored an initially missed repeated header on pages 2–3 and corrected text widths/placement and local icon geometry. Final **39** native text objects matched its independent visual transcription and also the parent's original fixture transcription. Zero pictures; final RGB MAE values: **0.006137 / 0.003142 / 0.003616**. Artifacts and limitations: examples/agent-pdf/quality-report.md.

These are two local agent tasks (one PNG, one PDF), sharing three original source designs. Completion counts are **1/1 PNG task and 1/1 PDF task**, not a general success-rate estimate. The agents were instructed to read only the skill and raw input; task source code/expected answers were withheld. Native structure and successful headless rendering still do not establish Microsoft UI editability.

## Distribution and structure

Candidate ZIPs contain bundled JavaScript and full notices, exclude native executables/fonts/Preview, and use platform-specific entry conventions. In a temporary extracted folder, both archives loaded both authoring engines, generated two small decks each and rendered all four outputs **without npm install and without project node_modules imports**. Host Node 24 remains provided; these are clean-folder tests, not a fresh OS image or Claude.ai runtime.

Portable plugin.json passed official Agent Plugins 1.0.0 JSON Schema with jsonschema 4.25.1. This does not validate every OpenAI extension field or establish uploader/review acceptance. Skill quick_validate passed; metadata description is 183 characters.

Tests: quality negative cases, text/native object inventory, chart cache/XLSX, original hashes, PDF order/text, pixel/aspect handling, archive layout and PNG embedding. Final run: **11/11 unittest methods passed**, including truncated-ZIP rejection. Exact output is in examples/results/test-run.txt. A candidate ZIP was once found truncated after packaging; the cause was not established. Packaging now validates a temporary archive before atomically replacing the final file; regeneration, CRC/hash checks and isolated loading all passed.

Production npm audit after the image-size 2.0.4 override reports zero known vulnerabilities. A PNG image embedding probe passed package inspection. This scan does not cover external renderers or prove absence of vulnerabilities. See SECURITY.md.

## Explicitly unverified

- Test 6: **BLOCKED**. Claude.ai browser navigation ended at a sign-in page. No ZIP upload, no real Claude conversion, no usage/token measurement.
- Microsoft PowerPoint for the web display/manual edit test: not executed.
- LibreOffice Impress manual selection/editing UI: not executed (headless rendering and programmatic editing completed).
- Full ECMA-376/Open XML SDK schema validation: not executed.
- Windows private-machine workflow: not executed in this Linux environment.
- User manual correction minutes and Claude execution cost: not measured.
- Public directory access, verified publisher identity, upload scans/review: not executed.
- Rights/contractual clearance, final project LICENSE, formal release: pending maintainer confirmation.
- GitHub upload/PR/merge explicitly authorized on 2026-10-25; remote completion is tracked separately from the local validation results.
- On 2026-10-25, all 11 unittest methods passed again before upload.

Final candidate sizes: Claude 537,473 B; OpenAI 541,834 B. Hashes are in docs/DISTRIBUTION-EVIDENCE.json. Both production-only and full npm audits reported zero known vulnerabilities after the override.
