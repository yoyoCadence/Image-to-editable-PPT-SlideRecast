# Third-party notices

SlideRecast's authored logic is independent. It uses publicly licensed packages, whose implementation remains the property of their authors. Do not describe bundled library code or copied public API declarations as newly authored project code.

The build derives the actual bundled package list, versions, hashes and full license texts from esbuild's input graph. See `skills/image-to-editable-ppt/vendor/manifest.json` and `THIRD-PARTY-NOTICES.txt` in each candidate ZIP. Those generated files are recreated from locked dependencies, rather than checked into the source tree.

- `@office-kit/pptx` 0.24.1: MIT, copyright 2026 Yuichiro Yamashita. Its upstream notices are retained.
- `pptxgenjs` 4.0.1: MIT. Its package license is retained.
- PptxGenJS's already embedded Microsoft JavaScript helper notice is retained separately. It is upstream helper code, not an official Microsoft partnership.
- `jszip`: available under MIT or GPL-3.0-or-later; this distribution selects the **MIT** option.
- `pako`: MIT and zlib-derived code. Both MIT and the original zlib notice from `lib/zlib/deflate.js` are retained.
- `inherits`: ISC. Remaining bundled dependency notices are listed in the generated manifest/text.
- `references/office-kit-api.d.ts` and `references/pptxgenjs-api.d.ts` are unmodified public package declarations, redistributed under their upstream MIT terms. Their original licenses are in the generated notices.

Preview (MIT), native resvg, fontkit and Preview's font assets are developer-test dependencies and are **not** included in Skill/Plugin ZIPs. Sharp (Apache-2.0) is used only to create developer fixtures; its native binaries are not bundled. LibreOffice is an external application, not redistributed. The separately obtained Noto Sans CJK TC font is used for testing, not embedded or redistributed. Font substitution remains a runtime limitation.

Skill Python helpers use optional Pillow (MIT-CMU), pypdf (BSD-3-Clause), and external Poppler/LibreOffice processes. They do not import or bundle PyMuPDF. The developer-only synthetic PDF fixture generator uses PyMuPDF (AGPL-3.0/commercial dual-licensed); its library/binaries are not shipped in Skill/Plugin ZIPs. openpyxl and python-pptx are developer QA dependencies and are not redistributed. No final project license is granted by this document. See docs/LICENSING.md.
