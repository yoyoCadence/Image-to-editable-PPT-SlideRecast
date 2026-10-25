# Architecture and product boundary

Decision date: 2026-10-13. Status: development preview.

The product is an Agent Skill, not a fixed recognition/conversion engine. The agent visually inspects the user's source, chooses suitable native objects or isolated decoration layers, writes source-specific generation code, renders the output, compares it, and corrects faulty regions. A future stronger model can improve the reconstruction without changing a prescribed scene schema.

## Components

| Component | Responsibility | Excluded |
| --- | --- | --- |
| `skills/image-to-editable-ppt/SKILL.md` | Fidelity policy, task workflow, tool selection, truthful delivery | Redesign, forced templates, forced native charts |
| `references/` | Pinned APIs, rendering caveats, acceptance method | Mandatory layout/scene model |
| `vendor/` built by esbuild | Offline JavaScript authoring dependencies plus full notices | Preview native binaries, desktop Office, bundled fonts |
| `scripts/` inside skill | Capability probe, PDF preparation, read-only package inventory, copy rendering, image comparisons | OCR model, universal vision detector, automatic trading/API/MCP services |
| Root `scripts/`, `tests/`, `examples/` | Reproducible developer benchmarks and distribution assembly | End-user development environment |
| `plugin.json` | Portable Agent Plugins manifest and OpenAI presentation metadata | Directory approval or automatic publication |

Workflow: inspect source → choose technique → generate → package check → render → compare and inspect → correct local objects → repeat → report. Major text and geometry should be native and individually editable. An SVG or bitmap is explicitly an image layer. Rendering and object structure are separate evidence.

## Tool decision

Office Kit 0.24.1 has direct native gradient/effect authoring and read/edit/save capabilities. Its package requires Node >=22.18.0 and its pre-1.0 API is not frozen. PptxGenJS 4.0.1 is a viable new-deck authoring alternative. The same fixtures showed comparable simple shape/text output, while a PptxGenJS native-band approximation produced inferior gradient boundaries. Prefer Office Kit for gradient-rich tasks on a supported runtime, without making it mandatory for every input. Do not force a two-engine round trip. The native-edit proof deliberately tests Office Kit editing both generators' output, independently of the default conversion flow.

Preview 0.15.0 is a development cross-check. It requires native resvg and font resources on its Node path and is omitted from compact skill bundles. Its output is approximate and its layout audit depends on font metrics. LibreOffice is an independent renderer, not Microsoft PowerPoint. Neither proves Microsoft fidelity.

## Charts

Default chart visuals are reconstructed with native shapes, labels and lines. The test fixture's bar appearance is retained without inserting native chart data. Native charts are opt-in when the user specifically asks to edit underlying values/Excel series or agrees to the tradeoff. The optional native-chart probe verifies capability, not the product's primary success metric. Chart-to-Excel is a future independent product boundary.

## Packaging and runtime

Two bundles total approximately 1.28 MB before ZIP compression. Their exact byte counts and hashes are generated. They require no package installation for JavaScript authoring. Node and rendering/font capabilities remain host-provided. Detect capabilities instead of assuming Claude.ai equals this Linux runtime. Python rendering uses external LibreOffice/Poppler processes and optional Pillow/pypdf. PyMuPDF is confined to developer fixture generation and is not bundled/imported by skill helpers.

No server, credentials, telemetry, website, account or MCP service is needed for the MVP. Files remain in the current platform's execution environment. No uploaded source is transmitted by helper scripts.
