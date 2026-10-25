# Feasibility research

Checked 2026-10-13. Official documentation and actual installed npm packages were used. No private source repository was consulted. This snapshot is not a promise of future platform availability.

## Packages actually resolved from npm

| Package | Latest queried version | Declared Node engine | npm unpacked size | License |
| --- | --- | --- | --- | --- |
| `@office-kit/pptx` | 0.24.1 | >=22.18.0 | 6,072,509 B | MIT |
| `pptxgenjs` | 4.0.1 | No `engines` value returned | 2,605,307 B | MIT |
| `@office-kit/pptx-preview` | 0.15.0 | >=22.18.0 | 8,326,867 B | MIT |

Commands: `npm view <package> version engines license dist.unpackedSize --json`, then exact-version installation. Source installation and lockfile, rather than GitHub's current main branch, define the tested APIs. Office Kit's docs state that minor releases can break its 0.x public API. Its core ESM bundle uses fflate. Preview's Node path also uses native resvg/fontkit and fonts. A docs helper requiring Node 22.18+ is not evidence that Claude.ai currently has that version.

The authoring comparison covers text, native shapes, shadow, gradient, layered native icons, line connectors, and an optional column-chart probe. It does not test every chart type, SmartArt, equations, animations, unsupported fonts, or all freeform curve cases. PptxGenJS cannot read existing decks; Office Kit can. A library's own CI validation claims were not adopted as validation of our outputs.

Official sources:

- https://github.com/office-kit/pptx and installed 0.24.1 README/types/license
- https://github.com/office-kit/pptx/tree/main/packages/preview and installed 0.15.0 README/types
- https://github.com/gitbrent/PptxGenJS and https://gitbrent.github.io/PptxGenJS/docs/api-shapes/
- https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html
- https://www.microsoft.com/en-us/microsoft-365/free-office-online-for-the-web

## Claude.ai

The current Help Center says Skills and code execution are available on Free, Pro, Max, Team and Enterprise, subject to settings/admin policy. Its current custom-skill guide describes a ZIP containing a single matching skill folder, required name/description metadata, and a description limit of 200 characters. The developer/API overview still contains older plan/settings text. Prefer the current product Help Center for web installation, but verify the actual account UI.

The Help Center permits Claude/Claude Code to install packages from approved repositories; it distinguishes API Skills, whose dependencies must be preinstalled. Code-execution docs describe network controls and a 30 MB **general file** limit. The custom-skill Help Center mentions ZIP size limits but does not establish a numeric **custom-skill ZIP** cap in the fetched pages. Do not substitute the API upload cap or general file cap. Our candidates are deliberately compact; uploader acceptance is unverified.

The web guide spells `skill.md`, whereas the Agent Skills specification uses `SKILL.md`. The repository and OpenAI package use uppercase; the Claude ZIP emits lowercase as described in the web upload guide. Actual case-sensitive acceptance is still a release check.

Attempted direct browser navigation to Claude.ai's Skills surface reached its sign-in page on 2026-10-13. No skill was installed, no conversion was run in Claude.ai, and no tokens/usage were measured there. Local clean-folder tests cannot substitute for that test.

Sources:

- https://support.claude.com/en/articles/12512180-use-skills-in-claude
- https://support.claude.com/en/articles/12512198-how-to-create-custom-skills
- https://support.claude.com/en/articles/12111783-create-and-edit-files-with-claude
- https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview
- https://agentskills.io/home

## OpenAI plugin distribution

Current official documentation supports skills-only packages and a portable root `plugin.json`, with `skills/`, optional `assets/`, and OpenAI metadata under `extensions.com.openai`. New packages should use the portable format. This project needs no MCP or lifecycle hooks. A legacy `.codex-plugin/plugin.json` is unnecessary for the chosen format.

Public submission requires the proper organization/project role (owner or Apps Management Write), verified individual/business developer identity, automated metadata/skill scans, review and a separate publish step. A skills-only plugin does not need MCP app test cases or an MCP demo recording. Adding MCP to an existing skills-only plugin is currently unsupported. Icons must exist at referenced relative paths and meet documented size/format constraints. Preparing a valid ZIP does not automatically grant publishing access or directory approval.

No submission portal was accessed and no directory review was requested. Publisher identity/access and actual upload acceptance remain pending.

Sources:

- https://developers.openai.com/plugins/build/plugins
- https://developers.openai.com/plugins/deploy/submission
- https://agent-plugins.org/schemas/1.0.0/plugin.schema.json
