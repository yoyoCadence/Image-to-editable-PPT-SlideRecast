# Development instructions

Use the current product direction: fidelity first, editable major elements, agent reasoning, render/compare/local correction. Native charts are opt-in. Never redesign source documents by default.

Develop all project logic, documentation, branding and test materials independently. Do not access, clone, fork, copy, translate, or rewrite YoyoVector/PPT_Master or any other private code as a source. Use only this repository and documented public official APIs/licensed dependencies. Do not use company/client documents. Keep provenance in docs/PROVENANCE.md.

Do not publish a release, grant the proposed Apache license, upload a public directory submission, or change repository visibility without the maintainer's explicit confirmation. Current public GitHub writes were blocked by automatic approval review; continue local work and request confirmation for the concrete reviewed commit/PR.

Use pinned dependencies. Keep source PPTX bytes intact during rendering. Test changes with `npm run build`, `npm run fixtures`, `npm run benchmark`, `node scripts/edit-proof.mjs`, `npm test`, and `npm run package` as relevant. Inspect comparison images. Re-run only affected checks after subsequent changes.

Never report known-coordinate fixtures as blind conversion quality, package checks as schema validation, programmatic edits as Microsoft UI verification, or local bundle tests as Claude.ai tests. Mark untested items explicitly. Update TASK.md and test evidence with every milestone.
