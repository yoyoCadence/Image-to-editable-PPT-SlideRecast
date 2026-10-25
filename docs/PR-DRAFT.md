# Pull request description

Title: Add fidelity-first image/PDF reconstruction skill and verifiable examples

Base: `main`

Head: `feat/fidelity-first-mvp`

## Description

SlideRecast introduces an independently authored Agent Skill that recreates uploaded images and PDFs as visually faithful editable PowerPoint presentations. The workflow preserves wording/layout, chooses native objects according to the source, renders the output and corrects local differences. Native chart data reconstruction is opt-in.

The change bundles pinned Office Kit and PptxGenJS dependencies with full notices, adds original source examples, separate structure/visual checks, paired-engine and independent agent evidence, native edit/read-back tests, English/Chinese documentation and private-test Claude/OpenAI package assembly. The image-size dependency is pinned to a patched version; package audit and PNG embedding evidence are retained.

Validation: automated quality/distribution tests, paired-engine all-page LibreOffice/Preview checks, independent PNG/PDF forward runs, extracted-ZIP authoring without npm installs, and portable manifest schema validation. Exact results and sample denominators are in docs/TEST-LOG.md.

Limitations: Claude.ai remains behind an authentication wall; Microsoft PowerPoint UI, Windows execution and full Open XML schema checks are unverified. License grant/formal publication require maintainer rights confirmation. Neither candidate ZIP was submitted to a public directory.

Maintainer authorized uploading these changes, opening a PR and merging on 2026-10-25. This is a development preview; formal release, directory submission and the planned Apache-2.0 grant remain pending. All 11 local automated tests passed again before upload.
