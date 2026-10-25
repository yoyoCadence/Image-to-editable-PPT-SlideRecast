# Dependency and input checks

The project processes user files in the current platform environment. It has no externally reachable server and no telemetry. Uploaded source text is treated as content, not instructions.

On 2026-10-13, `npm audit --omit=dev --json` identified two high-severity entries: image-size and its parent PptxGenJS, associated with image parser denial of service. The advisories identify patched image-size 2.0.3+. The lockfile overrides PptxGenJS's declared dependency to 2.0.4. The tested PptxGenJS dist bundle imports JSZip and the actual esbuild input graph did not include image-size; the package installation is nevertheless remediated rather than suppressing the alert. A PNG image embedding probe verifies the selected runtime path.

Advisories:

- https://github.com/advisories/GHSA-5p2g-fcmc-qvqq
- https://github.com/advisories/GHSA-w3rx-r6r6-pgpr

The generated audit evidence records the post-change result and exact scope. A zero known-advisory count is not a guarantee of complete security. Developer dependencies, native renderers, platform sandboxing and uploaded file parsers need ongoing updates and separate review. No dependencies or credentials from private projects are used.

ZIP packages exclude native executables/fonts and do not contain network secrets. Source PPTX files are not overwritten by rendering. Published issue reports should use synthetic reproduction inputs and avoid sensitive documents.
