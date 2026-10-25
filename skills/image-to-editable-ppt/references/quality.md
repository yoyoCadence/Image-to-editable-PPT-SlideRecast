# Acceptance and targeted correction

Review **both** fidelity and editability. No single combined percentage is defined.

| Check | Evidence | Limitation |
| --- | --- | --- |
| Text | Independent transcription, paragraph/object text comparison, readable render | XML text alone cannot detect clipped visible text |
| Page order | PDF manifest vs PPTX presentation relationships | File naming alone is insufficient |
| Editable major content | Native text/shapes/lines; select and edit in an application when available | Object count is not semantic editability or Microsoft validation |
| Visual fidelity | All-page comparison and detailed review of important regions | RGB error includes font and rasterizer differences |
| Missing content | Visual source checklist matched to output layers | No hardcoded object detector is supplied |
| Shadows/gradients/curves/icons | Compare details; inspect native fill/path/effect structure | SVG is a picture until converted to native geometry |
| Overflow | Render inspection plus font-aware layout audit where available | Bounds checks do not detect all text clipping |
| Corrections | Before/after render and local code edit log | A scripted defect test is not evidence of autonomous vision |

Suggested report fields: source pages, output slides, source aspect ratios, verified text objects, important native elements, picture-only layers and reason, renderer/version, inspected pages, corrections, remaining differences, and Microsoft verification status. Use measured values only.

`compare.py` reports normalized RGB MAE: sum of absolute RGB differences / (width × height × 3 × 255). It also reports fraction of pixels whose maximum channel difference exceeds 24. It resamples the render to source dimensions only if aspect ratio differs by at most 0.005, and does **not** shift-align images. Ranking 6×6 tiles helps locate differences; human/agent judgment decides which errors matter. A tile dominated by gradient or shadow differences may hide a small but serious text error. Always inspect text independently.

Correction example: if a title has the right wording but wraps to two lines, check the actual font and original line break first, then correct its width, margins, size or placement. Do not rewrite the sentence to fit. If only a connector points in the wrong direction, fix its arrow ends and preserve the rest of the slide.

When verification is unavailable, still deliver a useful draft if possible, label exactly which checks were skipped, and never promote the draft to a verified result.
