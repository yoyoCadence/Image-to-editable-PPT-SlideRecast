# 安裝與使用 / Installation

這是 alpha 測試候選包，不是已通過 Claude.ai 驗收的正式版本。

## Claude.ai：一般使用者

1. 下載 `image-to-editable-ppt-claude-alpha.zip`。
2. 在 Claude 的 Settings > Capabilities 確認 Code execution and file creation 已啟用。團隊帳號可能需要管理員允許。
3. 到 Customize > Skills，按 `+`，選 Create skill > Upload a skill，上傳 ZIP 並啟用。
4. 在新的對話上傳你有權使用的 PNG/JPEG 或 PDF。
5. 輸入下列要求。下載輸出的 PPTX，查看比較圖與品質報告。

> 請使用 Image to Editable PPT，把附件忠實還原成可編輯 PowerPoint。保持原始文字、配色、比例和版面，不重新設計。主要文字和圖形要能個別修改。請渲染全部頁面、與原稿比對，修正差異後交付 PPTX，並說明不可編輯的裝飾層或尚未驗證的部分。

不用在自己的電腦安裝 Node、Python、VS Code 或 Microsoft Office。Claude 需要它自身雲端工具的支援；若該帳號缺少渲染或字型能力，它應明確說明限制，不能宣稱檢查通過。ZIP 已包含兩套 JavaScript 製作函式庫，不需由使用者執行 npm install。

若要可修改圖表底層數值，另加：「請將圖表重建為可編輯資料系列的原生圖表；若外观無法忠實保留，先說明折衷。不可推測不存在的精確數值。」此功能不是預設流程。

## English quick start

Enable code execution/file creation, open Customize > Skills, upload the Claude ZIP and enable it. Start a new conversation, attach an image/PDF, and request:

> Use Image to Editable PPT to recreate this attachment as a faithful editable PowerPoint. Keep the exact wording, layout, colors and proportions. Preserve individually editable text and major visual elements. Render every slide, compare against the source and fix local differences. Return the PPTX and disclose image-only layers, font substitutions and unverified checks.

All authoring occurs in Claude's cloud execution environment. No local coding tools or desktop Office subscription are required. Platform capabilities, account settings and fonts still matter. The current Claude.ai login wall prevented direct acceptance testing of this candidate.

## OpenAI candidate

`sliderecast-openai-plugin-alpha.zip` has root `plugin.json`, `skills/image-to-editable-ppt/SKILL.md`, bundled resources and square SVG icons. It is prepared for private package testing. No public submission or publishing is authorized. See RESEARCH.md and RELEASE-CHECKLIST.md for developer verification and review steps; do not assume uploading an archive automatically lists it publicly.

## Free application checks

Use LibreOffice Impress on a copy to select/change native text and shapes. Use a free Microsoft account with PowerPoint for the web to check how it displays the original PPTX; record restrictions if particular object editing is unavailable there. Keep original downloads untouched, and save application-edited copies under new filenames. No Microsoft desktop subscription is required for these basic checks. This project's Linux headless tests do not establish Windows or Microsoft UI behavior.
