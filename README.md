# Image to Editable PPT — SlideRecast

從圖片與 PDF 忠實重建可編輯 PowerPoint 的 Agent Skill。優先保留外觀與文字，讓主要元素可單獨修改。預設不重新設計，也不提取圖表底層數據。

**開發中，尚未正式發行。** Claude.ai 網頁版驗收與 Microsoft PowerPoint 相容性檢查尚未完成。預計 Apache-2.0，目前尚未授予該授權，須先完成權利來源確認。

Developed independently by the project maintainer.

已實作 Agent 自主重建流程、兩套離線 JavaScript 製作函式庫、PDF 頁面準備、原生物件檢查、LibreOffice 渲染、原稿差異圖與局部修正驗收。製作與 QA 腳本保持精簡，沒有強制 Scene Graph、固定卡片模板、API Server 或 MCP Server。

本地測試包括兩套引擎對同一組自製原稿的重建、原生文字／圖形／圖表數值修改後讀回、負向測試及独立 Agent 從原始 PNG/PDF 重新製作。**已知座標的函式庫測試與只看原稿的 Agent 測試分開記錄**，不以物件數量或像素誤差宣稱「100% 可編輯」或「還原率」。

一般使用者預期只需在 Claude.ai 啟用 Code execution and file creation，上傳 Skill ZIP，再上傳圖片或 PDF；不用在自己的電腦裝 Coding 工具。候選包內含 Office Kit `0.24.1` 與 PptxGenJS `4.0.1`，不必由使用者執行 npm install。Claude 雲端的 Node、字型、渲染工具及帳號政策仍需實測。

使用範例：

> 請將附件忠實還原成可編輯 PowerPoint，保持原始文字、配色、比例與版面，不重新設計。主要元素要能個別編輯。請渲染全部頁面與原稿比對，修正差異後交付 PPTX，並說明不可編輯的裝飾層或尚未驗證的部分。

![自製原稿與可編輯 PPTX 渲染比較](examples/results/office-kit-comparison-1/comparison.png)

- [安裝指南及更多提示詞](docs/INSTALLATION.md)
- [Skill 主檔](skills/image-to-editable-ppt/SKILL.md)
- [實測紀錄與限制](docs/TEST-LOG.md)
- [架構](docs/ARCHITECTURE.md)、[官方研究](docs/RESEARCH.md)、[私人電腦開發流程](docs/DEVELOPMENT.md)
- [待辦](TASK.md)、[發布檢查](docs/RELEASE-CHECKLIST.md)、[隱私](docs/PRIVACY.md)
- [獨立開發來源](docs/PROVENANCE.md)、[第三方授權](THIRD-PARTY-NOTICES.md)、[授權確認](docs/LICENSING.md)
- [English README](README.en.md)

統計圖表預設以忠實還原外觀為目標。只有使用者要求修改底層數值或同意採用原生圖表時，才使用資料系列。圖片、SVG 與不可編輯的內部細節會明確揭露。

候選 ZIP 是測試成果，沒有自動上架資格。Claude.ai 真實帳號、Microsoft PowerPoint 網頁版、個人 Windows 環境及公開目錄驗收尚未完成。維護者已於 2026-10-25 授權上傳此批成果、建立 PR 並合併；正式發行及目錄提交仍待另行確認。未宣稱與 Anthropic、OpenAI 或 Microsoft 有官方合作。
