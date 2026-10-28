# 狀態

本檔記 w-docx2html 各功能之**現況**(做到哪裡、對應之規則帳本條目與測試檔、已知缺口與待辦、最近一次驗證);只記現況、不寫過程(過程見 `sgud/exp/experience.md`).

本檔 2026-10-28 新建(此前無 sgud).

## docx 轉 html(`WDocx2html`)

- 現況: 僅 Windows;以 `docx2html.exe`(srcPython 以 win32com 驅動 Word,PyInstaller 打包)轉換,輸入 base64 JSON(`fpIn`、`fpOut`、`execFontGrow`).
- 測試: `test/WDocx2html.test.mjs`(需本機有 Word;執行檔以本機鏡像供檔,斷言取自校驗後之快取).指令: `npm test`.

## 外部執行檔 `docx2html.exe`(R-BIN-DOCX2HTML-01)

- 現況: 已改制式規格(2026-10-28,主控任務 20261028054319,B 類→釘選模式): `bin/docx2html.exe`(主控已提交推送,提交 `924a7a516ef2ac01b6566e866b13383a2e8bc2d3`)＋`bin/manifest.json`(由 `toolg/genBinManifest.mjs` 產生)＋`src/ensureBins.mjs`(轉呼叫 wsemi `fsEnsureBins`,wsemi 1.9.27).`WDocx2html()` 經 `ensureBins({ names: 'docx2html.exe' })` 取得路徑;`WDocx2html.ensureBins` 與 `src/ensureBins.mjs` 供部署前預先取得;`autoDownloadFiles()` 保留為相容層.已移除 `postinstall`、`scripts/install.mjs`、`src/downloadFiles.mjs`(自寫之 `fsDownloadFile` 直連).
- 舊版相容: `src/docx2html.exe` 保留於 repo(舊版自主幹下載),不入 npm 包.
- 規則: `sgud/rule/rulebook.md` R-BIN-DOCX2HTML-01.
- 測試: `test/unit-binManifest.test.mjs`、`test/unit-ensureBins.test.mjs`(本機 HTTP `127.0.0.1:8951`、不連外網;子行程工具 `test/tools/ensureBinsChild.mjs`).
- 已知缺口與待辦: ①真實 GitHub 網址之下載未實測(測試不得連外網);②至舊版淘汰後移除 `src/docx2html.exe`(另案);③`dist/` 隨發布重新產製(`toolg/gDistRollup.mjs` 已增列 `url`、`crypto`、`stream`、`stream/promises` 為 external/globals;以該設定建置到暫存夾並模擬安裝實測: `import.meta.url` 於 UMD 可用、`ensureBins` 靜態屬性可用、經 UMD 實際轉換之輸出等於 `test/docoutTrue.html`;repo 內 `dist/` 仍為舊版,待發布時產製).
- 最近一次驗證(2026-10-28): `npm test` 21 passing(3 檔: `unit-binManifest`、`unit-ensureBins`、`WDocx2html`);變異 13 式皆殺;報告見主控任務 20261028054319 `agent-out/bin-docx2html-修正報告.md`.
