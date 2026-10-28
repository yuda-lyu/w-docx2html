# 累積經驗

本檔為執行工作時特有之累積經驗,留給後續之嘗試、執行與判斷.

## 外部執行檔改制式規格(2026-10-28,bin-docx2html)

- 範本來源: 主控任務 20261028054319 `agent-out/wsemi-run04-修正報告.md` §8;規則見 `sgud/rule/rulebook.md` R-BIN-DOCX2HTML-01.
- `.npmignore` 須同時列 `test/_tmp`: 有 `.npmignore` 時 `npm pack` 不讀 `.gitignore`,而測試夾具(`test/_tmp/*/…/docx2html.exe`)在 `mocha --parallel` 下與 T1(`npm pack --dry-run`)並行,夾具中之 `.exe` 會被列入清單使 T1 偶發失敗.
- 主檔 `dist/w-docx2html.umd.js` 為預設匯出單一函數之 UMD;`ensureBins` 以靜態屬性 `WDocx2html.ensureBins` 掛上,不加具名匯出(具名匯出會使 `require()` 回傳物件而非函數,破壞既有 CJS 呼叫).
- `pkgRoot` 以 `import.meta.url` 推得(`src/` 與打包後 `dist/` 皆在套件根下一層;rollup 3 於 UMD 會把 `import.meta.url` 轉為取 `__filename`,以現行 `toolg/gDistRollup.mjs` 設定建置後模擬安裝實測可用);舊版 `autoDownloadFiles` 以 `process.cwd()` 找 `src/` 或 `node_modules/w-docx2html/src/`,工作目錄不是專案根時會找不到.新版與工作目錄無關.
- 驗證 `pkgRoot` 與工作目錄無關,**不能在測試行程內 `process.chdir` 後再呼叫**: `pkgRoot` 於模組載入時就固定,載入後才 chdir 抓不到「改取 `process.cwd()`」之退化(變異存活).須以子行程於他處工作目錄載入模組(`test/tools/ensureBinsChild.mjs`).
- 驗證「套件端不轉傳 `floatingManifestUrl`」須傳**無效值**(如 `not-a-valid-url`): 該選項只在主幹模式生效,傳有效網址時轉傳與否行為相同(等價變異);傳無效值則轉傳時 wsemi 以 `E_ARG` 拒絕.
- 以 tmp 副本跑變異(`.npmignore` 與清單、`ensureBins` 各選項、主檔屬性、相容層)13 式: 初跑 11 殺 2 存(上兩項),補強後 13/13 殺.`.npmignore`、`.gitignore`、`WDocx2html.mjs`、`README.md` 為 CRLF,以字串比對式變異須用 `\r?\n`.
- 功能測試(`test/WDocx2html.test.mjs`)需要實際執行檔: 以 `W_BIN_MIRROR`(本地資料夾,結構 `<鏡像>/w-docx2html/<sha256>/docx2html.exe`)＋`W_BIN_CACHE`＋`W_BIN_OFFLINE=1` 供檔,測試結束還原環境變數(mocha `--parallel` 之 worker 會被後續測試檔重用,環境變數不還原會外溢).
- 舊版相容: `src/docx2html.exe` 留在 repo(舊版自主幹下載它),不入 npm 包;更新執行檔時須同步 `bin/` 與 `src/` 兩處(`test/unit-binManifest.test.mjs` 比對兩者相同).
