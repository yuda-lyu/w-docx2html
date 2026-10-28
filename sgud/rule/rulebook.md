# 規則帳本(w-docx2html)

本檔以規則為單位記程式碼須長成什麼樣;**上游真理源為主控制式規格**,本檔只記本套件之實作須長成什麼樣並指回規格.

## 索引

| 規則 | 內容 |
|---|---|
| R-BIN-DOCX2HTML-01 | 外部執行檔 `docx2html.exe` 之存放、取得、校驗與離線(B 類,釘選模式) |

## R-BIN-DOCX2HTML-01 外部執行檔 docx2html.exe 一律經 `ensureBins()`(wsemi `fsEnsureBins`)取得

- 依據: 制式規格 `C:\opensrcOper\dspec\外部執行檔之存放下載校驗與離線.md`(使用者 2026-10-28 裁示;以下「規格」皆指此檔,條號 S/M/D/O/E/T 為其條號);主控任務 20261028054319(bin-docx2html).本套件於主控帳本 G10 為 **B 類**(改造前自主幹 `src/docx2html.exe` 直連下載、不校驗),採**釘選模式**.
- 單一擁有者: `src/ensureBins.mjs`(轉呼叫 wsemi `fsEnsureBins`,`pkgRoot` 由本檔位置推得);清單 `bin/manifest.json`(`toolg/genBinManifest.mjs` 產生,不手寫).
- 規則:
  1. **存放**(S1、S2): 執行檔放 repo `bin/docx2html.exe`(納入 git,`.gitattributes` 設 `bin/** binary`);npm 包只含 `bin/manifest.json`(`.npmignore`: `/bin/*`、`!/bin/manifest.json`、`/bin-cache/`、`src/docx2html.exe`、`sgud`).
  2. **舊版相容**(S2): 舊版執行時自主幹 `src/docx2html.exe` 下載,故 `src/docx2html.exe` **保留於 repo、與 `bin/docx2html.exe` 位元相同**(`test/unit-binManifest.test.mjs` 比對),不入 npm 包;至舊版淘汰後另案移除.
  3. **清單**(M1、M2、M4): `mode: "pinned"`,`commit` 為含 `bin/docx2html.exe` 之提交(40 碼),sha256 與大小由 `node toolg/genBinManifest.mjs --commit <提交>` 依實檔計算;`test/unit-binManifest.test.mjs` 比對清單與實檔.更新流程(U1): 替換 `bin/` 檔(並同步 `src/docx2html.exe`,見規則 2)→ 提交並推 GitHub → `genBinManifest --commit <該提交>` → 跑測試 → 發布新版.
  4. **取得**(D1~D4、O1~O3): 只經 `ensureBins({ names: 'docx2html.exe' })`;首次使用時下載,**不在 `postinstall` 下載**(已移除 `postinstall` 與 `scripts/install.mjs`);快取 `<套件根>/bin-cache/<sha256>/docx2html.exe`(`W_BIN_CACHE` 時為 `<W_BIN_CACHE>/w-docx2html/<sha256>/docx2html.exe`);鏡像 `W_BIN_MIRROR`;離線 `W_BIN_OFFLINE=1`.**不得**另寫下載、快取、鏡像、校驗(`src/downloadFiles.mjs` 之 `fsDownloadFile` 直連已刪除).
  5. **匯出**(D1、E2③): `src/ensureBins.mjs` 之預設匯出;主檔 `WDocx2html.ensureBins`(靜態屬性,因 UMD 主檔為預設匯出單一函數,不加具名匯出以免改變 `require()` 之形狀).`autoDownloadFiles()`(`src/autoDownloadFiles.mjs`)保留為相容層,回傳 `{ fpExe }`,內部亦經 `ensureBins`.
  6. `ensureBins()` 不轉傳 `opt.floatingManifestUrl`(改變清單信任錨).
- 盤點指令: Grep `fsDownloadFile|downloadFiles|raw/refs/heads` 於 `src/`、`scripts/`、`toolg/`(應無下載站點);`node C:/opensrcOper/tools/scan-bins.mjs --only w-docx2html`(應為 OK).
- 站點全集與判準: 外部執行檔 1 個(`docx2html.exe`,自有檔);取得站點 1 個(`src/WDocx2html.mjs` 之 `ensureBins`,以及相容層 `src/autoDownloadFiles.mjs`);下載站點 0 個(下載只在 wsemi `fsEnsureBins` 內).
- 測試: `test/unit-binManifest.test.mjs`(M1/M2/M4、T1、T2、S2 舊版相容)、`test/unit-ensureBins.test.mjs`(T3~T7、環境變數、pkgRoot 與載入時之工作目錄無關〔子行程工具 `test/tools/ensureBinsChild.mjs`〕、不轉傳 `floatingManifestUrl`、`WDocx2html.ensureBins` 與相容層;本機 HTTP `127.0.0.1:8951`、本地資料夾鏡像、不連外網;T8 為主幹模式專屬,不適用)、`test/WDocx2html.test.mjs`(功能測試,以本機鏡像供檔並斷言執行檔取自校驗後之快取).夾具 `test/_tmp/ensureBins/`、`test/_tmp/binManifest/`、`test/_tmp/WDocx2html/`,測試自行刪除.
- 變更紀錄: 2026-10-28 改造(主控任務 20261028054319,bin-docx2html).
