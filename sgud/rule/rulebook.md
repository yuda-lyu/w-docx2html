# 規則帳本(w-docx2html)

本檔以規則為單位記程式碼須長成什麼樣;**上游真理源為主控制式規格**,本檔只記本套件之實作須長成什麼樣並指回規格.

## 索引

| 規則 | 內容 |
|---|---|
| R-BIN-DOCX2HTML-01 | 外部執行檔 `docx2html.exe` 之存放、取得、校驗與離線(B 類,釘選模式) |
| R-WORD | 轉檔器須自行建立 Word 實例、只結束自己者、輸入檔唯讀開啟(R-WORD-01~05) |
| R-TEST | 測試:需實際 `docx2html.exe` 之功能測試、〈Word instance isolation〉 |

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
- 變更紀錄: 2026-10-28 改造(主控任務 20261028054319,bin-docx2html);2026-11-10 換版為 Word 實例隔離之轉檔器(R-WORD,主控排程第 408 項;`bin/` 與 `src/` 兩處同步,清單待重產).

## R-WORD Word 實例須自行建立、只結束自己者

來源:主控排程第 408 項(2026-11-10,源自使用端之調查:兩行程同時轉檔互相關掉對方文件).範本為 w-html2docx 之 `sgud/rule/rulebook.md`〈R-WORD〉(同形站點另有 w-docx-split、w-docx2pdf);本節條號與其一致,內容依本套件(docx 輸入、html 輸出)調整.

| 規則 | 內容 | 站點 | 測試 |
| --- | --- | --- | --- |
| R-WORD-01 | 取得 Word 一律 `win32com.client.DispatchEx('Word.Application')`(CoCreateInstanceEx,每次另起 WINWORD;實測已有其他 Word 時亦另起).**不得**用 `Dispatch`:pywin32 之 `Dispatch` 內部先 `pythoncom.connect`(GetActiveObject)連上已在執行之 Word(他行程之轉檔或使用者開啟者),轉完 `Quit(0)` 即以不儲存結束對方 | `srcPython/docx2html.py` 之 `docx2html` | 〈Word instance isolation〉①②③④ |
| R-WORD-02 | 只結束自己之實例:收尾一律經 `closeOwnWord`——關閉本行程之文件(關閉失敗者仍屬本行程,以 COM 物件身分比對,不以路徑:他人可能重開同一路徑),再以 `countOtherDocs` 盤點實例內非本行程之文件(含不在 `Documents` 內之受保護檢視視窗 `ProtectedViewWindows`);有他人文件**或無法判定**(讀取失敗,有界重試 10×0.5 秒後)時**不 Quit**(stderr 印 `Word instance is kept since …`),實例已不存在(RPC 斷線類錯誤)則不再動作,否則 `Quit(0)`.why:`DispatchEx` 所建之實例仍登錄於 ROT,他行程之 `Dispatch` 可連上並於其中開文件(實測);寧可留下實例,不關他人之文件.限制:盤點與 Quit 之間他人仍可連入;他人取用本行程已開之同一文件物件時,本行程之 Close 仍會關掉它(須對方改用 `DispatchEx`) | `closeOwnWord`,`countOtherDocs` | ④ |
| R-WORD-03 | 例外路徑不殘留:`Visible`/`DisplayAlerts` 之設定與轉檔同在 `try` 內,任何例外(含開檔失敗、`SaveAs2` 失敗)皆經 `finally` 之 `closeOwnWord` | `docx2html` | ③ |
| R-WORD-04 | 輸入檔以唯讀開啟 `Documents.Open(fpIn, False, not isSamePath(fpIn, fpOut))`(位置參數依序為 FileName,ConfirmConversions,ReadOnly);轉檔器只 `SaveAs2` 至 `fpOut`.why:各轉檔各一個實例,兩轉檔以可寫方式同開同一檔時,後開者跳出「檔案使用中」對話框(可見)而停住(實測;DisplayAlerts=False 擋不住).**例外**:`fpOut` 即輸入檔時可寫開啟(同範本之同檔例外);於 Word 16.0 此情境改前即回報 error(儲存與開啟中之檔同名,`-2146821994`),唯讀開啟亦同一 error,故例外在本機不可觀察(等價變異),保留以維持改前語意.限制:唯讀只解此鎖檔情境,編碼、密碼、修復、安全警告等互動仍可能出現(改前即有) | `docx2html`,`isSamePath` | ②(同一輸入檔同時,錯開 1.5 秒,交錯);⑤(保留測試) |
| R-WORD-05 | 不開可見視窗:`app.Visible = False`,`app.DisplayAlerts = False`;測試中扮演他行程之 Word 亦一律隱藏 | `docx2html`,`test/tools/wordInstance.mjs` | 全組 |

盤點指令:

```
git grep -n "Dispatch(\|DispatchEx(\|GetActiveObject\|\.Quit(" -- srcPython
git grep -n "Documents.Open(" -- srcPython
```

判準:`srcPython/` 取得 Word 只有 `wc.DispatchEx('Word.Application')` 一處,無 `Dispatch(`;`.Quit(` 只在 `closeOwnWord`;`Documents.Open(` 以第三個位置參數 `not isSamePath(fpIn, fpOut)` 開啟.

建置:改 python 後須重建 exe 並同步 `bin/docx2html.exe` 與 `src/docx2html.exe`(R-BIN-DOCX2HTML-01 規則 2、3).`srcPython/build.bat` 為舊式多步驟(含 `pause`、會重產 `docx2html.spec`),不直接執行;實際建置指令:`python -m PyInstaller srcPython/docx2html.spec --noconfirm --clean --distpath <暫存夾>/dist --workpath <暫存夾>/build`(與 `pyinstaller -F docx2html.py` 同選項),再複製 `docx2html.exe` 至上述兩處.

驗證法:`test/tools/wordInstance.mjs` 為範本之可移植輔助(複製自 w-html2docx,另加 `startBorrower` 之 `attachMs`、`waitDocs` 選項與 RPC 斷線時不重試);`test/tools/docxFixture.mjs` 組最小 docx(僅段落文字,不含外部連結:含外部連結之 docx 開啟時 Word 會跳可見之「安全性注意事項」對話框);`test/tools/binEnv.mjs` 自 `bin/` 建本機鏡像.輸入為 docx 而無 http 資源可扣住轉檔,故以「多段落 docx＋大的 `execFontGrow`」拉長轉檔,並以可觀察條件同步(轉檔器之 WINWORD 出現;借用者見實例內文件數達 1=輸入檔已開啟).

## R-TEST 測試

- 需實際 `docx2html.exe` 之功能測試(`test/WDocx2html.test.mjs`)以本機鏡像供檔(不連外網):`convert` 案內嵌鏡像設定;〈Word instance isolation〉以 `test/tools/binEnv.mjs` 之 `setBinEnv`.
- 〈Word instance isolation〉(6 案:①③②a②b④⑤)須在**無其他 WINWORD**時執行(有則整組略過並印出原因:修前之轉檔器會借用並結束之,且④無法確定連上本行程之實例);略過或初始化前失敗時 `after` 不結束任何 WINWORD.收尾以 `killNewAutomationWords` 只結束本組新出現之自動化 Word;有強制結束者(含可見視窗監看所結束者)再以 `clearDisabledItems` 清除指向 repo 內檔案之 Word 停用項目——強制結束開著文件之 Word 後,Word 將該文件列入 `HKCU\Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems`,之後開該檔即跳出可見訊息框而停住(實測).測試是否可自動刪除使用者登錄中指向 repo 之停用項目,待主控裁示(範本 hd408 報告〈需裁示〉).
- 〈Word instance isolation〉與 `convert` 同檔:mocha `--parallel` 下不同檔會同時用 Word.
- 改動範圍 → 測試:改 `srcPython/docx2html.py` → 重建 exe 並同步兩處 → 清單同步後跑 `WDocx2html.test`(`convert` 與〈Word instance isolation〉)與 `unit-binManifest`;本套件無瀏覽器 e2e.
