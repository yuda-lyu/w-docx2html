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

## Word 實例隔離(2026-11-10,dh408)

- 範本為 w-html2docx 之 hd408(規則 R-WORD,機轉探測見其 `sgud/exp/experience.md`〈Word 實例隔離〉):`Dispatch` 先 GetActiveObject 連上已登錄於 ROT 之 Word,轉完 `Quit(0)` 關掉他人文件;`DispatchEx` 另起實例但仍可被他行程之 `Dispatch` 連上,故收尾須盤點他人文件;兩獨立實例可寫開同一檔會跳可見之「檔案使用中」對話框,故輸入檔唯讀開啟.本套件依其改寫,未另行發明.
- **docx 輸入無 http 資源可扣住轉檔**(範本以 html 內的 http 圖扣住 Word 匯入).試過含外部連結圖之 docx(`a:blip r:link`、`TargetMode="External"`):Word 開啟時確會向網址取圖(扣住可行),但其後跳出可見之「Microsoft Word 安全性注意事項」對話框(`NUIDialog`),轉檔失敗(看門狗結束後 RPC call failed)或停住至逾時;一次懸住的實例還使後續 `Dispatch` 連上它而一併停住.故棄用,改以「多段落 docx＋大的 `execFontGrow`」拉長轉檔.
- 拉長轉檔的量:3000 段落 `execFontGrow` 10/40/60 約 3/6/8 秒;100 次且借用者同時輪詢曾達 96 秒(字級放大使版面暴增,不可預期);最終 5000 段落、60 次(約 10 秒,整案 12~14 秒).同步點用可觀察條件,不用固定等待:轉檔器之 WINWORD 出現(輪詢 `tasklist`);借用者以 GetActiveObject 連上(有界重試)後等 `Documents.Count >= 1`(輸入檔已開啟),再 `Documents.Add()`;對方新增文件時 `conv.done === false` 為斷言前提.
- 借用者讀文件數:`$w.Documents.Count` 可;經 `GetType().InvokeMember('Count', GetProperty, …)` 讀集合物件的 `Count` 回 `DISP_E_UNKNOWNNAME`(0x80020006),`InvokeMember` 只適用 Word 忙碌時讀寫 `Content.Text`.實例已被 Quit 時寫入回 `0x800706BE`(RPC call failed),重試無益,借用者遇 RPC 斷線類(`0x800706BA`/`0x800706BE`/`0x80010108`)即不再重試,修前紅測因此在數秒內以 `SETFAIL` 乾淨失敗而非逾時.
- `WDocx2html` 對轉換失敗仍回 `ok`:exe 以標準輸出回報 `error: …`、離開碼 0,JS 端不檢查(程式內檢查已註解掉).故成功案以「輸出檔是否產生且含來源標記」判定,錯誤案(輸出路徑無效、輸入檔不存在)直接呼叫 exe 檢查標準輸出.Word 找不到輸入檔之訊息含該檔名(路徑中段被截成 `...`).
- 輸出為 html:Word 的 filtered HTML 會在文字中折行,斷言前先 `replace(/\s+/g, ' ')` 並以單詞標記(無空白)比對.
- ⑤ `fpOut` 即輸入檔:改前即回報 error(儲存與開啟中之檔同名,`-2146821994`)且輸入檔不變;唯讀開啟亦同一 error.`isSamePath` 例外於 Word 16.0 不可觀察(變異體「一律唯讀」存活),測試只能是改前行為之保留測試.
- 看門狗結束的可見視窗所屬 Word 亦屬「被強制結束」:開著文件者會被列入 `DisabledItems`(②b 修前紅跑出現一筆,指向 `test/_tmp/…/race.docx`),範本之 `after` 只在自己 `taskkill` 時清除,本套件改為「自己結束或看門狗結束皆清」.紅跑結束後仍須以 `clearDisabledItems` 回驗.
- 新 exe 為 9515276 位元組(前版 11183560):建置環境 Python 3.14.3、PyInstaller 6.22.3(前版為 2025-12 建置);輸出經 L3 比對與前版位元組相同(4 組輸入).
- `srcPython/build.bat` 不可直接執行:三個 `pyinstaller` 指令依序為 `-F docx2html.py`、不存在的 `build.spec`、`docx2html.py`(onedir),各接 `pause`,且會重產 `docx2html.spec`.實際以 `python -m PyInstaller srcPython/docx2html.spec …` 建到暫存夾再複製.
- 模擬夾做法(第二段前驗證全套):依 `git ls-files -co --exclude-standard` 複製 repo 至 `./tmp/<任務>/rc/`,換入新 exe(`bin/` 與 `src/`)並把清單的 sha256/大小改為其值;不需 node_modules 連結(Node 自上層目錄解析 repo 的 `node_modules`),以 `spawnSync(node, [mocha.js, …], { cwd: rc })` 執行,清理時無連結待拆.
- Git Bash 對以 `/` 起首的參數做路徑轉換:`--grep "/isolation (…)/"` 被轉成 `C:/Program Files/Git/isolation …`,mocha 得 0 案;排除單案改用 `--grep 關鍵字 --invert`,中文 grep 字串直接傳遞可用.
- 變異驗證:變異版 py 複製到 `./tmp/<任務>/mut/<名>/` 以 `PyInstaller -F --specpath` 建單獨 exe,模擬夾換入後跑相關案:收尾一律 Quit 被④殺,可寫開啟被②b 殺(跳可見對話框,看門狗結束),一律唯讀存活(等價).
