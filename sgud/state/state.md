# 狀態

本檔記 w-docx2html 各功能之**現況**(做到哪裡、對應之規則帳本條目與測試檔、已知缺口與待辦、最近一次驗證);只記現況、不寫過程(過程見 `sgud/exp/experience.md`).

本檔 2026-10-28 新建(此前無 sgud).最近更新: 2026-11-10(dh408 第一段).`package.json` 之 `version` 未動(1.0.39).

## docx 轉 html(`WDocx2html`)

- 現況: 僅 Windows;以 `docx2html.exe`(srcPython 以 win32com 驅動 Word,PyInstaller 打包)轉換,輸入 base64 JSON(`fpIn`、`fpOut`、`execFontGrow`).
- 測試: `test/WDocx2html.test.mjs`(需本機有 Word;執行檔以本機鏡像供檔,斷言取自校驗後之快取).指令: `npm test`.
- 已知行為(改前即有,未改): `WDocx2html` 於轉換失敗時 exe 仍以標準輸出回報 `error: …`、離開碼 0,JS 端不檢查,故仍回 `ok`(無輸出檔);另案處理須經主控裁示(會改變對外行為,見 `sgud/feedback/feedback.md`).

## Word 實例(規則 R-WORD-01~05)

| 項目 | 現況 | 規則 | 測試 |
| --- | --- | --- | --- |
| 取得 Word | `srcPython/docx2html.py` 之 `docx2html` 以 `wc.DispatchEx('Word.Application')` 另起實例(原 `Dispatch` 會連上已在執行之 Word) | R-WORD-01 | 〈Word instance isolation〉①②④ |
| 收尾 | `closeOwnWord`:只關本行程之文件;`countOtherDocs` 以物件身分盤點他人文件(含受保護檢視);有他人文件或無法判定時不 Quit;`Visible`/`DisplayAlerts` 已移入 `try` | R-WORD-02,03 | ③④ |
| 開檔 | 輸入檔唯讀開啟;`fpOut` 即輸入檔時可寫開啟(`isSamePath`,本機不可觀察,見規則帳本) | R-WORD-04 | ②(同一輸入檔同時,錯開 1.5 秒,交錯),⑤ |
| 執行檔 | `bin/docx2html.exe` 與 `src/docx2html.exe` 已為 dh408 重建者(sha256 `2a34ef38d2ad8e4d9df2c1b0117528198da7b1fe46eaca3101a3ee36a8df79bb`,9515276 位元組;前版 `cc9bb43b…`,11183560),**待主控提交**;`bin/manifest.json` 仍為前版(commit `924a7a5…`),**待第二段以 `genBinManifest --commit` 重產**;清單未重產前 `npm test` 為 10 passing/12 failing(皆為清單與實檔不符之預期失敗) | R-BIN-DOCX2HTML-01,R-WORD | 模擬夾(新 exe＋清單改為其值)全套 27 passing/0 failing |
| 驗證法 | `test/tools/wordInstance.mjs`(可移植輔助,含可見視窗監看)＋`docxFixture.mjs`＋`binEnv.mjs` ＋ `WDocx2html.test` 〈Word instance isolation〉6 案 | — | 修前 exe 紅(①③②a②b④),新 exe 綠 |

- 已知限制:他行程(舊寫法)先 Quit 本行程之實例,或取用本行程已開之同一文件物件時,本行程仍會失敗或關掉該文件,須對方改用 `DispatchEx`(w-html2docx 已修;w-docx-split、w-docx2pdf 由主控另派);盤點與 Quit 之間他人仍可連入;保留實例時之 stderr 提示 JS 端不轉出;`DispatchEx` 本身於取得控制代碼前失敗時無從收尾.
- 未實測:使用者開著可見之互動 Word 時(不得開可見視窗,依機轉推論不受影響);轉檔中使用者自檔案總管開 docx 是否經 DDE 進入本行程之隱藏實例;`countOtherDocs` 之「無法判定→保留實例」分支無法注入,只以讀碼驗證;輸出等價(L3)僅以不含圖片與表格之 4 組輸入比對(HTML 位元組全同).

## 外部執行檔 `docx2html.exe`(R-BIN-DOCX2HTML-01)

- 現況: 已改制式規格(2026-10-28,主控任務 20261028054319,B 類→釘選模式): `bin/docx2html.exe`(主控已提交推送,提交 `924a7a516ef2ac01b6566e866b13383a2e8bc2d3`)＋`bin/manifest.json`(由 `toolg/genBinManifest.mjs` 產生)＋`src/ensureBins.mjs`(轉呼叫 wsemi `fsEnsureBins`,wsemi 1.9.27).`WDocx2html()` 經 `ensureBins({ names: 'docx2html.exe' })` 取得路徑;`WDocx2html.ensureBins` 與 `src/ensureBins.mjs` 供部署前預先取得;`autoDownloadFiles()` 保留為相容層.已移除 `postinstall`、`scripts/install.mjs`、`src/downloadFiles.mjs`(自寫之 `fsDownloadFile` 直連).
- 舊版相容: `src/docx2html.exe` 保留於 repo(舊版自主幹下載),不入 npm 包.
- 規則: `sgud/rule/rulebook.md` R-BIN-DOCX2HTML-01.
- 測試: `test/unit-binManifest.test.mjs`、`test/unit-ensureBins.test.mjs`(本機 HTTP `127.0.0.1:8951`、不連外網;子行程工具 `test/tools/ensureBinsChild.mjs`).
- 已知缺口與待辦: ①真實 GitHub 網址之下載未實測(測試不得連外網);②至舊版淘汰後移除 `src/docx2html.exe`(另案);③`dist/` 隨發布重新產製(`toolg/gDistRollup.mjs` 已增列 `url`、`crypto`、`stream`、`stream/promises` 為 external/globals;以該設定建置到暫存夾並模擬安裝實測: `import.meta.url` 於 UMD 可用、`ensureBins` 靜態屬性可用、經 UMD 實際轉換之輸出等於 `test/docoutTrue.html`;repo 內 `dist/` 仍為舊版,待發布時產製).
- 最近一次驗證(2026-10-28): `npm test` 21 passing(3 檔: `unit-binManifest`、`unit-ensureBins`、`WDocx2html`);變異 13 式皆殺;報告見主控任務 20261028054319 `agent-out/bin-docx2html-修正報告.md`.
- 最近一次驗證(2026-11-11,dh408 第一段): 模擬夾(新 exe＋清單改為其 sha256/大小)`npm test` 同指令 27 passing/0 failing(含〈Word instance isolation〉6 案);repo 內 `npm test`(清單未重產)10 passing/12 failing,12 條皆為清單與 `bin/` 不符;變異 M2(收尾一律 Quit)被④殺、M3(可寫開啟)被②b 殺、M4(一律唯讀)存活(等價).
