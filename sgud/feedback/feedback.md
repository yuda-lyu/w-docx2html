# 回饋

任務中發現之自有套件、全域技能、全域設定之調整與抽提提案(只增不改;處置結果另記一列並寫明承接哪一列).

| # | 日期 | 類別 | 對象 | 內容 | 依據 |
| --- | --- | --- | --- | --- | --- |
| 1 | 2026-11-11 | 自有套件 | w-html2docx(`test/tools/wordInstance.mjs` 與其測試之 `after`) | 範本之收尾只在自己 `taskkill` 時才清 Word 停用項目(DisabledItems);但可見視窗監看(`startVisibleWordWatch`)結束之自動化 Word 若開著文件,同樣會被列入(本套件②b 修前紅跑實見一筆),`after` 應改為「自己結束或監看結束(`KILLED` 行)皆清」,已在本套件 `test/WDocx2html.test.mjs` 採用.同步複製該輔助之 w-docx-split、w-docx2pdf 宜一併採用 | `test/WDocx2html.test.mjs` 〈Word instance isolation〉之 `after`;本套件 dh408 ②b 修前紅跑後 DisabledItems 殘留一筆之紀錄 |
| 2 | 2026-11-11 | 自有套件 | w-docx2html(`src/WDocx2html.mjs`) | `WDocx2html` 於轉換失敗(輸入檔無法開啟、輸出路徑無效等)時仍回 `ok`:轉檔器以標準輸出回報 `error: …`、離開碼 0,JS 端不檢查(`src/WDocx2html.mjs` 之 output 檢查已註解掉),呼叫端只能以輸出檔是否產生判斷.修正會改變對外行為(失敗改為 reject),須主控裁示是否另案 | `src/WDocx2html.mjs:113-129`;本套件測試③直接呼叫 exe 取得 `error: …` |
| 3 | 2026-11-11 | 自有套件 | w-docx2html(`srcPython/build.bat`) | `build.bat` 為舊式手動多步驟(`pyinstaller -F`、不存在之 `build.spec`、onedir 各接 `pause`,並會重產 `docx2html.spec`),不能無人值守建置;宜改為與 w-html2docx 同形之單一指令(`python -m PyInstaller docx2html.spec --noconfirm --clean --distpath ..\bin`)並同步 `src/docx2html.exe`,另案 | `srcPython/build.bat`;本輪改以 `python -m PyInstaller srcPython/docx2html.spec …` 建到暫存夾再複製 |
| 4 | 2026-11-11 | 全域技能 | role-command-windows §2.3(或新建之 Office COM 自動化技能) | 補:Git Bash 會對以 `/` 起首之參數做路徑轉換,不只路徑——`mocha --grep "/pattern/"` 被轉成 `C:/Program Files/Git/pattern/` 而選不到任何案;排除單案可用 `--grep 關鍵字 --invert`,或 `MSYS_NO_PATHCONV=1` | 本輪 `tmp/quiet-run/…-dh408-red.log`:`0 passing (1ms)`,指令列顯示 `C:/Program Files/Git/isolation …` |
| 5 | 2026-11-11 | 全域技能 | 新建之 Office(Word)COM 自動化規範技能(承 w-html2docx feedback #4) | 補兩點:①輸入為 docx(無 http 資源可扣住轉檔)時,含外部連結(`a:blip r:link` External)之 docx 雖會使 Word 向網址取圖,但開啟後跳出可見之「安全性注意事項」對話框,不可用作扣住點;改以「多段落 docx＋拉長轉檔之參數」配合可觀察條件同步;②自 PowerShell 借用者讀 `Documents.Count` 須用 `$w.Documents.Count`,經 `InvokeMember('Count')` 回 `DISP_E_UNKNOWNNAME`,而寫 `Content.Text` 於 Word 忙碌時才需 `InvokeMember` | `sgud/exp/experience.md`〈Word 實例隔離〉;本輪 probe1/probe2/probe6 紀錄 |
