# w-docx2html
A tool for docx2html.

![language](https://img.shields.io/badge/language-JavaScript-orange.svg) 
[![npm version](http://img.shields.io/npm/v/w-docx2html.svg?style=flat)](https://npmjs.org/package/w-docx2html) 
[![license](https://img.shields.io/npm/l/w-docx2html.svg?style=flat)](https://npmjs.org/package/w-docx2html) 
[![npm download](https://img.shields.io/npm/dt/w-docx2html.svg)](https://npmjs.org/package/w-docx2html) 
[![npm download](https://img.shields.io/npm/dm/w-docx2html.svg)](https://npmjs.org/package/w-docx2html) 
[![jsdelivr download](https://img.shields.io/jsdelivr/npm/hm/w-docx2html.svg)](https://www.jsdelivr.com/package/npm/w-docx2html)

## Documentation
To view documentation or get support, visit [docs](https://yuda-lyu.github.io/w-docx2html/global.html).

## Core
> `w-docx2html` is based on the `win32com` in `python`, and only run in `Windows`.

## Installation

### Using npm(ES6 module):
```alias
npm i w-docx2html
```

#### Example:
> **Link:** [[dev source code](https://github.com/yuda-lyu/w-docx2html/blob/master/g.mjs)]
```alias
import w from 'wsemi'
import WDocx2html from './src/WDocx2html.mjs'
//import WDocx2html from 'w-docx2html/src/WDocx2html.mjs'
//import WDocx2html from 'w-docx2html'

async function test() {

    let fpIn = `./test/docin.docx`
    let fpOut = `./test/docout.html`
    let opt = {
        execFontGrow: 1,
    }

    let r = await WDocx2html(fpIn, fpOut, opt)
    console.log(r)
    // => ok

    w.fsDeleteFile(fpOut)

}
test()
    .catch((err) => {
        console.log('catch', err)
    })
```

## Breaking changes

### 轉換失敗改為 reject（下一版）

- 舊版：`docx2html.exe` 內部失敗（如輸入非 docx、docx 損毀、Word 無法開啟或無法另存）時，只在 stdout 印 `error: …` 且離開碼為 0，`WDocx2html()` 仍 resolve `'ok'`，呼叫端無從得知失敗；`fpOut` 已有舊檔者還會讀到舊內容而不自知。
- 新版：`WDocx2html()` 依 exe 輸出判定成敗，須有 `success` 行且無 `error:` 行才 resolve `'ok'`；`error:` 行（即使同時有 `success`）、空輸出、未知輸出一律以字串 reject，格式為 `[E_CONVERT_FAILED] 原因`，原因為 exe 輸出之診斷（已依系統字碼頁解碼，不回顯輸入與輸出路徑；但 Word 之診斷可能自行含檔名）。
- exe 報成功後另驗輸出檔：`fpOut` 須存在且非空，否則以 `[E_CONVERT_FAILED] output file was not generated or is empty` reject。例如 `fpOut` 無副檔名時 Word 會另存為 `<fpOut>.htm`，舊版回 `'ok'`、新版 reject；請給 `.html` 或 `.htm` 副檔名。
- 取不到執行檔：舊版 reject `ensureBins()` 拋出之 Error 物件（帶 `code`）；新版改為字串 `[code] 訊息`（如 `[E_OFFLINE_MISSING] …`），與 w-docx-split、w-html2docx 同型，本函數之 reject 一律為字串。
- 其他 reject 型態不變：參數檢測失敗為說明字串；exe 無法執行或離開碼非 0 為 `code[N]:\n…` 字串。轉換成功時之回傳值 `'ok'` 與輸出內容不變。
- 下游影響：原本以 `.then` 視為成功之呼叫端，遇上述失敗時改進 `.catch`；需區分失敗種類時以 `err.startsWith('[E_CONVERT_FAILED]')` 判斷。失敗時 `fpOut` 之內容不可採用（本套件不刪除 `fpOut` 既有之舊檔）。
- `docx2html.exe` 本身未重建（離開碼仍恆為 0），判定只靠輸出。

## 外部執行檔

本套件執行時需要下列外部執行檔，不含於 npm 包：

| 檔名 | 用途 | 來源 |
|---|---|---|
| docx2html.exe | 以 Word 將 docx 轉為 html（僅 Windows） | 本套件 GitHub repo 之 `bin/`（釘選提交） |

- 首次使用到轉換功能時自動下載（或預先呼叫本套件匯出之 `ensureBins()`），以 `bin/manifest.json` 所載之 SHA-256 與大小校驗後存於 `<套件根>/bin-cache/`；之後自快取取得（亦校驗）。安裝時不下載。
- 環境變數：
  - `W_BIN_CACHE=<資料夾>`：快取改放 `<資料夾>/w-docx2html/`（套件所在之 `node_modules` 唯讀時使用，多套件可共用同一資料夾）。
  - `W_BIN_MIRROR=<資料夾或 http(s) 網址>`：內網鏡像，先找鏡像再連 GitHub；結構為 `<鏡像>/w-docx2html/<sha256>/docx2html.exe`。鏡像之檔同樣以清單校驗，不符即不用。
  - `W_BIN_OFFLINE=1`：不連網，只用快取與鏡像；缺檔時錯誤（`E_OFFLINE_MISSING`）會列出缺哪些檔、sha256、大小與應放之完整路徑。
- 離線部署：
  1. 在可連網之電腦設 `W_BIN_CACHE=<空資料夾>`，呼叫本套件之 `ensureBins()` 一次（或使用到轉換功能一次）。
  2. 把該資料夾整個帶到內網電腦，設 `W_BIN_MIRROR=<該資料夾>`（全離線時另設 `W_BIN_OFFLINE=1`）。

```alias
import ensureBins from 'w-docx2html/src/ensureBins.mjs'
//import WDocx2html from 'w-docx2html'
//亦可用 WDocx2html.ensureBins()

async function test() {

    //部署前預先取得外部執行檔
    let bins = await ensureBins()
    console.log(bins['docx2html.exe'])
    // => D:\xxx\node_modules\w-docx2html\bin-cache\<sha256>\docx2html.exe

}
test()
    .catch((err) => {
        console.log('catch', err)
    })
```
