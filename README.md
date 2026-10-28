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
