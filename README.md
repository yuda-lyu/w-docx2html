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
