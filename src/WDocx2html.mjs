import path from 'path'
import process from 'process'
import get from 'lodash-es/get.js'
import isestr from 'wsemi/src/isestr.mjs'
import isp0int from 'wsemi/src/isp0int.mjs'
import str2b64 from 'wsemi/src/str2b64.mjs'
import execProcess from 'wsemi/src/execProcess.mjs'
import getPathParent from 'wsemi/src/getPathParent.mjs'
import fsIsFile from 'wsemi/src/fsIsFile.mjs'
import fsIsFolder from 'wsemi/src/fsIsFolder.mjs'
import fsCreateFolder from 'wsemi/src/fsCreateFolder.mjs'
import ensureBins from './ensureBins.mjs'


//fnExe, docx2html執行檔名稱(須與bin/manifest.json之name一致)
let fnExe = 'docx2html.exe'


function isWindows() {
    return process.platform === 'win32'
}


/**
 * 依照指定字串分割Docx檔
 *
 * @param {String} fpIn 輸入來源Html檔位置字串
 * @param {String} fpOut 輸入轉出Docx檔位置字串
 * @param {Object} [opt={}] 輸入設定物件，預設{}
 * @param {Integer} [opt.execFontGrow=0] 輸入全文件放大字型次數整數，預設0
 * @returns {Promise} 回傳Promise，resolve回傳成功訊息，reject回傳錯誤訊息
 * @example
 *
 * import w from 'wsemi'
 * import WDocx2html from './src/WDocx2html.mjs'
 * //import WDocx2html from 'w-docx2html/src/WDocx2html.mjs'
 * //import WDocx2html from 'w-docx2html'
 *
 * async function test() {
 *
 *     let fpIn = `./test/docin.docx`
 *     let fpOut = `./test/docout.html`
 *     let opt = {
 *         execFontGrow: 1,
 *     }
 *
 *     let r = await WDocx2html(fpIn, fpOut, opt)
 *     console.log(r)
 *     // => ok
 *
 *     w.fsDeleteFile(fpOut)
 *
 * }
 * test()
 *     .catch((err) => {
 *         console.log('catch', err)
 *     })
 *
 */
async function WDocx2html(fpIn, fpOut, opt = {}) {
    let errTemp = null

    //isWindows
    if (!isWindows()) {
        return Promise.reject('operating system is not windows')
    }

    //check
    if (!fsIsFile(fpIn)) {
        return Promise.reject(`fpIn[${fpIn}] does not exist`)
    }

    //check
    if (!isestr(fpOut)) {
        return Promise.reject(`fpOut[${fpOut}] is not an effective string`)
    }

    //execFontGrow
    let execFontGrow = get(opt, 'execFontGrow')
    if (!isp0int(execFontGrow)) {
        execFontGrow = 0
    }

    //轉絕對路徑
    fpIn = path.resolve(fpIn)
    fpOut = path.resolve(fpOut)

    //fdOut
    let fdOut = getPathParent(fpOut)
    if (!fsIsFolder(fdOut)) {
        fsCreateFolder(fdOut)
    }

    //prog, 經ensureBins取得docx2html.exe: 依bin/manifest.json校驗, 首次使用時下載(或自鏡像取得), 之後自快取取得, 見README〈外部執行檔〉
    let bins = await ensureBins({ names: fnExe })
    let prog = bins[fnExe]
    // console.log('prog', prog)

    //inp
    let inp = {
        fpIn,
        fpOut,
        execFontGrow,
    }
    // console.log('inp', inp)

    //input to b64
    let cInput = JSON.stringify(inp)
    let b64Input = str2b64(cInput)
    // console.log('b64Input', b64Input)

    //execProcess
    await execProcess(prog, b64Input)
        .catch((err) => {
            console.log('execProcess catch', err)
            errTemp = err.toString()
        })

    //check
    if (errTemp) {
        return Promise.reject(errTemp)
    }

    // //check
    // if (!isestr(output)) {
    //     return Promise.reject(`output[${cstr(output)}] is not an effective string`)
    // }

    return 'ok'
}


//ensureBins, 供部署前預先取得外部執行檔(以靜態屬性匯出, 不加具名匯出以免改變UMD主檔之形狀)
WDocx2html.ensureBins = ensureBins


export default WDocx2html
