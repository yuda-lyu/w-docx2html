import ensureBins from './ensureBins.mjs'


//fnExe, docx2html執行檔名稱(須與bin/manifest.json之name一致)
let fnExe = 'docx2html.exe'


/**
 * 取得docx2html.exe，回傳docx2html.exe的絕對路徑
 *
 * 相容層：本檔為舊版對外介面，內部改經ensureBins取得，依bin/manifest.json之SHA-256與大小校驗，
 * 首次使用時下載(或自鏡像取得)，之後自快取取得，不再偵測當前工作路徑之src/與node_modules/w-docx2html/src/，
 * 環境變數W_BIN_CACHE、W_BIN_MIRROR、W_BIN_OFFLINE見README〈外部執行檔〉
 *
 * 因docx2html.exe只能用於Windows作業系統，故調用前須自行檢核作業系統
 *
 * @returns {Promise} 回傳Promise，resolve回傳執行檔路徑物件，內含fpExe為docx2html.exe的絕對路徑字串，reject回傳帶code之Error(E_OFFLINE_MISSING等)
 * @example
 * import autoDownloadFiles from 'w-docx2html/src/autoDownloadFiles.mjs'
 *
 * async function test() {
 *
 *     //autoDownloadFiles, 取得docx2html.exe, 取得失敗則reject
 *     let { fpExe } = await autoDownloadFiles()
 *
 *     console.log('fpExe', fpExe)
 *     // fpExe D:\xxx\node_modules\w-docx2html\bin-cache\<sha256>\docx2html.exe
 * }
 * test()
 *     .catch((err) => {
 *         console.log('catch', err)
 *     })
 *
 */
async function autoDownloadFiles() {

    //ensureBins
    let bins = await ensureBins({ names: fnExe })

    //fpExe
    let fpExe = bins[fnExe]

    return {
        fpExe,
    }
}


export default autoDownloadFiles
