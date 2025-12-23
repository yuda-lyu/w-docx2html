import fsDownloadFile from 'wsemi/src/fsDownloadFile.mjs'


async function downloadFiles(fdBase) {

    //url
    let url = `https://github.com/yuda-lyu/w-docx2html/raw/refs/heads/master/src/docx2html.exe`
    // console.log('url',url)

    //fn
    let fn = `docx2html.exe`

    //fp
    let fp = `${fdBase}${fn}`

    //fsDownloadFile
    console.log(`downloading url[${url}]...`, `to fp[${fp}]`)
    await fsDownloadFile(url, fp)

}


export default downloadFiles
