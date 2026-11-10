import fs from 'fs'
import isestr from 'wsemi/src/isestr.mjs'


//docx2html.exe(srcPython/docx2html.py之run)結束時於stdout印出單行結果: 成功為`success`, 失敗為`error: <訊息>`, 且不論成敗離開碼皆為0, 故離開碼無法判定成敗, 須依輸出判定
let codeFail = 'E_CONVERT_FAILED'


//failMsg, 組出失敗訊息`[E_CONVERT_FAILED] 原因`, 內部函數(不入JSDoc), 失敗碼與格式僅此一處
function failMsg(detail) {
    return `[${codeFail}] ${detail}`
}


//judgeConvertOutput, 依docx2html.exe之輸出(stdout與stderr合併之字串)判定轉換成敗, 內部函數(不入JSDoc), 回傳{ ok, msg }
//判準為fail-closed: 輸出中任一行以`error:`開頭視為失敗(即使同時有`success`行); 無`error:`行時須有整行為`success`之行才視為成功; 其餘(空輸出, 未知輸出, 非字串)一律視為失敗
//失敗訊息為`[E_CONVERT_FAILED] 原因`, 原因取自exe之`error:`行, 無`error:`行時為固定說明, 不回顯其餘輸出內容
//stderr可能有`Word instance is kept since ...`(R-WORD-02, 不以`error:`起首), 合併後不影響判定
function judgeConvertOutput(out) {
    let lines = isestr(out) ? out.split(/\r?\n/).map((v) => v.trim()) : []

    //exe自報失敗, 原因為`error:`之後之文字
    let lineErr = lines.find((v) => v.startsWith('error:'))
    if (lineErr !== undefined) {
        let detail = lineErr.slice('error:'.length).trim()
        return { ok: false, msg: failMsg(detail !== '' ? detail : 'docx2html.exe reported an error') }
    }

    //成功須有明確之success行
    if (lines.includes('success')) {
        return { ok: true, msg: '' }
    }

    return { ok: false, msg: failMsg('docx2html.exe returned no success result') }
}


//judgeConvertFile, exe報成功後另驗輸出檔: fpOut須為存在且大小>0之檔案, 否則視為失敗, 內部函數(不入JSDoc), 回傳{ ok, msg }
//失敗訊息為固定說明, 不回顯路徑; 例: fpOut無副檔名時Word另存為`<fpOut>.htm`, exe仍報success而fpOut不存在
function judgeConvertFile(fp) {
    let ok = false
    try {
        if (isestr(fp)) {
            let st = fs.statSync(fp)
            ok = st.isFile() && st.size > 0
        }
    }
    catch (err) {
        ok = false
    }
    if (ok) {
        return { ok: true, msg: '' }
    }
    return { ok: false, msg: failMsg('output file was not generated or is empty') }
}


export default judgeConvertOutput
export { judgeConvertFile }
