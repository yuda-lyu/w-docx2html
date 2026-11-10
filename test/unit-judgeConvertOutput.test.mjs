import fs from 'fs'
import path from 'path'
import assert from 'assert'
import judgeConvertOutput, { judgeConvertFile } from '../src/judgeConvertOutput.mjs'


//純函數測試(不需Word): docx2html.exe不論成敗離開碼皆為0, 成敗只能依輸出判定(R-CONVERT-RESULT-01)
describe('judgeConvertOutput', function() {

    let prefix = '[E_CONVERT_FAILED] '

    describe('success (output carries the "success" line and no "error:" line)', function() {

        it('judges the exe success output "success\\r\\n" as ok (R02: a valid conversion is not rejected)', function() {
            assert.strict.deepStrictEqual(judgeConvertOutput('success\r\n'), { ok: true, msg: '' })
            assert.strict.deepStrictEqual(judgeConvertOutput('success\n'), { ok: true, msg: '' })
            assert.strict.deepStrictEqual(judgeConvertOutput('success'), { ok: true, msg: '' })
        })

        it('tolerates extra lines around the success line, e.g. the stderr note "Word instance is kept since ..." merged by execProcess (R02, R-WORD-02)', function() {
            assert.strict.equal(judgeConvertOutput('Word instance is kept since 1 document(s) opened by another client\r\nsuccess\r\n').ok, true)
            assert.strict.equal(judgeConvertOutput('success\r\nWord instance is kept since its documents can not be checked\r\n').ok, true)
            assert.strict.equal(judgeConvertOutput('\r\n\r\n  success  \r\nanother line\r\n').ok, true)
        })

        it('counts only lines that start with "error:" (a note that merely mentions "error:" does not fail a success)', function() {
            assert.strict.equal(judgeConvertOutput('note: the previous error: was ignored\r\nsuccess\r\n').ok, true)
        })

    })

    describe('failure reported by the exe (the "error:" line, exit code 0)', function() {

        it('judges the exe output "error: ..." as failure with the code and the exe diagnosis (R01)', function() {
            let r = judgeConvertOutput(`error: (-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The file is corrupt.', 'wdmain11.chm', 25272, -2146822496), None)\r\n`)
            assert.strict.equal(r.ok, false)
            assert.strict.equal(r.msg, `${prefix}(-2147352567, 'Exception occurred.', (0, 'Microsoft Word', 'The file is corrupt.', 'wdmain11.chm', 25272, -2146822496), None)`)
        })

        it('judges the invalid-argv output of the exe as failure (R01)', function() {
            assert.strict.deepStrictEqual(judgeConvertOutput('error: invalid length of argv\r\n'), { ok: false, msg: `${prefix}invalid length of argv` })
        })

        it('gives a non-empty fixed reason when the "error:" line carries no text (R01)', function() {
            for (let s of ['error:', 'error:\r\n', 'error:   \r\n']) {
                let r = judgeConvertOutput(s)
                assert.strict.equal(r.ok, false, s)
                assert.strict.equal(r.msg.startsWith(prefix), true, s)
                assert.strict.equal(r.msg.length > prefix.length, true, s)
            }
        })

        it('fails closed: an "error:" line wins even when a "success" line is also present (R01, R03)', function() {
            assert.strict.equal(judgeConvertOutput('success\r\nerror: x\r\n').ok, false)
            assert.strict.equal(judgeConvertOutput('error: x\r\nsuccess\r\n').ok, false)
            assert.strict.equal(judgeConvertOutput('success\r\nerror: x\r\n').msg, `${prefix}x`)
        })

        it('takes the first "error:" line as the reason when there are several', function() {
            assert.strict.equal(judgeConvertOutput('error: first\r\nerror: second\r\n').msg, `${prefix}first`)
        })

        it('still reports the error when the stderr note is merged before the error line', function() {
            assert.strict.deepStrictEqual(judgeConvertOutput('Word instance is kept since 1 document(s) opened by another client\r\nerror: x\r\n'), { ok: false, msg: `${prefix}x` })
        })

    })

    describe('fails closed on anything that is not the success output (R03)', function() {

        let msgNo = `${prefix}docx2html.exe returned no success result`

        it('judges empty or blank output as failure', function() {
            for (let s of ['', ' ', '\r\n', '\n\n']) {
                assert.strict.deepStrictEqual(judgeConvertOutput(s), { ok: false, msg: msgNo }, JSON.stringify(s))
            }
        })

        it('judges unknown output as failure and does not echo it (R04)', function() {
            let s = 'Traceback: C:\\secret\\input.docx failed'
            let r = judgeConvertOutput(s)
            assert.strict.deepStrictEqual(r, { ok: false, msg: msgNo })
            assert.strict.equal(r.msg.includes('secret'), false)
        })

        it('requires the whole line to be "success" (lookalikes are not success)', function() {
            for (let s of ['successful', 'not success', 'success: no', 'Success', 'SUCCESS', 'succes', 'xsuccess']) {
                assert.strict.equal(judgeConvertOutput(s).ok, false, s)
            }
        })

        it('judges non-string input as failure without throwing (R05)', function() {
            let vs = [undefined, null, 0, 1, true, false, {}, [], ['success'], () => 'success', Symbol('x'), NaN]
            for (let v of vs) {
                let r = null
                assert.strict.doesNotThrow(() => {
                    r = judgeConvertOutput(v)
                })
                assert.strict.deepStrictEqual(r, { ok: false, msg: msgNo }, String(typeof v))
            }
        })

    })

})


//判成功後另驗輸出檔(R-CONVERT-RESULT-01 規則4): exe報success仍須fpOut存在且大小>0, 否則同碼、固定說明(不回顯路徑)
describe('judgeConvertFile', function() {

    let prefix = '[E_CONVERT_FAILED] '
    let msgNo = `${prefix}output file was not generated or is empty`
    let fdt = path.resolve('./test/_tmp/judgeConvertFile')

    before(function() {
        fs.rmSync(fdt, { recursive: true, force: true })
        fs.mkdirSync(path.join(fdt, 'asdir.html'), { recursive: true })
        fs.writeFileSync(path.join(fdt, 'ok.html'), '<html></html>')
        fs.writeFileSync(path.join(fdt, 'one.html'), 'x')
        fs.writeFileSync(path.join(fdt, 'empty.html'), '')
    })

    after(function() {
        fs.rmSync(fdt, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    })

    it('judges an existing non-empty file as ok (a size of 1 byte is enough)', function() {
        assert.strict.deepStrictEqual(judgeConvertFile(path.join(fdt, 'ok.html')), { ok: true, msg: '' })
        assert.strict.deepStrictEqual(judgeConvertFile(path.join(fdt, 'one.html')), { ok: true, msg: '' })
    })

    it('judges a missing file as failure with the fixed reason and without echoing the path', function() {
        let fp = path.join(fdt, 'missing-secret-name.html')
        let r = judgeConvertFile(fp)
        assert.strict.deepStrictEqual(r, { ok: false, msg: msgNo })
        assert.strict.equal(r.msg.includes('secret'), false)
        assert.strict.equal(r.msg.includes(fdt), false)
    })

    it('judges an empty file as failure', function() {
        assert.strict.deepStrictEqual(judgeConvertFile(path.join(fdt, 'empty.html')), { ok: false, msg: msgNo })
    })

    it('judges a folder at the output path as failure', function() {
        assert.strict.deepStrictEqual(judgeConvertFile(path.join(fdt, 'asdir.html')), { ok: false, msg: msgNo })
    })

    it('judges a non-string or blank path as failure without throwing', function() {
        let vs = [undefined, null, '', 0, {}, []]
        for (let v of vs) {
            let r = null
            assert.strict.doesNotThrow(() => {
                r = judgeConvertFile(v)
            })
            assert.strict.deepStrictEqual(r, { ok: false, msg: msgNo }, String(typeof v))
        }
    })

})
