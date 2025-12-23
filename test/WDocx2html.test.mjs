import fs from 'fs'
import w from 'wsemi'
import assert from 'assert'
import WDocx2html from '../src/WDocx2html.mjs'


function isWindows() {
    return process.platform === 'win32'
}


describe('WDocx2html', function() {

    //check
    if (!isWindows()) {
        return
    }

    let fpOutTrue = `./test/docoutTrue.html`

    let fpIn = `./test/docin.docx`
    let fpOut = `./test/docout.html`
    let opt = {
        execFontGrow: 1,
    }

    it('convert', async function() {
        await WDocx2html(fpIn, fpOut, opt)
        let r = fs.readFileSync(fpOut, 'utf8')
        let rr = fs.readFileSync(fpOutTrue, 'utf8')
        let b = r === rr
        w.fsDeleteFile(fpOut)
        assert.strict.deepEqual(true, b)
    })

})
