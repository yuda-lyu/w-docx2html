import fs from 'fs'
import path from 'path'
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

    //外部執行檔(規格 O1、O3): 不連外網, 以本機鏡像供檔, 快取放test/_tmp/
    let fdt = path.resolve('./test/_tmp/WDocx2html')
    let fdMirror = path.join(fdt, 'mirror')
    let fdCache = path.join(fdt, 'cache')
    let manifest = JSON.parse(fs.readFileSync('bin/manifest.json', 'utf8'))
    let entry = manifest.files[0]
    let envNames = ['W_BIN_CACHE', 'W_BIN_MIRROR', 'W_BIN_OFFLINE']
    let envOld = {}

    before(function() {
        fs.rmSync(fdt, { recursive: true, force: true })
        let fdSha = path.join(fdMirror, manifest.pkg, entry.sha256)
        fs.mkdirSync(fdSha, { recursive: true })
        fs.copyFileSync('bin/docx2html.exe', path.join(fdSha, entry.name))
        for (let k of envNames) {
            envOld[k] = process.env[k]
        }
        process.env.W_BIN_CACHE = fdCache
        process.env.W_BIN_MIRROR = fdMirror
        process.env.W_BIN_OFFLINE = '1'
    })

    after(function() {
        for (let k of envNames) {
            if (envOld[k] === undefined) {
                delete process.env[k]
            }
            else {
                process.env[k] = envOld[k]
            }
        }
        fs.rmSync(fdt, { recursive: true, force: true })
    })

    it('convert', async function() {
        await WDocx2html(fpIn, fpOut, opt)
        let r = fs.readFileSync(fpOut, 'utf8')
        let rr = fs.readFileSync(fpOutTrue, 'utf8')
        let b = r === rr
        w.fsDeleteFile(fpOut)
        assert.strict.deepEqual(true, b)

        //轉換所用之執行檔應已經由ensureBins自鏡像取得並放入快取<W_BIN_CACHE>/<套件>/<sha256>/<name>(規格 D3), 而非套件內src/docx2html.exe
        let fpCached = path.join(fdCache, manifest.pkg, entry.sha256, entry.name)
        assert.strict.equal(fs.existsSync(fpCached), true)
        assert.strict.equal(fs.statSync(fpCached).size, entry.size)
    })

})
