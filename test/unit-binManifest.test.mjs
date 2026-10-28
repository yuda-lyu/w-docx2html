import fs from 'fs'
import assert from 'assert'
import { execSync } from 'child_process'
import fsGetBinManifestEntry from 'wsemi/src/fsGetBinManifestEntry.mjs'
import fsEnsureBins from 'wsemi/src/fsEnsureBins.mjs'


//規格: 外部執行檔之存放下載校驗與離線(制式規格) S1、M1、M2、M4、§8 T1、T2
describe(`bin/manifest.json`, function() {
    this.timeout(120000)

    let manifest = JSON.parse(fs.readFileSync('bin/manifest.json', 'utf8'))
    let fdt = './test/_tmp/binManifest'

    after(function() {
        fs.rmSync(fdt, { recursive: true, force: true })
    })

    it(`should pass the schema check of fsEnsureBins (M1, M2)`, async function() {
        //離線且快取為空資料夾: 清單有效時為E_OFFLINE_MISSING(缺檔), 清單無效時為E_MANIFEST; 不連網
        let err = null
        try {
            await fsEnsureBins('bin/manifest.json', { pkgRoot: '.', cacheDir: `${fdt}/cache`, mirror: '', offline: true })
        }
        catch (e) {
            err = e
        }
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
    })

    it(`should be a pinned manifest of this package whose commit holds bin/ (M1, M2, U1)`, function() {
        //M1欄位, M2釘選: mode為pinned, commit為40碼小寫十六進位, pkg與repo對應本套件
        let pkg = JSON.parse(fs.readFileSync('package.json', 'utf8'))
        assert.strict.equal(manifest.schema, 1)
        assert.strict.equal(manifest.pkg, pkg.name)
        assert.strict.equal(manifest.mode, 'pinned')
        assert.strict.equal(manifest.repo, `yuda-lyu/${pkg.name}`)
        assert.strict.equal(/^[0-9a-f]{40}$/.test(manifest.commit), true, manifest.commit)
        //S2: 清單所列之檔皆於bin/下, 且為自有檔(path)而非廠商檔(url)
        assert.strict.deepStrictEqual(manifest.files.map((v) => v.path), ['bin/docx2html.exe'])
        assert.strict.deepStrictEqual(manifest.files.map((v) => v.name), ['docx2html.exe'])
    })

    it(`should match the real files in bin/ (M4, T2)`, async function() {
        for (let f of manifest.files) {
            for (let p of [f, ...(f.parts || [])]) {
                if (typeof p.path === 'string') {
                    let e = await fsGetBinManifestEntry(p.path)
                    assert.strict.deepStrictEqual([e.sha256, e.size], [p.sha256, p.size], p.path)
                }
            }
        }
    })

    it(`should keep the legacy file src/docx2html.exe identical to bin/ (S2 舊版相容)`, async function() {
        //舊版於執行時自主幹src/docx2html.exe下載, 該檔須保留於repo且與bin/之檔位元相同
        let a = await fsGetBinManifestEntry('src/docx2html.exe')
        let b = await fsGetBinManifestEntry('bin/docx2html.exe')
        assert.strict.deepStrictEqual([a.sha256, a.size], [b.sha256, b.size])
    })

    it(`should not pack external executables into the npm package, only bin/manifest.json (S1, T1)`, function() {
        let r = JSON.parse(execSync('npm pack --dry-run --json', { encoding: 'utf8', windowsHide: true }))
        let info = Array.isArray(r) ? r[0] : Object.values(r)[0] //npm 11以前為陣列, npm 12起為以套件名為鍵之物件
        let fps = info.files.map((v) => v.path)
        assert.strict.deepStrictEqual(fps.filter((v) => /\.(exe|dll|msi|zip|7z|\d{3})$/i.test(v)), [])
        assert.strict.deepStrictEqual(fps.filter((v) => v.startsWith('bin/') || v.startsWith('bin-cache/')), ['bin/manifest.json'])
    })

})
