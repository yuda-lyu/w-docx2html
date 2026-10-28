import fs from 'fs'
import os from 'os'
import path from 'path'
import http from 'http'
import assert from 'assert'
import { spawnSync } from 'child_process'
import ensureBins from '../src/ensureBins.mjs'
import autoDownloadFiles from '../src/autoDownloadFiles.mjs'
import WDocx2html from '../src/WDocx2html.mjs'


//規格: 外部執行檔之存放下載校驗與離線(制式規格) D1~D4、O1~O3、§8 T3~T7(T8為主幹模式專屬, 本套件為釘選模式故不適用)
//全部以本機HTTP伺服器(扮GitHub)與本地資料夾(扮鏡像與快取)進行, 不連外網
describe(`ensureBins`, function() {
    this.timeout(180000)

    let port = 8951
    let RAW = 'https://raw.githubusercontent.com'
    let fdt = path.resolve('./test/_tmp/ensureBins')
    let fpSrc = path.resolve('bin/docx2html.exe')
    let bufGood = fs.readFileSync(fpSrc)
    let manifest = JSON.parse(fs.readFileSync('bin/manifest.json', 'utf8'))
    let entry = manifest.files[0]
    let name = entry.name
    let sha256 = entry.sha256
    let pkgName = manifest.pkg
    let urlPinned = `${RAW}/${manifest.repo}/${manifest.commit}/${entry.path}`

    //flip, 回傳同長度但內容不同之副本
    let flip = (buf) => {
        let b = Buffer.from(buf)
        b[1000] = b[1000] ^ 0xff
        return b
    }

    //伺服器狀態: mode為good、tampered(同長度不同內容)、long(多1位元組)、missing(404)
    let mode = 'good'
    let hits = []
    let server = null

    //fetch, 把GitHub原始檔網址改指本機伺服器, 並記錄被請求之原網址
    let urls = []
    let fetch = (url, init) => {
        urls.push(url)
        return globalThis.fetch(url.replace(RAW, `http://127.0.0.1:${port}`), init)
    }

    //fetchBan, 任何網路請求皆視為錯誤(用於斷言未連網)
    let fetchBan = () => {
        throw new Error('network access is not allowed in this case')
    }

    let seq = 0
    let newDir = (tag) => {
        seq++
        return path.join(fdt, `${tag}-${seq}`)
    }

    //listFiles, 遞迴列出資料夾內所有檔案(不存在回傳空陣列)
    let listFiles = (fd) => {
        let r = []
        if (!fs.existsSync(fd)) {
            return r
        }
        for (let n of fs.readdirSync(fd)) {
            let fp = path.join(fd, n)
            if (fs.statSync(fp).isDirectory()) {
                r.push(...listFiles(fp))
            }
            else {
                r.push(fp)
            }
        }
        return r
    }

    //mkMirror, 建立本地資料夾鏡像(結構<鏡像>/<套件>/<sha256>/<name>)
    let mkMirror = (buf) => {
        let fdMirror = newDir('mirror')
        let fd = path.join(fdMirror, pkgName, sha256)
        fs.mkdirSync(fd, { recursive: true })
        fs.writeFileSync(path.join(fd, name), buf)
        return fdMirror
    }

    //withEnv, 暫時設定環境變數(值為undefined代表刪除), 結束後還原
    let withEnv = async(vars, fun) => {
        let old = {}
        for (let k of Object.keys(vars)) {
            old[k] = process.env[k]
            if (vars[k] === undefined) {
                delete process.env[k]
            }
            else {
                process.env[k] = vars[k]
            }
        }
        try {
            return await fun()
        }
        finally {
            for (let k of Object.keys(vars)) {
                if (old[k] === undefined) {
                    delete process.env[k]
                }
                else {
                    process.env[k] = old[k]
                }
            }
        }
    }

    let catchErr = async(fun) => {
        try {
            await fun()
        }
        catch (err) {
            return err
        }
        return null
    }

    before(async function() {
        fs.rmSync(fdt, { recursive: true, force: true })
        fs.mkdirSync(fdt, { recursive: true })
        server = http.createServer((req, res) => {
            hits.push(req.url)
            if (mode === 'missing') {
                res.writeHead(404)
                res.end()
                return
            }
            let body = bufGood
            if (mode === 'tampered') {
                body = flip(bufGood)
            }
            else if (mode === 'long') {
                body = Buffer.concat([bufGood, Buffer.from([0])])
            }
            res.writeHead(200, { 'content-length': String(body.length) })
            res.end(body)
        })
        await new Promise((resolve, reject) => {
            server.once('error', reject)
            server.listen(port, '127.0.0.1', resolve)
        })
    })

    after(async function() {
        await new Promise((resolve) => {
            server.close(resolve)
        })
        fs.rmSync(fdt, { recursive: true, force: true })
    })

    beforeEach(function() {
        mode = 'good'
        hits = []
        urls = []
    })

    it(`should download from the pinned commit, verify and cache; the second call hits the cache without network (T3, D1, D3, D4)`, async function() {
        let cacheDir = newDir('cache')

        let r1 = await ensureBins({ cacheDir, fetch, mirror: '', offline: false })
        assert.strict.deepStrictEqual(Object.keys(r1), [name])
        assert.strict.equal(r1[name], path.join(cacheDir, sha256, name)) //D3: <快取>/<sha256>/<name>
        assert.strict.equal(fs.readFileSync(r1[name]).equals(bufGood), true)
        assert.strict.deepStrictEqual(urls, [urlPinned]) //O2: 釘選之提交網址, 非master、非releases/latest
        assert.strict.equal(hits.length, 1)

        let r2 = await ensureBins({ cacheDir, fetch: fetchBan, mirror: '', offline: false })
        assert.strict.equal(r2[name], r1[name]) //D4: 命中且不連網
        assert.strict.equal(hits.length, 1)
    })

    it(`should reject a download whose bytes differ from the manifest and leave no file (T4, M2, D2)`, async function() {
        mode = 'tampered'
        let cacheDir = newDir('cache')
        let err = await catchErr(() => ensureBins({ cacheDir, fetch, mirror: '', offline: false, maxAttempts: 1 }))
        assert.strict.equal(err && err.code, 'E_HASH_MISMATCH', err && err.message)
        assert.strict.deepStrictEqual(listFiles(cacheDir), [])
    })

    it(`should reject a download whose size differs from the manifest and leave no file (T4, M2, D2)`, async function() {
        mode = 'long'
        let cacheDir = newDir('cache')
        let err = await catchErr(() => ensureBins({ cacheDir, fetch, mirror: '', offline: false, maxAttempts: 1 }))
        assert.strict.equal(err && err.code, 'E_SIZE_MISMATCH', err && err.message)
        assert.strict.deepStrictEqual(listFiles(cacheDir), [])
    })

    it(`should reject an HTTP error status without leaving a file (T4)`, async function() {
        mode = 'missing'
        let cacheDir = newDir('cache')
        let err = await catchErr(() => ensureBins({ cacheDir, fetch, mirror: '', offline: false, maxAttempts: 1 }))
        assert.strict.equal(err && err.code, 'E_HTTP', err && err.message)
        assert.strict.deepStrictEqual(listFiles(cacheDir), [])
    })

    it(`should get the file from a local folder mirror without network (T5, O1, O2)`, async function() {
        let fdMirror = mkMirror(bufGood)
        let cacheDir = newDir('cache')
        let r = await ensureBins({ cacheDir, fetch: fetchBan, mirror: fdMirror, offline: true })
        assert.strict.equal(r[name], path.join(cacheDir, sha256, name))
        assert.strict.equal(fs.readFileSync(r[name]).equals(bufGood), true)
        assert.strict.equal(hits.length, 0)
    })

    it(`should refuse a tampered mirror file: offline fails with E_OFFLINE_MISSING and caches nothing (T5, O2)`, async function() {
        let fdMirror = mkMirror(flip(bufGood))
        let cacheDir = newDir('cache')
        let err = await catchErr(() => ensureBins({ cacheDir, fetch: fetchBan, mirror: fdMirror, offline: true }))
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
        assert.strict.deepStrictEqual(listFiles(cacheDir), [])
    })

    it(`should fall back to GitHub when the mirror file is tampered and online, still verified (T5, O2)`, async function() {
        let fdMirror = mkMirror(flip(bufGood))
        let cacheDir = newDir('cache')
        let r = await ensureBins({ cacheDir, fetch, mirror: fdMirror, offline: false })
        assert.strict.equal(fs.readFileSync(r[name]).equals(bufGood), true)
        assert.strict.equal(hits.length, 1)
    })

    it(`should reject offline with E_OFFLINE_MISSING naming the file, sha256, size and the cache path when nothing is available (T6, O3)`, async function() {
        let cacheDir = newDir('cache')
        let err = await catchErr(() => ensureBins({ cacheDir, fetch: fetchBan, mirror: '', offline: true }))
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
        let cachePath = path.join(cacheDir, sha256, name)
        assert.strict.equal(err.message.includes(name), true, err.message)
        assert.strict.equal(err.message.includes(sha256), true, err.message)
        assert.strict.equal(err.message.includes(`${entry.size} bytes`), true, err.message)
        assert.strict.equal(err.message.includes(cachePath), true, err.message)
        assert.strict.equal(err.message.includes(`<W_BIN_MIRROR>/${pkgName}/${sha256}/${name}`), true, err.message)
        assert.strict.deepStrictEqual(err.missing.map((v) => [v.name, v.sha256, v.size, v.cachePath]), [[name, sha256, entry.size, cachePath]])
        assert.strict.equal(hits.length, 0)
    })

    it(`should re-obtain the file when the cached file was altered (T7, D4)`, async function() {
        let cacheDir = newDir('cache')
        let r1 = await ensureBins({ cacheDir, fetch, mirror: '', offline: false })
        assert.strict.equal(hits.length, 1)

        fs.writeFileSync(r1[name], flip(bufGood)) //同長度不同內容

        let r2 = await ensureBins({ cacheDir, fetch, mirror: '', offline: false })
        assert.strict.equal(r2[name], r1[name])
        assert.strict.equal(fs.readFileSync(r2[name]).equals(bufGood), true)
        assert.strict.equal(hits.length, 2)
    })

    it(`should re-obtain the file when the cached file was truncated (T7, D4)`, async function() {
        let cacheDir = newDir('cache')
        let r1 = await ensureBins({ cacheDir, fetch, mirror: '', offline: false })

        fs.writeFileSync(r1[name], bufGood.subarray(0, 100)) //大小不符

        let r2 = await ensureBins({ cacheDir, fetch, mirror: '', offline: false })
        assert.strict.equal(fs.readFileSync(r2[name]).equals(bufGood), true)
        assert.strict.equal(hits.length, 2)
    })

    it(`should read W_BIN_CACHE, W_BIN_MIRROR and W_BIN_OFFLINE (D3, O1, O3)`, async function() {
        let fdMirror = mkMirror(bufGood)
        let fdCache = newDir('envcache')
        let r = await withEnv({ W_BIN_CACHE: fdCache, W_BIN_MIRROR: fdMirror, W_BIN_OFFLINE: '1' }, () => {
            return ensureBins({ fetch: fetchBan })
        })
        //D3: 共用快取為<W_BIN_CACHE>/<套件>/<sha256>/<name>
        assert.strict.equal(r[name], path.join(fdCache, pkgName, sha256, name))
        assert.strict.equal(fs.readFileSync(r[name]).equals(bufGood), true)
        assert.strict.equal(hits.length, 0)

        //W_BIN_OFFLINE=1且無快取無鏡像: E_OFFLINE_MISSING, 不連網
        let fdCache2 = newDir('envcache')
        let err = await withEnv({ W_BIN_CACHE: fdCache2, W_BIN_MIRROR: undefined, W_BIN_OFFLINE: '1' }, () => {
            return catchErr(() => ensureBins({ fetch: fetchBan }))
        })
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
        assert.strict.equal(err.message.includes(path.join(fdCache2, pkgName, sha256, name)), true, err.message)
    })

    it(`should place the default cache under the package root even when the module is loaded from another working directory (D3)`, function() {
        //pkgRoot於模組載入時決定, 故以子行程於他處工作目錄載入(被安裝為相依套件時, 載入時之工作目錄為使用者專案而非套件根)
        let env = { ...process.env }
        for (let k of ['W_BIN_CACHE', 'W_BIN_MIRROR', 'W_BIN_OFFLINE']) {
            delete env[k]
        }
        let cwdOther = os.tmpdir()
        let r = spawnSync(process.execPath, [path.resolve('test/tools/ensureBinsChild.mjs'), path.resolve('src/ensureBins.mjs')], { cwd: cwdOther, env, encoding: 'utf8', windowsHide: true })
        assert.strict.equal(r.status, 0, r.stderr)
        let o = JSON.parse(r.stdout.trim().split('\n').pop())
        assert.strict.equal(path.resolve(o.cwd), path.resolve(cwdOther)) //確實於他處載入
        assert.strict.equal(o.code, 'E_OFFLINE_MISSING')
        assert.strict.equal(o.cachePath, path.join(path.resolve('.'), 'bin-cache', sha256, name))
    })

    it(`should not forward floatingManifestUrl, which would change the trust anchor of the manifest (E2)`, async function() {
        //floatingManifestUrl只在主幹模式生效, 但被轉傳時無效值會使fsEnsureBins以E_ARG拒絕; 套件端丟棄它故仍正常取得
        let fdMirror = mkMirror(bufGood)
        let cacheDir = newDir('cache')
        let r = await ensureBins({ cacheDir, fetch: fetchBan, mirror: fdMirror, offline: true, floatingManifestUrl: 'not-a-valid-url' })
        assert.strict.equal(fs.readFileSync(r[name]).equals(bufGood), true)
        assert.strict.equal(hits.length, 0)
    })

    it(`should export ensureBins from the main file WDocx2html.ensureBins (D1, E2)`, function() {
        assert.strict.equal(WDocx2html.ensureBins, ensureBins)
    })

    it(`should keep autoDownloadFiles as a compatible layer that returns the verified path through ensureBins`, async function() {
        let fdMirror = mkMirror(bufGood)
        let fdCache = newDir('envcache')
        let r = await withEnv({ W_BIN_CACHE: fdCache, W_BIN_MIRROR: fdMirror, W_BIN_OFFLINE: '1' }, () => {
            return autoDownloadFiles()
        })
        assert.strict.deepStrictEqual(r, { fpExe: path.join(fdCache, pkgName, sha256, name) })
        assert.strict.equal(fs.readFileSync(r.fpExe).equals(bufGood), true)

        //無快取無鏡像之離線: 與ensureBins相同之E_OFFLINE_MISSING(不再回傳字串'can not find docx2html.exe')
        let err = await withEnv({ W_BIN_CACHE: newDir('envcache'), W_BIN_MIRROR: undefined, W_BIN_OFFLINE: '1' }, () => {
            return catchErr(() => autoDownloadFiles())
        })
        assert.strict.equal(err && err.code, 'E_OFFLINE_MISSING', err && err.message)
    })

})
