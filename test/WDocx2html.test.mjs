import fs from 'fs'
import path from 'path'
import { spawnSync } from 'child_process'
import w from 'wsemi'
import assert from 'assert'
import WDocx2html from '../src/WDocx2html.mjs'
import { setBinEnv } from './tools/binEnv.mjs'
import { genDocx } from './tools/docxFixture.mjs'
import { sleep, wordPids, isAutomationWord, waitWordGone, killNewAutomationWords, clearDisabledItems, startHeldWord, startBorrower, startVisibleWordWatch } from './tools/wordInstance.mjs'


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


describe('WDocx2html Word instance isolation', function() {

    //check
    if (!isWindows()) {
        return
    }

    //規格: 主控排程第408項(dh408, 2026-11-10, 源自下游w-md2docx之調查; 範本為w-html2docx之hd408): 轉檔器每次另開獨立之Word實例(DispatchEx), 只結束本行程所開者, 輸入檔以唯讀開啟
    //  ①他行程已開之Word(含其未存內容)不被借用, 不被結束; ②兩行程並行轉檔(甲進行中乙完整轉完、同一輸入檔同時與錯開1.5秒)皆成功;
    //  ③例外路徑(輸出路徑無效, 輸入檔不存在)亦只結束自己之實例且不殘留WINWORD; ④本行程之實例被他行程(舊寫法Dispatch)連上並開文件時, 轉完不結束該實例;
    //  ⑤輸出路徑即輸入檔時照改前行為(同一錯誤, 輸入檔不變)
    //  輸入為docx, 無http資源可扣住轉檔: 以大量段落之docx配合execFontGrow次數使轉檔持續數秒至十餘秒, 並以可觀察條件同步(轉檔器之WINWORD出現; 借用者見到實例內文件數達1=輸入檔已開啟)
    //  他行程之Word以PowerShell經COM扮演(見tools/wordInstance.mjs), 一律隱藏; 本組須在無其他Word時執行(見before)

    let fdTmp = path.resolve('./test/_tmp/WDocx2html-word')
    let fdOut = path.join(fdTmp, 'out')

    //nParas, 重型docx之段落數; growLong, 其轉檔使用之execFontGrow次數(使轉檔持續十餘秒); growMid, 並行案使用者
    let nParas = 5000
    let growLong = 60
    let growMid = 40

    //pre, 本組開始時已存在之WINWORD(須為空, 見before); null為本組未執行(略過或初始化前失敗), 此時after不結束任何WINWORD
    let pre = null

    //restoreBinEnv, docx2html.exe自本機鏡像(取自repo內bin/)取得, 不連外網
    let restoreBinEnv = null

    //watch, 可見Word視窗之監看(規格: 不得開啟可見之Word視窗)
    let watch = null

    //prog, docx2html.exe之路徑(直接呼叫用)
    let prog = null

    let fpStopOf = (name) => {
        return path.join(fdTmp, `${name}.stop`)
    }

    //genHeavy, 首段為單詞標記marker, 其後nParas個填充段落之docx
    let genHeavy = (name, marker) => {
        let texts = [marker]
        for (let i = 0; i < nParas; i++) {
            texts.push(`filler paragraph ${i} lorem ipsum dolor sit amet consectetur adipiscing elit`)
        }
        return genDocx(path.join(fdTmp, `${name}.docx`), texts)
    }

    //startConv, 以WDocx2html轉檔(本套件之對外API; 轉換失敗時其亦回ok, 故以輸出檔是否產生判定), 立即回傳狀態物件c, c.p於轉檔結束時resolve為c, 不拋錯
    let startConv = (fpIn, name, grow = 0) => {
        let fpOut = path.join(fdOut, `${name}.html`)
        let c = { ok: false, err: '', fpOut, html: '', done: false, p: null }
        c.p = WDocx2html(fpIn, fpOut, { execFontGrow: grow })
            .then((v) => {
                c.ok = v === 'ok'
            })
            .catch((e) => {
                c.err = String(e)
            })
            .then(() => {
                if (fs.existsSync(fpOut)) {
                    c.html = fs.readFileSync(fpOut, 'utf8').replace(/\s+/g, ' ')
                }
                c.ok = c.ok && c.html !== ''
                c.done = true
                return c
            })
        return c
    }

    //runExe, 直接呼叫轉檔器(略過WDocx2html), 回傳{ status, out, err }; 轉檔器以標準輸出回報success或error: ...
    let runExe = (inp) => {
        let r = spawnSync(prog, [Buffer.from(JSON.stringify(inp), 'utf8').toString('base64')], { encoding: 'latin1', windowsHide: true, timeout: 120000 })
        return { status: r.status, out: String(r.stdout).trim(), err: String(r.stderr).trim() }
    }

    //newPids, 本組開始後新出現之WINWORD
    let newPids = () => {
        return wordPids().filter((v) => !pre.includes(v))
    }

    //untilNewWord, 等(except以外之)新WINWORD出現(即轉檔器所開之實例), 轉檔先結束則回傳false
    let untilNewWord = async (c, except = [], ms = 60000) => {
        let t0 = Date.now()
        while (Date.now() - t0 < ms) {
            if (newPids().filter((v) => !except.includes(v)).length > 0) {
                return true
            }
            if (c.done) {
                return false
            }
            await sleep(100)
        }
        return false
    }

    //assertHeldIntact, H於轉檔後仍在, 其後連續2次讀取皆為文件數1且未存文字不變, 自始無讀取失敗
    let assertHeldIntact = async (h) => {
        let i = h.lines.length
        for (let k = 0; k < 2; k++) {
            let { i: j, line } = await h.waitLine(/^(DOCS|ERR)/, i)
            assert.strict.deepEqual(`DOCS 1 ${h.text}`, line)
            i = j + 1
        }
        assert.strict.deepEqual([], h.lines.filter((v) => /^(ERR|STDERR|SPAWNERR)/.test(v)))
        assert.strict.deepEqual(true, wordPids().includes(h.pid), `held WINWORD ${h.pid}`)
    }

    //assertNoVisible, 規格: 自監看第iw行起無可見之Word視窗(等監看再完成一輪後判定)
    let assertNoVisible = async (iw) => {
        assert.strict.deepEqual([], await watch.visibleSince(iw))
    }

    before(async function() {
        this.timeout(120000)

        //前一組之轉檔器所開之Word可能仍在結束中
        await waitWordGone(wordPids().filter(isAutomationWord), 30000)

        //有其他Word(含使用者開啟者)時不跑本組: 修前(Dispatch)之轉檔器會借用並以不儲存結束之; ④之他行程以GetActiveObject連上時亦無法確定連上本行程之實例
        let pids = wordPids()
        if (pids.length > 0) {
            console.log(`skip [Word instance isolation]: other WINWORD is running(${pids.join(',')})`)
            this.skip()
        }
        pre = pids

        restoreBinEnv = setBinEnv('WDocx2html-word-bin')

        //並行轉檔前先取得轉檔器至快取, 免兩個轉檔同時寫快取
        prog = (await WDocx2html.ensureBins({ names: 'docx2html.exe' }))['docx2html.exe']

        fs.rmSync(fdTmp, { recursive: true, force: true })
        fs.mkdirSync(fdOut, { recursive: true })

        //先前被強制結束之Word所留下、指向本套件檔案之停用項目, 會使開檔跳出可見訊息框而停住, 先清除
        let rd = clearDisabledItems(path.resolve('.'))
        if (rd.removed.length > 0 || rd.failed.length > 0) {
            console.log(`Word disabled items under this package: ${JSON.stringify(rd)}`)
        }

        watch = await startVisibleWordWatch(fpStopOf('watch'))
    })

    after(async function() {
        this.timeout(60000)

        //nKilledByWatch, 監看結束之可見視窗所屬Word數(被強制結束者亦會留下停用項目)
        let nKilledByWatch = 0
        if (watch) {
            await watch.stop()
            nKilledByWatch = watch.lines.filter((v) => /^KILLED/.test(v)).length
        }

        //收尾: 只結束本組新出現且經COM自動化啟動之WINWORD(各案已斷言不殘留, 此處為失敗時之清理); 有強制結束者(含監看所結束者)再清其可能留下之停用項目
        if (pre !== null) {
            let rk = await killNewAutomationWords(pre)
            if (rk.killed.length > 0 || nKilledByWatch > 0) {
                let rd = clearDisabledItems(path.resolve('.'))
                console.log(`cleanup: killed WINWORD ${JSON.stringify(rk)}, Word disabled items ${JSON.stringify(rd)}`)
            }
        }
        if (restoreBinEnv) {
            restoreBinEnv()
        }
        fs.rmSync(fdTmp, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 })
        try {
            fs.rmdirSync(path.resolve('./test/_tmp')) //尚有其他測試之暫存時非空, 略過
        }
        catch (err) {}
    })

    it('①他行程已開之Word(含未存文件)不被借用亦不被結束: 轉檔期間另起WINWORD, 轉完原實例仍在且文件數與未存文字不變, 轉檔器之WINWORD不殘留', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let h = await startHeldWord(fpStopOf('held1'))
        try {
            let c = startConv(genHeavy('held1', 'MarkHeld1'), 'held1', growLong)

            //規格①: 轉檔進行中有H以外之WINWORD, 即轉檔器另起之實例(未借用H)
            assert.strict.deepEqual(true, await untilNewWord(c, [h.pid]), 'conversion ended before another WINWORD appeared')
            let during = newPids()
            assert.strict.deepEqual(true, during.includes(h.pid) && during.filter((v) => v !== h.pid).length >= 1, JSON.stringify({ during, held: h.pid }))
            let mine = during.filter((v) => v !== h.pid)
            let r = await c.p

            //規格①: 轉檔成功, 輸出含來源內容
            assert.strict.deepEqual(true, r.ok, r.err)
            assert.strict.deepEqual(true, r.html.includes('MarkHeld1'))

            //規格①: 轉完H仍在, 文件數與未存文字不變
            await assertHeldIntact(h)

            //規格③: 轉檔器所開之WINWORD轉完即結束
            assert.strict.deepEqual([], await waitWordGone(mine))
        }
        finally {
            await h.stop()
        }
        assert.strict.deepEqual([], await waitWordGone(newPids()))

        //規格: 不得開啟可見之Word視窗
        await assertNoVisible(iw)
    })

    it('③例外路徑亦只結束自己之實例: 輸出路徑無效(文件已開啟後失敗)、輸入檔不存在皆回報error且不產生輸出檔, 他行程之Word仍在且文件不變, 轉檔器之WINWORD不殘留', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let h = await startHeldWord(fpStopOf('held3'))
        try {
            let fpOk = genDocx(path.join(fdTmp, 'ok.docx'), ['MarkFail'])

            //(a)輸出路徑無效: 輸入檔已開啟後SaveAs2失敗, (b)輸入檔不存在: 開啟失敗; 轉檔器以標準輸出回報error: ...(開啟失敗之訊息含該檔名)
            let cases = [
                ['saveas', { fpIn: fpOk, fpOut: path.join(fdTmp, 'no-such-dir', 'o.html'), execFontGrow: 0 }, ''],
                ['open', { fpIn: path.join(fdTmp, 'no-such.docx'), fpOut: path.join(fdOut, 'o.html'), execFontGrow: 0 }, 'no-such.docx'],
            ]
            for (let [name, inp, fn] of cases) {
                let rb = runExe(inp)
                assert.strict.deepEqual(true, rb.out.startsWith('error: ') && rb.out.includes(fn), `${name}: ${rb.status} ${rb.out} ${rb.err}`)
                assert.strict.deepEqual(false, fs.existsSync(inp.fpOut), name)

                //規格③: 每一例外之後他行程之Word皆仍在且文件不變
                await assertHeldIntact(h)
            }

            //規格③: 轉檔器所開之WINWORD皆已結束
            assert.strict.deepEqual([], await waitWordGone(newPids().filter((v) => v !== h.pid)))
        }
        finally {
            await h.stop()
        }
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

    it('②兩行程交錯轉檔: 甲轉檔進行中時乙完整轉完一份, 兩者皆成功且輸出完整, 不殘留WINWORD', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let a = startConv(genHeavy('interleaveA', 'MarkInterleaveA'), 'interleaveA', growLong)

        //甲之Word已啟動(轉檔進行中)
        assert.strict.deepEqual(true, await untilNewWord(a), 'conversion A ended before its WINWORD appeared')
        let b = startConv(genDocx(path.join(fdTmp, 'interleaveB.docx'), ['MarkInterleaveB']), 'interleaveB', 0)
        await b.p

        //乙轉完時甲仍在進行(兩者確實交錯)
        assert.strict.deepEqual(false, a.done, 'conversion A ended before B did')
        await a.p

        //規格②: 兩者皆成功, 各自輸出含其來源內容且不含對方之內容
        assert.strict.deepEqual([true, true], [a.ok, b.ok], JSON.stringify({ A: a.err, B: b.err }))
        assert.strict.deepEqual([true, false], [a.html.includes('MarkInterleaveA'), a.html.includes('MarkInterleaveB')])
        assert.strict.deepEqual([true, false], [b.html.includes('MarkInterleaveB'), b.html.includes('MarkInterleaveA')])

        //規格③: 不殘留WINWORD
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

    it('②兩行程並行轉檔同一輸入檔(同時開始、錯開1.5秒, 各2次): 皆成功且輸出完整, 每次結束後不殘留WINWORD', async function() {
        this.timeout(900000)
        let iw = watch.lines.length
        let fpIn = genHeavy('race', 'MarkRace')
        let rs = []
        for (let rep = 0; rep < 2; rep++) {
            for (let off of [0, 1500]) {
                let name = `race-${off}-${rep}`

                //maxWord, 轉檔期間同時存在之轉檔器WINWORD數(取樣觀察, 用以確認兩個實例確有重疊)
                let maxWord = 0
                let smp = { on: true }
                let sampler = (async () => {
                    while (smp.on) {
                        maxWord = Math.max(maxWord, newPids().length)
                        await sleep(300)
                    }
                })()

                //off為情境參數(乙晚甲off毫秒開始), 非同步點
                let [a, b] = await Promise.all([
                    startConv(fpIn, `${name}-a`, growMid).p,
                    sleep(off).then(() => startConv(fpIn, `${name}-b`, growMid).p),
                ])
                smp.on = false
                await sampler
                let full = (c) => c.ok && c.html.includes('MarkRace')
                let left = await waitWordGone(newPids())
                rs.push({ name, a: full(a) || a.err, b: full(b) || b.err, maxWord, left })
            }
        }

        //規格②: 0失敗, 兩轉檔各自一個實例且期間重疊(同一輸入檔, 輸入檔以唯讀開啟故不跳「檔案使用中」); 規格③: 每次結束後不殘留WINWORD
        assert.strict.deepEqual([], rs.filter((v) => v.a !== true || v.b !== true || v.maxWord < 2 || v.left.length > 0), JSON.stringify(rs))
        await assertNoVisible(iw)
    })

    it('④本行程之Word被他行程(舊寫法Dispatch)連上並開文件時: 轉檔成功, 轉完不結束該實例, 對方之文件保留, 對方用畢結束後不殘留', async function() {
        this.timeout(180000)
        let iw = watch.lines.length
        let b = null
        let x = []
        let c = startConv(genHeavy('borrowed', 'MarkBorrowed'), 'borrowed', growLong)
        try {

            //x, 轉檔器之實例(本組開始時無其他Word, 故此時僅此一個)
            assert.strict.deepEqual(true, await untilNewWord(c), 'conversion ended before its WINWORD appeared')
            x = newPids()
            assert.strict.deepEqual(1, x.length, JSON.stringify(x))

            //他行程以GetActiveObject連上(轉檔器之Word啟動需時, 有界重試), 待實例內文件數達1(輸入檔已開啟)後新增文件、寫入未存文字
            b = startBorrower(fpStopOf('borrower'), { text: 'unsaved-text-of-the-borrowing-client-408', attachMs: 30000, waitDocs: 1 })
            let { line } = await b.waitLine(/^(ADDED|NOATTACH|VISIBLE|EXIT)/)
            assert.strict.deepEqual('ADDED', line, b.lines.join(' / '))

            //對方新增文件時轉檔仍在進行(其後轉檔器收尾才盤點實例內他人之文件)
            assert.strict.deepEqual(false, c.done, 'conversion ended before the borrower added its document')
            await c.p

            //規格④: 轉檔成功, 輸出不含對方文件之文字
            assert.strict.deepEqual(true, c.ok, c.err)
            assert.strict.deepEqual([true, false], [c.html.includes('MarkBorrowed'), c.html.includes(b.text)])

            //規格④: 對方於轉完後仍持有其同一份文件(文件數1, 以其文件參考讀得其未存文字), 即實例與該文件皆未被轉檔器關閉
            let { line: lt } = await b.waitLine(/^(TEXTSET|SETFAIL|EXIT)/)
            assert.strict.deepEqual('TEXTSET', lt, b.lines.join(' / '))
            await b.stop()
            assert.strict.deepEqual(true, b.lines.includes(`DOCS 1 ${b.text}`), b.lines.join(' / '))
            b = null
        }
        finally {
            if (b) {
                await b.stop()
            }
            await c.p
        }

        //對方用畢以Quit結束該實例, 不殘留
        assert.strict.deepEqual([], await waitWordGone(x))
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

    it('⑤輸出路徑即輸入檔時照改前行為: 回報與改前相同之error(儲存與開啟中之檔同名), 輸入檔不變, 不殘留WINWORD', async function() {
        this.timeout(180000)
        let iw = watch.lines.length

        //規格(不得改變轉檔之輸出與對外API): 改前fpOut與fpIn同檔時Word回報error(wdmain11.chm 25774, -2146821994), 輸入檔不變; 此為改前行為之保留測試(修前修後皆綠)
        //  轉檔器於此情境仍以可寫開啟(isSamePath例外, 同範本w-html2docx之同檔例外); 於Word 16.0唯讀開啟亦得同一error, 故本案無法區分兩者(變異體「一律唯讀」存活, 屬等價變異)
        let fp = path.join(fdTmp, 'same-path.docx')
        fs.copyFileSync(path.resolve('./test/docin.docx'), fp)
        let bytes = fs.readFileSync(fp)
        let rb = runExe({ fpIn: fp, fpOut: fp, execFontGrow: 0 })
        assert.strict.deepEqual(true, rb.out.startsWith('error: ') && rb.out.includes('-2146821994'), `${rb.status} ${rb.out} ${rb.err}`)
        assert.strict.deepEqual(true, bytes.equals(fs.readFileSync(fp)))
        assert.strict.deepEqual([], await waitWordGone(newPids()))
        await assertNoVisible(iw)
    })

})
