import w from 'wsemi'
import WDocx2html from './src/WDocx2html.mjs'
//import WDocx2html from 'w-docx2html/src/WDocx2html.mjs'
//import WDocx2html from 'w-docx2html'


async function test() {

    let fpIn = `./test/docin.docx`
    let fpOut = `./test/docout.html`
    let opt = {
        execFontGrow: 1,
    }

    let r = await WDocx2html(fpIn, fpOut, opt)
    console.log(r)
    // => ok

    w.fsDeleteFile(fpOut)

}
test()
    .catch((err) => {
        console.log('catch', err)
    })


//node g.mjs
