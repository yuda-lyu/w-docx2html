import rollupFiles from 'w-package-tools/src/rollupFiles.mjs'
import getFiles from 'w-package-tools/src/getFiles.mjs'


let fdSrc = './src'
let fdTar = './dist'


rollupFiles({
    fns: 'WDocx2html.mjs',
    fdSrc,
    fdTar,
    hookNameDist: () => 'w-docx2html',
    // nameDistType: 'kebabCase', //直接由hookNameDist給予
    globals: {
        'path': 'path',
        'fs': 'fs',
        'process': 'process',
        'child_process': 'child_process',
        'url': 'url',
        'crypto': 'crypto',
        'stream': 'stream',
        'stream/promises': 'stream/promises',
    },
    external: [
        'path',
        'fs',
        'process',
        'child_process',
        'url',
        'crypto',
        'stream',
        'stream/promises',
    ],
})

