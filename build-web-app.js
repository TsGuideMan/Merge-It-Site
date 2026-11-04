'use strict';

const fs = require('node:fs/promises');
const path = require('node:path');
const JSZip = require('./app/vendor/jszip.min.js');

async function main() {
  const root = __dirname;
  const files = [
    'app/index.html',
    'app/app.css',
    'app/app.js',
    'app/merge-engine.js',
    'app/ads-config.js',
    'app/ads.js',
    'app/vendor/LICENSE-JSZIP.txt',
    'app/vendor/LICENSE-XLSX.txt',
    'app/vendor/jszip.min.js',
    'app/vendor/xlsx.full.min.js',
    'pics/marge-it-72.png'
  ];
  const archive = new JSZip();
  archive.file('README.txt', [
    'Merge It - Offline Browser App',
    '',
    '1. Extract this ZIP file.',
    '2. Open app/index.html in a modern web browser.',
    '3. Choose a spreadsheet and a .pptx template, then create your presentations.',
    '',
    'Your files are processed locally in your browser and are not uploaded.',
    'Ads are disabled in this offline file:// copy; ads only load on the hosted website when valid AdSense IDs are configured.',
    'For the Windows desktop app and help, visit https://tsguideman.github.io/Merge-It-Site/'
  ].join('\r\n'));

  for (const relativePath of files) {
    const absolutePath = path.join(root, relativePath);
    archive.file(relativePath.replace(/\\/g, '/'), await fs.readFile(absolutePath));
  }

  const outputPath = path.join(root, 'downloads', 'MergeIt-WebApp.zip');
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const output = await archive.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
  await fs.writeFile(outputPath, output);
  console.log(`Created ${outputPath}`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
