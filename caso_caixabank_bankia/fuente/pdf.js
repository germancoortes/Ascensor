const { chromium } = require('playwright');
(async () => { const b = await chromium.launch(); const p = await b.newPage();
for (const [i, o] of [[process.argv[2], process.argv[3]], [process.argv[4], process.argv[5]]]) {
  await p.goto('file://' + i); await p.pdf({ path: o, format: 'A4', printBackground: true, preferCSSPageSize: true,
    displayHeaderFooter: true, headerTemplate: '<span></span>', footerTemplate: '<div style="font-size:8px;color:#5C6B7A;width:100%;text-align:center;font-family:Arial"><span class="pageNumber"></span></div>' }); }
await b.close(); })();
