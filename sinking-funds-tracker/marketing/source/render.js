const {chromium}=require('playwright');
(async()=>{const scale=+process.argv[2]||1;const out=process.argv[3]||'prev';
const b=await chromium.launch({args:['--allow-file-access-from-files']});
const p=await b.newPage({viewport:{width:1100,height:1100},deviceScaleFactor:scale});
await p.goto('file://'+__dirname+'/mockups.html');await p.evaluate(()=>document.fonts.ready);await p.waitForTimeout(600);
const ids=await p.$$eval('.slide',e=>e.map(x=>x.id));
for(const id of ids){await (await p.$('#'+id)).screenshot({path:`${out}_${id}.png`});}
await b.close();})();
