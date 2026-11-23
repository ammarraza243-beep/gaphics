const {chromium}=require('playwright');
(async()=>{const s=+process.argv[2]||1,out=process.argv[3]||'b';const b=await chromium.launch({args:['--allow-file-access-from-files']});
const p=await b.newPage({viewport:{width:3400,height:1000},deviceScaleFactor:s});
await p.goto('file://'+__dirname+'/brand.html');await p.evaluate(()=>document.fonts.ready);await p.waitForTimeout(500);await p.evaluate(()=>{document.body.style.background='transparent';document.documentElement.style.background='transparent'});
for(const id of ['logo','markonly','wordmark','banner','mini']){await (await p.$('#'+id)).screenshot({path:`${out}_${id}.png`,omitBackground:true});}
await b.close();})();
