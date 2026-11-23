const {chromium}=require('playwright');const fs=require('fs');
(async()=>{const mode=process.argv[2];const b=await chromium.launch({args:['--allow-file-access-from-files']});
const p=await b.newPage({viewport:{width:1080,height:1080},deviceScaleFactor:1});
await p.goto('file://'+__dirname+'/video.html');await p.evaluate(()=>document.fonts.ready);await p.waitForTimeout(500);
const st=await p.$('#stage');
if(mode==='preview'){for(const t of [2.2,4.6,6.5,8.2,10.8,12.9,14.9]){await p.evaluate(t=>render(t),t);await st.screenshot({path:`vprev_${t}.png`});}}
else{fs.mkdirSync('frames',{recursive:true});const N=450;for(let i=0;i<N;i++){await p.evaluate(t=>render(t),i/30);await st.screenshot({path:`frames/f${String(i).padStart(4,'0')}.png`});}}
await b.close();})();
