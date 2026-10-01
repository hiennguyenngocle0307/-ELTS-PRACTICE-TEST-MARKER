// usage: node html2pdf.js <htmlRoot> <outRoot>   (htmlRoot/Midterm_N/<sub>/*.html -> outRoot/Midterm_N/<sub>/*.pdf)
const {chromium}=require('playwright');const fs=require('fs');const path=require('path');
const [,, src, dst]=process.argv;
function walk(d){return fs.readdirSync(d,{withFileTypes:true}).flatMap(e=>e.isDirectory()?walk(path.join(d,e.name)):[path.join(d,e.name)]);}
(async()=>{const b=await chromium.launch();const pg=await b.newPage();
for(const f of walk(src).filter(f=>f.endsWith('.html'))){
 const out=path.join(dst,path.relative(src,f)).replace(/\.html$/,'.pdf');fs.mkdirSync(path.dirname(out),{recursive:true});
 await pg.goto('file://'+f);await pg.pdf({path:out,format:'A4',printBackground:true,preferCSSPageSize:true});}
await b.close();})();
