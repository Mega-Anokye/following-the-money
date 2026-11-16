// Part of "Following the money" (MoMo + IC Securities reconciliation). Personal identifiers are placeholders:
// set STATEMENTS_DIR, OWN_ACCT and EMPLOYER_ACCT (environment variables) and the OWN_BANK_REF / HOLDER patterns before running on your own statements.
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import fs from 'fs';
const COLS=['TRANSACTION DATE','FROM ACCT','FROM NAME','FROM NO.','TRANS. TYPE','AMOUNT','FEES','E-LEVY','BAL BEFORE','BAL AFTER','TO NO.','TO NAME','TO ACCT','F_ID','REF','OVA'];
const oldRe=/^\d{2}-[A-Za-z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M$/;
// newer statements (from Jun 2026): "2026-09-18 18:55:36.870" on its own line above the row, "UTC" below it
const isoRe=/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(\.\d+)?$/;
const dateRe={test:s=>oldRe.test(s)||isoRe.test(s)};
const MON=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
// normalise the new format to the old "dd-Mon-yyyy hh:mm:ss AM" so every statement reads the same
function normDate(s){ const m=s.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})/); if(!m) return s;
  let h=+m[4]; const ap=h>=12?'PM':'AM'; h=h%12||12;
  return `${m[3]}-${MON[+m[2]-1]}-${m[1]} ${String(h).padStart(2,'0')}:${m[5]}:${m[6]} ${ap}`; }
export async function extract(file){
  const doc=await pdfjs.getDocument({data:new Uint8Array(fs.readFileSync(file)),verbosity:0}).promise;
  const out=[]; let meta={}; let prevHdr=null;
  for(let pn=1;pn<=doc.numPages;pn++){
    const items=(await (await doc.getPage(pn)).getTextContent()).items.filter(i=>i.str.trim())
      .map(i=>({s:i.str.trim(),x:i.transform[4],y:i.transform[5],c:i.transform[4]+i.width/2}));
    if(pn===1){const t=items.map(i=>i.s);meta.from=t[t.indexOf('From:')+1];meta.to=t[t.indexOf('To:')+1];}
    const td=items.find(i=>i.s===COLS[0]); let hdr=COLS.map(n=>td&&items.find(i=>i.s===n&&Math.abs(i.y-td.y)<2)); if(hdr.some(h=>!h)){ if(!prevHdr) throw new Error("no header"); hdr=prevHdr.map(h=>({...h,y:1e9})); } prevHdr=hdr;
    const hy=hdr[0].y;
    const body=items.filter(i=>i.y<hy-3 && !/^Powered by/.test(i.s));
    const anchors=body.filter(i=>dateRe.test(i.s)&&Math.abs(i.c-hdr[0].c)<40).sort((a,b)=>b.y-a.y);
    if(!anchors.length){ if(body.length) console.error("page w/o rows",file,pn,body.map(i=>i.s).join("|")); continue; }
    // the new format's data line sits ~3pt below its date line, so centre the row on the data line
    const rows=anchors.map(a=>({y:isoRe.test(a.s)?a.y-3:a.y,cells:COLS.map(()=>[])}));
    const byCol=COLS.map(()=>[]);
    for(const it of body) byCol[hdr.reduce((b,h,k)=>Math.abs(h.c-it.c)<Math.abs(hdr[b].c-it.c)?k:b,0)].push(it);
    byCol.forEach((list,ci)=>{
      list.sort((a,b)=>b.y-a.y); const blocks=[];
      for(const it of list){const B=blocks.at(-1); if(B&&B.at(-1).y-it.y<7.5&&!(ci===0)) B.push(it); else blocks.push([it]);}
      for(const B of blocks){const cy=(B[0].y+B.at(-1).y)/2;
        const r=rows.reduce((b,r)=>Math.abs(r.y-cy)<Math.abs(b.y-cy)?r:b);
        if(Math.abs(r.y-cy)>3) console.error('offcentre',file,pn,ci,B.map(i=>i.s).join(' '),r.y,cy);
        r.cells[ci].push(...B);}
    });
    for(const r of rows){ const v=r.cells.map(c=>c.sort((a,b)=>b.y-a.y||a.x-b.x).map(i=>i.s).join(' '));
      const d0=v[0].replace(/\s*UTC$/,''); v[0]=isoRe.test(d0)?normDate(d0):v[0];
      out.push({page:pn,v}); }
  }
  return {meta,rows:out,pages:doc.numPages};
}
