// Part of "Following the money" (MoMo + IC Securities reconciliation). Personal identifiers are placeholders:
// set STATEMENTS_DIR, OWN_ACCT and EMPLOYER_ACCT (environment variables) and the OWN_BANK_REF / HOLDER patterns before running on your own statements.
import {extract} from './extract.mjs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import fs from 'fs';
const D=(process.env.STATEMENTS_DIR || './statements/');
const OWN=(process.env.OWN_ACCT || 'YOUR_WALLET_ACCT');
const FILES=[4,1,2,3,5,6];          // oldest first
const NEW_FORMAT=new Set([5,6]);   // statements in MTN's newer layout (from Jun 2026)
const MON={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
const parse=s=>{const m=s.match(/^(\d\d)-(\w{3})-(\d{4}) (\d\d):(\d\d):(\d\d) ([AP])M$/);let h=+m[4]%12;if(m[7]==='P')h+=12;return new Date(Date.UTC(+m[3],MON[m[2]],+m[1],h,+m[5],+m[6]));};
const r2=x=>Math.round(x*100)/100;
const fail=[], notes=[];const ok=(c,msg)=>{if(!c){fail.push(msg);console.log('FAIL',msg);}};
// independent count of transaction dates in raw PDF text (both layouts)
async function rawDateCount(f){const doc=await pdfjs.getDocument({data:new Uint8Array(fs.readFileSync(f)),verbosity:0}).promise;let n=0;
  for(let p=1;p<=doc.numPages;p++){for(const i of (await (await doc.getPage(p)).getTextContent()).items){const s=i.str.trim();
    if(/^\d{2}-[A-Za-z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M$/.test(s)||/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(\.\d+)?$/.test(s))n++;}}return n;}
const stmts=[];let swaps=[];
for(const n of FILES){const f=D+`MomoStatementReport (${n}).pdf`;const {meta,rows,pages}=await extract(f);
  const tx=rows.map(r=>{const v=r.v;return{date:parse(v[0]),dateStr:v[0],fromAcct:v[1],fromName:v[2],fromNo:v[3],type:v[4],amount:+v[5],fees:+v[6],elevy:+v[7],bb:+v[8],ba:+v[9],toNo:v[10],toName:v[11],toAcct:v[12],fid:v[13],ref:v[14],ova:v[15],src:`MomoStatementReport (${n}).pdf`,fmt:NEW_FORMAT.has(n)?'new':'old'};}).reverse(); // oldest first
  tx.forEach(t=>{ok([t.amount,t.fees,t.elevy,t.bb,t.ba].every(Number.isFinite),`non-numeric value ${t.fid}`);t.dir=t.toAcct===OWN?'In':'Out';
    // newer layout sometimes prints an incoming row's balances the wrong way round: fix and record it
    if(t.dir==='In'&&t.bb>t.ba&&Math.abs((t.bb-t.ba)-t.amount)<0.005){swaps.push(`${t.dateStr} ${t.fid} ${t.amount}`);[t.bb,t.ba]=[t.ba,t.bb];t.balSwapped=true;}
    ok((t.dir==='In')===(t.ba>t.bb)||t.ba===t.bb,`direction vs balance ${t.fid} ${t.dateStr}`);});
  for(let i=1;i<tx.length;i++) ok(tx[i].date>=tx[i-1].date,`order ${n} ${tx[i].dateStr}`);
  const s={n,file:`MomoStatementReport (${n}).pdf`,fmt:NEW_FORMAT.has(n)?'new':'old',meta,pages,tx,raw:await rawDateCount(f)};
  s.inT=r2(tx.filter(t=>t.dir==='In').reduce((a,t)=>a+t.amount,0));s.outT=r2(tx.filter(t=>t.dir==='Out').reduce((a,t)=>a+t.amount,0));
  s.net=r2(tx.reduce((a,t)=>a+t.ba-t.bb,0));s.open=tx[0].bb;s.close=tx.at(-1).ba;s.first=tx[0].date;s.last=tx.at(-1).date;s.firstFid=tx[0].fid;s.lastFid=tx.at(-1).fid;
  s.breaks=0;for(let i=1;i<tx.length;i++) if(Math.abs(tx[i].bb-tx[i-1].ba)>0.005) s.breaks++;
  s.unlisted=r2(s.close-s.open-s.net);
  ok(s.raw===tx.length,`${s.file}: PDF has ${s.raw} dates but ${tx.length} rows extracted`);
  if(Math.abs(s.unlisted)>=0.005) notes.push(`${s.file}: opening + changes differs from closing by ${s.unlisted} across ${s.breaks} balance break(s)`);
  stmts.push(s);}
// names: the newer layout shows only first names, and on incoming rows FROM NAME is the account holder ("HOLDER").
// Learn full names per account number from the older statements and apply them.
const names=new Map();
for(const s of stmts.filter(s=>s.fmt==='old')) for(const t of s.tx){
  if(t.dir==='In'&&t.fromName&&t.fromAcct!==OWN) names.set(t.fromAcct,t.fromName);
  if(t.dir==='Out'&&t.toName&&t.toName!=='0'&&t.toAcct!==OWN) names.set(t.toAcct,t.toName);}
let resolvedIn=0,resolvedOut=0;
for(const s of stmts.filter(s=>s.fmt==='new')) for(const t of s.tx){
  if(t.dir==='In'){ t.fromNameRaw=t.fromName; t.fromName=names.get(t.fromAcct)||''; if(t.fromName) resolvedIn++; }
  else if(names.has(t.toAcct)){ t.toNameRaw=t.toName; t.toName=names.get(t.toAcct); resolvedOut++; }}
// combine, de-duplicate by F_ID, check duplicates are identical
const seen=new Map();const all=[];let dups=0;
for(const s of stmts){s.dupe=0;for(const t of s.tx){const k=t.fid;if(seen.has(k)){dups++;s.dupe++;const o=seen.get(k);
  ok(['dateStr','amount','bb','ba','type','toNo','fromNo'].every(f=>o[f]===t[f]),`F_ID ${k} differs between statements`);continue;}
  seen.set(k,t);all.push(t);}}
all.sort((a,b)=>a.date-b.date); // stable: keeps statement order for same-second rows
const chainBreaks=[];for(let i=1;i<all.length;i++) if(Math.abs(all[i].bb-all[i-1].ba)>0.005) chainBreaks.push({prev:all[i-1],cur:all[i],gap:r2(all[i].bb-all[i-1].ba)});
ok(new Set(all.map(t=>t.fid)).size===all.length,'duplicate F_ID left');
const keyDup=new Set(all.map(t=>t.dateStr+t.amount+t.bb+t.ba)).size;ok(keyDup===all.length,'duplicate date+amount+balance rows left');
for(let i=0;i+1<stmts.length;i++) ok(stmts[i+1].first<=stmts[i].last || (stmts[i+1].first-stmts[i].last)<86400000*3,`possible gap between ${stmts[i].file} and ${stmts[i+1].file}`);
// overall identity: opening + every balance change + every break = closing
const totalChange=r2(all.reduce((a,t)=>a+t.ba-t.bb,0)), totalGap=r2(chainBreaks.reduce((a,c)=>a+c.gap,0));
ok(Math.abs(r2(all[0].bb+totalChange+totalGap)-all.at(-1).ba)<0.01,'opening + changes + breaks != closing');
console.log('rows',all.length,'dups removed',dups,'| balance swaps fixed',swaps.length,'| names resolved in/out',resolvedIn,resolvedOut);
console.log('chain breaks',chainBreaks.length,'net gap',totalGap);chainBreaks.forEach(c=>console.log('  ',c.prev.dateStr,'->',c.cur.dateStr,'gap',c.gap,c.cur.fmt));
console.log(stmts.map(s=>`${s.file} [${s.fmt}] raw ${s.raw} rows ${s.tx.length} breaks ${s.breaks} in ${s.inT} out ${s.outT} open ${s.open} close ${s.close} unlisted ${s.unlisted}`).join('\n'));
console.log('NOTES',notes);console.log('FAILS',fail.length,fail);
fs.writeFileSync('combined.json',JSON.stringify({all,stmts:stmts.map(({tx,...s})=>s),chainBreaks,dups,fail,swaps,notes,namesLearned:names.size}));
