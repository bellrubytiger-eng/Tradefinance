# -*- coding: utf-8 -*-
"""고객 서식 HTML 생성: python3 tools/build_forms.py  →  screens/forms.html (A4 인쇄용, 단독 문서)"""
import json, pathlib
import model as M

ROOT = pathlib.Path(__file__).resolve().parent.parent
DATA = json.dumps({"forms": M.FORMS, "samples": M.SAMPLES, "branch": M.BRANCH}, ensure_ascii=False)

TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GBS India Customer Forms</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Sans+KR:wght@400;600&family=Source+Serif+4:opsz,wght@8..60,600;8..60,700&display=swap">
<style>
/* Layout: 책상 위 A4 한 장. 서식은 종이이므로 단일(라이트) 테마로 고정하고 모든 색을 명시한다. */
:root{
  --desk:#d8dde4; --paper:#ffffff; --ink:#1b2430; --sub:#556274; --rule:#a9b4c2; --hair:#d9e0e8;
  --navy:#163a63; --tint:#eef3f8; --seal:#a3312a; --map:#1d5fb3; --map-bg:#e6effa;
  --serif:"Source Serif 4",Georgia,serif; --sans:"IBM Plex Sans","IBM Plex Sans KR",Arial,sans-serif; --mono:"IBM Plex Mono",Consolas,monospace;
  color-scheme:light;
}
*{box-sizing:border-box}
[hidden]{display:none!important}
html,body{margin:0;background:var(--desk);color:var(--ink);font:12px/1.4 var(--sans)}
.tools{position:sticky;top:0;z-index:5;background:#1f2a37;color:#e9eef4;display:flex;flex-wrap:wrap;gap:8px 14px;align-items:center;
  padding-block:10px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));font-family:"IBM Plex Sans KR",var(--sans)}
.tools b{font-size:13px;margin-right:6px}
.tools .grp{display:flex;flex-wrap:wrap;gap:4px}
.tools button{font:500 12px "IBM Plex Sans KR",var(--sans);background:#2d3b4c;color:#e9eef4;border:1px solid #415267;border-radius:3px;padding:5px 10px;cursor:pointer}
.tools button.on{background:#e9eef4;color:#1f2a37;border-color:#e9eef4}
.tools label{display:flex;gap:6px;align-items:center;font-size:12px;cursor:pointer}
.tools .sp{flex:1}
.tools a{color:#9cc4ff}
.info{max-width:210mm;margin:14px auto 0;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));font:12.5px/1.5 "IBM Plex Sans KR",var(--sans);color:#2b3644}
.info dl{display:grid;grid-template-columns:auto 1fr;gap:2px 12px;margin:0}
.info dt{font-weight:600;color:#3d4b5c}
.desk{overflow-x:auto;padding:16px max(16px,env(safe-area-inset-right,0px)) 40px max(16px,env(safe-area-inset-left,0px))}
.page{width:210mm;min-height:297mm;margin:0 auto;background:var(--paper);box-shadow:0 2px 14px rgba(20,30,45,.18);padding:10mm 12mm 8mm;position:relative;display:flex;flex-direction:column;gap:6px;overflow:hidden}
.wm{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none;z-index:0}
.wm span{font:700 92px var(--serif);color:rgba(163,49,42,.07);transform:rotate(-32deg);letter-spacing:.08em}
.page>*:not(.wm){position:relative;z-index:1}
.head{display:grid;grid-template-columns:1fr auto;gap:12px;border-bottom:2px solid var(--navy);padding-bottom:6px}
.bank b{font:700 17px var(--serif);color:var(--navy);display:block}
.bank div{color:var(--sub);font-size:9.8px}
.badge{text-align:right;display:flex;flex-direction:column;gap:5px;align-items:flex-end}
.fid{font:500 10.5px var(--mono);color:var(--sub);border:1px solid var(--rule);padding:2px 6px}
.copy{font:600 11px var(--sans);letter-spacing:.14em;padding:3px 9px;border:1.5px solid var(--navy);color:var(--navy)}
.copy.dup{border-color:var(--seal);color:var(--seal)}
.title h1{margin:0;font:700 19px/1.15 var(--serif);color:var(--ink);text-wrap:balance}
.title p{margin:2px 0 0;font-size:10.5px;letter-spacing:.12em;text-transform:uppercase;color:var(--sub)}
.strip{display:grid;grid-template-columns:repeat(var(--n),minmax(0,1fr));border:1px solid var(--rule)}
.strip div{padding:3px 6px;border-right:1px solid var(--hair);min-width:0}
.strip div:last-child{border-right:0}
.lab{display:block;font-size:8.6px;letter-spacing:.08em;text-transform:uppercase;color:var(--sub);margin-bottom:1px}
.val{font-size:10.6px;font-weight:500;overflow-wrap:anywhere}
.val.m{font-family:var(--mono);font-size:10.2px}
.map{display:none;font:9px/1.3 var(--mono);color:var(--map);background:var(--map-bg);padding:1px 4px;margin-top:2px;overflow-wrap:anywhere}
body.mapping .map{display:block}
section h2{margin:0 0 3px;font:600 9.5px var(--sans);letter-spacing:.14em;text-transform:uppercase;color:var(--navy);display:flex;gap:8px;align-items:center}
section h2::after{content:"";flex:1;border-top:1px solid var(--rule)}
.cells{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));grid-auto-flow:row dense;column-gap:18px}
.cell{display:grid;grid-template-columns:34mm minmax(0,1fr);column-gap:8px;align-items:baseline;padding:2px 0;border-bottom:1px solid var(--hair);min-width:0}
.cell .lab{margin:0}
.cell .map{grid-column:2}
.c4{grid-column:1/-1}
.key{display:grid;grid-template-columns:auto 1fr;gap:4px 18px;align-items:end;background:var(--tint);border-left:4px solid var(--navy);padding:5px 10px}
.key .big{font:600 17px var(--mono);font-variant-numeric:tabular-nums;color:var(--navy);white-space:nowrap}
.key .words{font:italic 11px var(--serif);color:var(--ink)}
table.fees{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}
table.fees td{padding:1px 0;border-bottom:1px dotted var(--hair);font-size:10px}
table.fees td.i{color:var(--sub);font-style:italic}
table.legs{width:100%;border-collapse:collapse;font:10px var(--mono)}
table.legs th{font:600 8.4px var(--sans);letter-spacing:.06em;text-transform:uppercase;color:var(--sub);text-align:left;border-bottom:1px solid var(--hair)}
table.legs td{padding:1px 4px 1px 0}
table.legs td:last-child,table.legs th:last-child{padding-left:12px}
.uetr{font-size:9.6px;color:var(--sub);margin-top:-2px}
.uetr span{font-family:var(--mono);color:var(--ink)}
table.legs .n{text-align:right}
table.fees td:last-child{text-align:right;font-family:var(--mono)}
ol.st{margin:0;padding-left:16px;display:grid;gap:2px;font:9.7px/1.35 var(--serif);color:#26303c}
.sign{display:grid;grid-template-columns:repeat(var(--n),minmax(0,1fr));gap:18px;margin-top:auto;padding-top:4px}
.sign div{border-top:1px solid var(--ink);padding-top:3px;font-size:10px;color:var(--sub)}
.sign .seal{border:1px dashed var(--rule);height:40px;border-top-style:dashed;display:flex;align-items:center;justify-content:center}
.foot{display:grid;grid-template-columns:auto 1fr auto;gap:12px;align-items:center;border-top:1px solid var(--rule);padding-top:6px;font-size:9.5px;color:var(--sub)}
.foot canvas{width:44px;height:44px;image-rendering:pixelated;border:1px solid var(--hair)}
.foot .code{font:500 11px var(--mono);color:var(--ink)}
@media print{
  @page{size:A4;margin:0}
  html,body{background:#fff}
  .tools,.info{display:none}
  .desk{padding:0;overflow:visible}
  .page{box-shadow:none;margin:0}
}
@media (max-width:820px){.info dl{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="tools">
  <b>고객 서식</b>
  <div class="grp" id="fsel" role="tablist" aria-label="서식 선택"></div>
  <div class="grp" id="vsel" aria-label="변형 선택"></div>
  <span class="sp"></span>
  <label><input type="checkbox" id="tMap"> DB 필드 매핑 표시</label>
  <label><input type="checkbox" id="tDup"> DUPLICATE 재발급 보기</label>
  <a href="index.html">← 화면 시안</a>
</div>
<div class="info" id="info"></div>
<div class="desk"><div class="page" id="page"></div></div>
<script>
"use strict";
const DATA = __DATA__;
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const fUS = (n,d=2) => Number(n).toLocaleString("en-US",{minimumFractionDigits:d,maximumFractionDigits:d});
const fIN = (n,d=2) => Number(n).toLocaleString("en-IN",{minimumFractionDigits:d,maximumFractionDigits:d});
const ONES=["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen","Seventeen","Eighteen","Nineteen"];
const TENS=["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"];
const two = n => n<20?ONES[n]:TENS[Math.floor(n/10)]+(n%10?" "+ONES[n%10]:"");
const three = n => (n>=100?ONES[Math.floor(n/100)]+" Hundred"+(n%100?" ":""):"")+(n%100?two(n%100):"");
function inrWords(x){ // 인도식: Crore / Lakh / Thousand
  const r=Math.floor(x+1e-9), p=Math.round((x-r)*100); let n=r, out=[];
  const cr=Math.floor(n/1e7); n%=1e7; const lk=Math.floor(n/1e5); n%=1e5; const th=Math.floor(n/1e3); n%=1e3;
  if(cr) out.push((cr>=1000?inrWords(cr).replace(/^Rupees | Only$/g,""):three(cr))+" Crore");
  if(lk) out.push(two(lk)+" Lakh"); if(th) out.push(two(th)+" Thousand"); if(n) out.push(three(n));
  return "Rupees "+(out.join(" ")||"Zero")+(p?" and "+two(p)+" Paise":"")+" Only";
}
const SPAN = {c4:["ben_addr","rem_addr","inr_fixed","uetr","rmt_info","purpose","fees","ben_bank","rem_bank","intm","lrs_fy","tcs","kyc","inv_ref","form_a2","f15ca","gst_info","conv_legs","payout","charges","susp","nro_fy","fx_type","rbi_ref","sell_acct","buy_acct"], c2:[]};
let cur={form:"firc",variant:"fdi"};

function value(key, d){
  const v=d[key], ccy=d.ccy;
  const feeSum=(d.fees||[]).reduce((a,f)=>a+(f[1]||0),0);
  const netDed = d.fee_mode==="SEPARATE"?0:feeSum;
  if(typeof v==="string" && !["fees","conv_legs"].includes(key)) return esc(v);
  switch(key){
    case "instd": case "fcy": case "amt": case "fchg": return `${ccy} ${fUS(v)}`;
    case "inr": return `INR ${fIN(v)}`;
    case "rate": return `${fUS(v,4)} INR per ${ccy}`;
    case "total": return `INR ${fIN(d.inr+feeSum)}`;
    case "net": return `INR ${fIN(d.inr-netDed)}`+(d.fee_mode==="SEPARATE"?" (charges collected separately)":"");
    case "inr_words": return inrWords(d.inr);
    case "total_words": return inrWords(d.inr+feeSum);
    case "net_words": return inrWords(d.inr-netDed);
    case "buy_words": return d.buy_inr?inrWords(d.buy_inr):"Foreign currency credited - see Exchange";
    case "conv_legs": return `<table class="legs"><tr><th>Date</th><th class="n">${ccy}</th><th class="n">Rate</th><th class="n">INR</th><th>Deal</th></tr>${(v||[]).map(l=>`<tr><td>${l[0]}</td><td class="n">${fUS(l[1])}</td><td class="n">${fUS(l[2],4)}</td><td class="n">${fIN(l[3])}</td><td>${esc(l[4])}</td></tr>`).join("")}</table>`;
    case "chg_type": return {OUR:"OUR - all charges paid by remitter",SHA:"SHA - shared",BEN:"BEN - paid by beneficiary"}[v]||v;
    case "fees": return `<table class="fees">${(d.fees||[]).map(f=>f[1]===null?`<tr><td class="i" colspan="2">${esc(f[0])}</td></tr>`:`<tr><td>${esc(f[0])}</td><td>INR ${fIN(f[1])}</td></tr>`).join("")}<tr><td><b>Total charges and taxes</b></td><td><b>INR ${fIN(feeSum)}</b></td></tr></table>`;
  }
  return esc(v ?? "-");
}
const MONO = new Set(["cert_no","adv_no","our_ref","their_ref","fdi_ref","uetr","rem_acct","ben_acct","rem_pan","ben_pan","ben_cin","rem_iec","ben_iec","rem_lei","orm","irm","instd","fcy","fchg","amt","inr","rate","total","net","cust_no"]);

function cellHTML(f, d){
  const [key,en,ko,src]=f, span=SPAN.c4.includes(key)?"c4":SPAN.c2.includes(key)?"c2":"";
  return `<div class="cell ${span}"><span class="lab">${esc(en)}</span><div class="val ${MONO.has(key)?"m":""}">${value(key,d)}</div><div class="map">${esc(ko)} · ${esc(src)}</div></div>`;
}

function render(){
  const F=DATA.forms.find(x=>x.key===cur.form), d=Object.assign({},DATA.samples[F.key][cur.variant]), B=DATA.branch, dup=$("#tDup").checked;
  const show=f=>!f[4]||f[4].includes(cur.variant);
  const meta=F.sections[0], body=F.sections.slice(1);
  const keyAmt = {firc:["INR Amount Credited","inr","inr_words"], out:["Total Amount Debited","total","total_words"], in:["Net Amount Credited","net","net_words"], fx:["You Received","buy","buy_words"]}[F.key];
  const copyTxt = dup?"DUPLICATE":d.copy;
  const metaF = meta.f.filter(show).filter(f=>f[0]!=="uetr");
  let h=`<div class="wm"><span>SPECIMEN</span></div>
  <header class="head"><div class="bank"><b>${esc(B.bank)}</b><div>${esc(B.branch)} · Authorised Dealer Category-I</div><div>${esc(B.addr)}</div><div>AD Code ${esc(B.ad_code)} · SWIFT ${esc(B.bic)} · GSTIN ${esc(B.gstin)}</div></div>
    <div class="badge"><span class="fid">${F.id}</span><span class="copy ${dup?"dup":""}">${copyTxt}</span></div></header>
  <div class="title"><h1>${esc(F.title)}</h1><p>${esc(F.subtitle)}</p></div>
  <div class="strip" style="--n:${metaF.length}">${metaF.map(f=>`<div><span class="lab">${esc(f[1])}</span><div class="val m">${value(f[0],d)}</div><div class="map">${esc(f[2])} · ${esc(f[3])}</div></div>`).join("")}</div>`+
    (d.uetr&&meta.f.some(f=>f[0]==="uetr")?`<div class="uetr">UETR (SWIFT gpi tracking) <span>${esc(d.uetr)}</span></div>`:"");
  body.forEach(sec=>{
    const fs=sec.f.filter(show).filter(f=>!(f[0]===keyAmt[1]||f[0]===keyAmt[2]));
    if(sec.t.startsWith("Remittance and")||sec.t==="Amount"||sec.t==="Exchange") h+=`<div class="key"><div><span class="lab">${keyAmt[0]}</span><div class="big">${value(keyAmt[1],d)}</div></div><div class="words">${value(keyAmt[2],d)}</div></div>`;
    h+=`<section><h2>${esc(sec.t)}</h2><div class="cells">${fs.map(f=>cellHTML(f,d)).join("")}</div></section>`;
  });
  const st=F.statements.filter(s=>!s[1]||s[1].includes(cur.variant)).map(s=>s[0]);
  if(dup) st.push(`This is a duplicate of ${d.cert_no||d.adv_no} originally issued on ${d.issue_dt||d.adv_dt}, reissued on request. The original stands cancelled for any further use.`);
  h+=`<section><h2>${F.key==="firc"?"Certification":"Notes"}</h2><ol class="st">${st.map(s=>`<li>${esc(s)}</li>`).join("")}</ol></section>`;
  h+=`<div class="sign" style="--n:${F.sign.length}">${F.sign.map(s=>s==="Branch Seal"?`<div class="seal">${s}</div>`:`<div>${s}<br>Name / Employee ID</div>`).join("")}</div>`;
  h+=`<footer class="foot"><canvas id="qr" width="21" height="21" aria-label="Verification code pattern"></canvas><div>Verify at the branch or through internet banking with the code below.<br><span class="code">${esc(d.verify)}</span> · Issue ${dup?"2 (DUPLICATE)":"1 (ORIGINAL)"}<div class="map">${"TB_FX_ADVICE_ISSUE.VERIFY_CODE, ISSUE_TYPE, ISSUE_SEQ"}</div></div><div style="text-align:right">Computer-generated document<br>${F.id} · Page 1 of 1</div></footer>`;
  $("#page").innerHTML=h;
  drawCode(d.verify);
  $("#info").innerHTML=`<dl><dt>${esc(F.ko)}</dt><dd>${esc(F.title)} — ${esc(DATA.forms.find(x=>x.key===cur.form).variants.find(v=>v[0]===cur.variant)[1])}</dd><dt>발급 화면</dt><dd>${esc(F.screen)}</dd><dt>발급 시점</dt><dd>${esc(F.when)}</dd><dt>용도</dt><dd>${esc(F.use)}</dd><dt>보관</dt><dd>${esc(F.copies)}</dd></dl>`;
  document.title = `${F.id} ${F.title}`;
  $$("#fsel button").forEach(b=>b.classList.toggle("on",b.dataset.f===cur.form));
  $("#vsel").innerHTML=F.variants.map(v=>`<button data-v="${v[0]}" class="${v[0]===cur.variant?"on":""}">${esc(v[1])}</button>`).join("");
}
function $$(s){return [...document.querySelectorAll(s)];}
function drawCode(code){ // 진위확인 코드의 시각 패턴(실제 QR 아님 - 운영 시 QR 라이브러리로 대체)
  const c=$("#qr"), x=c.getContext("2d"); x.fillStyle="#ffffff"; x.fillRect(0,0,21,21); x.fillStyle="#1b2430";
  let h=0; for(const ch of code) h=(h*31+ch.charCodeAt(0))>>>0;
  const fin=(a,b)=>{x.fillRect(a,b,7,7);x.fillStyle="#ffffff";x.fillRect(a+1,b+1,5,5);x.fillStyle="#1b2430";x.fillRect(a+2,b+2,3,3);};
  fin(0,0);fin(14,0);fin(0,14);
  for(let i=0;i<21;i++)for(let j=0;j<21;j++){ if((i<8&&j<8)||(i>12&&j<8)||(i<8&&j>12)) continue; h=(h*1103515245+12345)>>>0; if(h&0x10000) x.fillRect(i,j,1,1); }
}
function setFromHash(){ const [f,v]=location.hash.slice(1).split("-"); const F=DATA.forms.find(x=>x.key===f);
  if(F){ cur.form=f; cur.variant=F.variants.some(x=>x[0]===v)?v:F.variants[0][0]; } render(); }
$("#fsel").innerHTML=DATA.forms.map(f=>`<button data-f="${f.key}">${esc(f.ko)}</button>`).join("");
document.addEventListener("click",e=>{ const b=e.target.closest("button"); if(!b) return;
  if(b.dataset.f){ const F=DATA.forms.find(x=>x.key===b.dataset.f); location.hash=`${F.key}-${F.variants[0][0]}`; }
  if(b.dataset.v){ location.hash=`${cur.form}-${b.dataset.v}`; } });
$("#tMap").addEventListener("change",e=>document.body.classList.toggle("mapping",e.target.checked));
$("#tDup").addEventListener("change",render);
window.addEventListener("hashchange",setFromHash);
setFromHash();
</script>
</body>
</html>
"""

out = ROOT / "screens/forms.html"
out.write_text(TEMPLATE.replace("__DATA__", DATA), encoding="utf-8")
print("saved", out)
