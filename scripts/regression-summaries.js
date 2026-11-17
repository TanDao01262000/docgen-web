#!/usr/bin/env node
/**
 * Regression: parse every summary under a fixtures dir (default /workspace/docgen-real).
 * Also: for each xlsx, mutate sheet names/order + add a junk sheet and assert
 * parse output is identical (content-only sheet detection).
 * Never commits xlsx. Run: node scripts/regression-summaries.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const {execFileSync} = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const FIXTURES = process.env.DOCGEN_FIXTURES || '/workspace/docgen-real';

function loadSandbox(){
  const sandbox = { console, window: {}, globalThis: {}, setTimeout, clearTimeout };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'vendor/xlsx.full.min.js'), 'utf8'), sandbox, {filename:'xlsx.full.min.js'});
  vm.runInContext(fs.readFileSync(path.join(ROOT, 'voucher.js'), 'utf8'), sandbox, {filename:'voucher.js'});
  if(!sandbox.Voucher || !sandbox.XLSX) throw new Error('Voucher/XLSX not loaded');
  return sandbox;
}

function listFixtures(dir){
  if(!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .filter(n => /\.(xlsx|xls|pdf)$/i.test(n))
    .map(n => path.join(dir, n))
    .sort();
}

function hotelsStub(){
  try { return JSON.parse(fs.readFileSync(path.join(ROOT, 'hotels.json'),'utf8')); }
  catch(e){ return []; }
}

function fingerprint(data){
  const acc=data.acc_rows||[];
  const trf=data.trf_rows||[];
  return {
    guests: data.guests||[],
    pax: String(data.pax||''),
    contact: data.contact||'',
    hotels: acc.filter(r=>r.is_hotel).map(r=>({name:r.hotel_name||r.matched_name||'', city:r.city||'', cin:r.check_in||'', cout:String(r.check_out||'')})),
    trains: acc.filter(r=>r.is_train).map(r=>r.train_text||r.city||''),
    transfers: trf.map(r=>({label:(r.label||'').replace(/\s+/g,' ').trim(), date:r.date||''}))
  };
}

function summarize(name, kind, data, extra){
  const fp=fingerprint(data);
  return Object.assign({
    file: name,
    kind,
    guests: fp.guests.length,
    guest_lines: fp.guests,
    pax: fp.pax,
    contact: fp.contact,
    hotels: fp.hotels.length,
    hotel_names: fp.hotels.map(h=>h.name),
    trains: fp.trains.length,
    acc: (data.acc_rows||[]).length,
    transfers: fp.transfers.length,
    trf_labels: fp.transfers.map(r=>r.label.split('\n')[0]||r.label),
    flags: (data.flags||[]).map(f=>f.level+': '+f.msg),
    empty: !fp.guests.length || !fp.hotels.length || !fp.transfers.length,
    fingerprint: fp
  }, extra||{});
}

function mulberry32(a){
  return function(){ let t=a+=0x6D2B79F5; t=Math.imul(t^t>>>15, t|1); t^=t+Math.imul(t^t>>>7, t|61); return ((t^t>>>14)>>>0)/4294967296; };
}
function hashStr(s){ let h=2166136261; for(let i=0;i<s.length;i++){ h^=s.charCodeAt(i); h=Math.imul(h,16777619);} return h>>>0; }
function shuffle(arr, rng){
  const a=arr.slice();
  for(let i=a.length-1;i>0;i--){ const j=Math.floor(rng()*(i+1)); [a[i],a[j]]=[a[j],a[i]]; }
  return a;
}

function mutateWorkbook(XLSX, wb, seed){
  const rng=mulberry32(hashStr(seed||'x')||1);
  const token=()=>Math.floor(rng()*1e9).toString(36);
  const oldNames=(wb.SheetNames||[]).slice();
  const order=shuffle(oldNames, rng);
  const newSheets={};
  const newNames=[];
  const meta=[];
  for(const old of order){
    const nn='x_'+token()+'_'+newNames.length;
    newSheets[nn]=wb.Sheets[old];
    newNames.push(nn);
    const prev=((wb.Workbook&&wb.Workbook.Sheets)||[]).find(s=>s && (s.name===old||s.Name===old));
    meta.push({name:nn, Hidden:(prev&&prev.Hidden)||0});
  }
  const junkName='junk_'+token();
  newSheets[junkName]=XLSX.utils.aoa_to_sheet([
    ['NOT A TOUR FILE','random',999],
    ['invoices','amount','date'],
    ['hello world','foo','bar'],
    [12,34,56],
    ['notes','lorem ipsum dolor sit amet']
  ]);
  const junkAt=Math.floor(rng()*(newNames.length+1));
  newNames.splice(junkAt, 0, junkName);
  meta.splice(junkAt, 0, {name:junkName, Hidden:0});
  return {SheetNames:newNames, Sheets:newSheets, Workbook:{Sheets:meta}, _mutation:{from:oldNames, to:newNames, junk:junkName}};
}

const EXPECT = {
  'sample-summary.xlsx': {minGuests:1, exactGuests:1, minHotels:1, exactHotels:1, minTransfers:2, exactTransfers:2},
  'summary.xlsx': {minGuests:2, exactGuests:2, minHotels:6, exactHotels:6, minTransfers:15, exactTransfers:15},
  'summary_cat_tien_pq.xlsx': {minGuests:2, exactGuests:2, minHotels:5, exactHotels:5, minTransfers:8, exactTransfers:8},
  'summary_guest_bug.xlsx': {minGuests:2, exactGuests:2, minHotels:6, exactHotels:6, minTransfers:12, exactTransfers:12},
  'summary_lookup_qa.xlsx': {minGuests:1, exactGuests:1, minHotels:1, exactHotels:1, minTransfers:2, exactTransfers:2},
  'summary_serper_test.xlsx': {minGuests:2, exactGuests:2, minHotels:6, exactHotels:6, minTransfers:15, exactTransfers:15},
  'summary_vesela.xlsx': {minGuests:4, exactGuests:4, minHotels:6, exactHotels:6, minTransfers:11, exactTransfers:11},
  'summary_wife_empty.xlsx': {minGuests:4, exactGuests:4, minHotels:6, exactHotels:6, minTransfers:12, exactTransfers:12},
  'summary_garbage.xlsx': {minGuests:2, exactGuests:2, minHotels:6, exactHotels:6, minTransfers:15, exactTransfers:15},
  'wife_cat_tien_pq.pdf': {minGuests:4, exactGuests:4, minHotels:6, exactHotels:6, minTransfers:11}
};

function checkExpect(r, exp){
  const problems=[];
  if(r.error && !exp.optional) problems.push(r.error);
  if(!exp.optional){
    if(r.guests < (exp.minGuests||1)) problems.push(`guests ${r.guests} < ${exp.minGuests||1}`);
    if((r.hotels||0) < (exp.minHotels||1)) problems.push(`hotels ${r.hotels} < ${exp.minHotels||1}`);
    if((r.transfers||0) < (exp.minTransfers||1)) problems.push(`transfers ${r.transfers} < ${exp.minTransfers||1}`);
    if(exp.exactGuests!=null && r.guests!==exp.exactGuests) problems.push(`guests ${r.guests} != ${exp.exactGuests}`);
    if(exp.exactHotels!=null && r.hotels!==exp.exactHotels) problems.push(`hotels ${r.hotels} != ${exp.exactHotels}`);
    if(exp.exactTransfers!=null && r.transfers!==exp.exactTransfers) problems.push(`transfers ${r.transfers} != ${exp.exactTransfers}`);
  }
  return problems;
}

function parseXlsx(sandbox, buf, hotels){
  const wb=sandbox.XLSX.read(buf, {type:'buffer', cellDates:true, raw:false});
  const data=sandbox.Voucher.parseSummary(wb, hotels);
  return {wb, data};
}

function main(){
  const sandbox=loadSandbox();
  const hotels=hotelsStub();
  const files=listFixtures(FIXTURES);
  if(!files.length){ console.error('No fixtures in', FIXTURES); process.exit(2); }
  const results=[];
  let failed=0, variantFails=0, variantRuns=0;

  for(const f of files){
    const name=path.basename(f);
    let r;
    try{
      if(/\.pdf$/i.test(f)){
        let text='';
        try{ text=execFileSync('pdftotext',['-layout',f,'-'],{encoding:'utf8', timeout:20000}); }
        catch(e){ r={file:name, kind:'pdf', error:'pdftotext failed: '+(e.message||e), guests:0, hotels:0, transfers:0, empty:true}; }
        if(!r){
          const data=sandbox.Voucher.parseSummaryFromPdfText(text, hotels);
          r=summarize(name,'pdf',data);
        }
      } else {
        const buf=fs.readFileSync(f);
        const {wb, data}=parseXlsx(sandbox, buf, hotels);
        const det=sandbox.Voucher.detectSheets(wb);
        r=summarize(name,'xlsx',data,{detect:{sumScore:det.sumScore, cliScore:det.cliScore, sumName:det.sumName, cliName:det.cliName, origSheets:wb.SheetNames}});
        // Content-only variants: rename + reorder + junk sheet
        const fp=JSON.stringify(fingerprint(data));
        const variantResults=[];
        for(let i=0;i<2;i++){
          variantRuns++;
          const mutated=mutateWorkbook(sandbox.XLSX, wb, name+'#'+i);
          const data2=sandbox.Voucher.parseSummary(mutated, hotels);
          const fp2=JSON.stringify(fingerprint(data2));
          const ok=fp===fp2;
          if(!ok) variantFails++;
          variantResults.push({
            ok,
            sheets:mutated.SheetNames,
            guests:(data2.guests||[]).length,
            hotels:(data2.acc_rows||[]).filter(x=>x.is_hotel).length,
            transfers:(data2.trf_rows||[]).length,
            diff: ok?null:{orig:JSON.parse(fp), got:fingerprint(data2)}
          });
          if(!ok){
            console.log('  VARIANT FAIL', name, 'seed', i, 'sheets', mutated.SheetNames.join(','));
          }
        }
        r.variants=variantResults;
        r.variantsOk=variantResults.every(v=>v.ok);
      }
    }catch(e){
      r={file:name, error:String(e.stack||e), guests:0, hotels:0, transfers:0, empty:true, variantsOk:false};
    }
    const exp=EXPECT[r.file] || {minGuests:1, minHotels:1, minTransfers:1};
    r.problems=checkExpect(r, exp);
    if(r.variants && !r.variantsOk) r.problems.push('sheet-name variants did not parse identically');
    r.ok=!r.problems.length;
    if(!r.ok) failed++;
    console.log((r.ok?'PASS':'FAIL')+' '+r.file+' guests='+r.guests+' hotels='+r.hotels+' trains='+(r.trains||0)+' trf='+r.transfers+' pax='+r.pax+' contact='+JSON.stringify(r.contact||'')+(r.variants?(' variants='+(r.variantsOk?'ok':'FAIL')):''));
    if(r.guest_lines) console.log('    guests:', r.guest_lines.join(' | '));
    if(r.hotel_names) console.log('    hotels:', r.hotel_names.join(' | '));
    if(r.trf_labels) console.log('    trf:', r.trf_labels.join(' | '));
    if(r.detect) console.log('    detect:', JSON.stringify({sum:r.detect.sumName, sumScore:r.detect.sumScore, cli:r.detect.cliName, cliScore:r.detect.cliScore}));
    if(r.flags && r.flags.length) console.log('    flags:', r.flags.join(' || '));
    if(r.problems && r.problems.length) console.log('    PROBLEMS:', r.problems.join('; '));
    results.push(r);
  }

  const out=path.join(FIXTURES, 'compare/regression_summaries.json');
  try{ fs.mkdirSync(path.dirname(out), {recursive:true}); fs.writeFileSync(out, JSON.stringify(results,null,2)); }catch(_){}
  console.log('\n'+(failed?failed+' FAILED':'ALL PASSED')+'  ('+results.length+' files, '+variantRuns+' sheet-name variants, '+variantFails+' variant failures)');
  process.exit(failed?1:0);
}
main();
