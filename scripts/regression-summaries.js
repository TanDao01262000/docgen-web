#!/usr/bin/env node
/**
 * Regression: parse every summary under a fixtures dir (default /workspace/docgen-real).
 * Never commits xlsx. Run: node scripts/regression-summaries.js
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const {execFileSync} = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const FIXTURES = process.env.DOCGEN_FIXTURES || '/workspace/docgen-real';

function loadSandbox(){
  const sandbox = {
    console,
    window: {},
    globalThis: {},
    setTimeout, clearTimeout,
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  const xlsx = fs.readFileSync(path.join(ROOT, 'vendor/xlsx.full.min.js'), 'utf8');
  vm.runInContext(xlsx, sandbox, {filename:'xlsx.full.min.js'});
  const voucher = fs.readFileSync(path.join(ROOT, 'voucher.js'), 'utf8');
  vm.runInContext(voucher, sandbox, {filename:'voucher.js'});
  if(!sandbox.Voucher) throw new Error('Voucher not loaded');
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
  const p = path.join(ROOT, 'hotels.json');
  try { return JSON.parse(fs.readFileSync(p,'utf8')); } catch(e){ return []; }
}

function parseFile(sandbox, file, hotels){
  const name = path.basename(file);
  if(/\.pdf$/i.test(file)){
    let text='';
    try{
      text = execFileSync('pdftotext', ['-layout', file, '-'], {encoding:'utf8', timeout:20000});
    }catch(e){
      return {file:name, kind:'pdf', error:'pdftotext failed: '+(e.message||e), guests:0, hotels:0, transfers:0};
    }
    const data = sandbox.Voucher.parseSummaryFromPdfText(text, hotels);
    return summarize(name, 'pdf', data);
  }
  const buf = fs.readFileSync(file);
  const wb = sandbox.XLSX.read(buf, {type:'buffer', cellDates:true, raw:false});
  const data = sandbox.Voucher.parseSummary(wb, hotels);
  return summarize(name, 'xlsx', data, wb);
}

function summarize(name, kind, data, wb){
  const guests = data.guests||[];
  const acc = data.acc_rows||[];
  const trf = data.trf_rows||[];
  return {
    file: name,
    kind,
    sheets: wb ? (wb.SheetNames||[]) : [],
    guests: guests.length,
    guest_lines: guests,
    pax: data.pax,
    contact: data.contact||'',
    hotels: acc.filter(r=>r.is_hotel).length,
    trains: acc.filter(r=>r.is_train).length,
    acc: acc.length,
    hotel_names: acc.filter(r=>r.is_hotel).map(r=>r.hotel_name||r.matched_name),
    transfers: trf.length,
    trf_labels: trf.map(r=>(r.label||'').split('\n')[0]),
    flags: (data.flags||[]).map(f=>f.level+': '+f.msg),
    empty: !guests.length || !acc.length || !trf.length
  };
}

// Expected floors. garbage/lookup fixtures may differ — discover then lock.
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

function main(){
  const sandbox = loadSandbox();
  const hotels = hotelsStub();
  const files = listFixtures(FIXTURES);
  if(!files.length){
    console.error('No fixtures in', FIXTURES);
    process.exit(2);
  }
  const results=[];
  let failed=0;
  for(const f of files){
    let r;
    try{ r=parseFile(sandbox,f,hotels); }
    catch(e){ r={file:path.basename(f), error:String(e.stack||e), guests:0, hotels:0, transfers:0, empty:true}; }
    results.push(r);
    const exp=EXPECT[r.file] || {minGuests:1, minHotels:1, minTransfers:1};
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
    r.ok = !problems.length;
    r.problems = problems;
    if(!r.ok){ failed++; }
    console.log((r.ok?'PASS':'FAIL')+' '+r.file+' guests='+r.guests+' hotels='+r.hotels+' trains='+(r.trains||0)+' trf='+r.transfers+' pax='+r.pax+' contact='+JSON.stringify(r.contact||''));
    if(r.guest_lines) console.log('    guests:', r.guest_lines.join(' | '));
    if(r.hotel_names) console.log('    hotels:', r.hotel_names.join(' | '));
    if(r.trf_labels) console.log('    trf:', r.trf_labels.join(' | '));
    if(r.flags && r.flags.length) console.log('    flags:', r.flags.join(' || '));
    if(problems.length) console.log('    PROBLEMS:', problems.join('; '));
  }
  const out = path.join(FIXTURES, 'compare/regression_summaries.json');
  try{ fs.mkdirSync(path.dirname(out), {recursive:true}); fs.writeFileSync(out, JSON.stringify(results,null,2)); }catch(_){}
  console.log('\n'+ (failed?failed+' FAILED':'ALL PASSED') + '  ('+results.length+' files)');
  process.exit(failed?1:0);
}
main();
