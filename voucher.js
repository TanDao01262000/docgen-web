/* Voucher mode: parse GO2 summary xlsx -> review -> 2 vouchers */
(function(global){
const MONTHS={JAN:1,FEB:2,MAR:3,APR:4,MAY:5,JUN:6,JUL:7,AUG:8,SEP:9,OCT:10,NOV:11,DEC:12};
const MON=['','JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
function norm(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D').toUpperCase().replace(/[^A-Z0-9]+/g,' ').trim();}
function todayVN(){const d=new Date();const p=n=>String(n).padStart(2,'0');return `${p(d.getDate())}.${p(d.getMonth()+1)}.${d.getFullYear()}`;}
function parseDayMon(s){// "12JUN" / "12JUN _06:05" / "16JUN - Late check out"
  s=String(s||'').trim();
  const m=s.match(/(\d{1,2})\s*([A-Za-z]{3})/);
  if(!m) return {raw:s, dayMon:'', late:/late/i.test(s)};
  const mon=m[2].toUpperCase().slice(0,3);
  return {raw:s, dayMon:`${m[1].padStart(2,'0')}-${mon}`, late:/late\s*check\s*out/i.test(s), time: (s.match(/(\d{1,2}:\d{2})/)||[])[1]||'', after:/after\s+cruise/i.test(s)};
}
function parseTimeFrom(s){
  s=String(s||'');
  const range=s.match(/(\d{1,2}:\d{2})\s*[-–]\s*(\d{1,2}:\d{2})/);
  if(range) return `${range[1]} - ${range[2]}`;
  if(/after\s+cruise/i.test(s)) return 'after cruise';
  const t=s.match(/(\d{1,2}:\d{2})/);
  return t?t[1]:'';
}
function fmtFlight(a){
  // Keep summary spelling (e.g. "VN 036" vs "VN119"); only collapse whitespace
  return String(a||'').trim().replace(/\s+/g,' ');
}


function looksLikeProseAddress(addr){
  const s=String(addr||'').trim();
  if(!s) return true;
  if(s.length>120) return true;
  if(/can be contacted|is located|please contact|the resort is|hotel is|\b(our|we|you)\b/i.test(s)) return true;
  if(/https?:\/\//i.test(s)) return true;
  // Prefer addresses with a digit or clear place token
  if(!/\d/.test(s) && !/(street|st\b|road|rd\b|ave|district|ward|ap\b|hamlet|island|phu quoc|sai gon|hanoi|danang|da nang)/i.test(s)) return true;
  return false;
}
function shortenAddress(addr){
  let a=String(addr||'').replace(/\s+/g,' ').trim();
  if(!a) return '';
  // Plus Codes e.g. 7P28+X5 / RJ8G+6V Phu Quoc
  a=a.replace(/\b[A-Z0-9]{4,}\+[A-Z0-9]{2,}\b/gi,'');
  // postal / zip
  a=a.replace(/\b\d{4,6}\b/g,(m,off,str)=>{
    // keep street numbers that are short and mid-string; drop 5-6 digit zips
    return m.length>=5?'':m;
  });
  a=a.replace(/,?\s*(Socialist Republic of )?Viet\s*Nam\b/ig,'');
  a=a.replace(/,?\s*Vietnam\b/ig,'');
  // collapse duplicate province/city tokens
  const parts=a.split(/,\s*/).map(x=>x.trim()).filter(Boolean);
  const seen=new Set(); const out=[];
  for(const part of parts){
    const k=norm(part);
    if(!k||seen.has(k)) continue;
    // drop pure "Province" noise if city already present
    seen.add(k); out.push(part);
  }
  a=out.join(', ').replace(/\s+,/g,',').replace(/,\s*,/g,',').replace(/^,\s*|,\s*$/g,'').trim();
  if(a.length>90){
    // keep first ~2-3 comma segments
    const segs=a.split(/,\s*/);
    a=segs.slice(0, Math.min(3, segs.length)).join(', ');
    if(a.length>90) a=a.slice(0,87).replace(/\s+\S*$/,'')+'…';
  }
  return a;
}
function canonicalCity(raw){
  const n=norm(raw);
  if(!n) return '';
  if(/HA\s*NOI|HANOI|\bHAN\b/.test(n) && !/HOTEL|AIRPORT/.test(n)) return 'HA NOI';
  if(/HA\s*LONG|HALONG/.test(n)) return 'HA LONG BAY';
  if(/NINH\s*BINH/.test(n)) return 'NINH BINH';
  if(/^HUE\b|HUE\b/.test(n) && !/HOTEL/.test(n)) return 'HUE';
  if(/HOI\s*AN/.test(n)) return 'HOI AN';
  if(/SAI\s*GON|SAIGON|HCMC|HO\s*CHI\s*MINH/.test(n)) return 'SAI GON';
  if(/CAT\s*TIEN|CATTIEN/.test(n)) return 'CAT TIEN';
  if(/PHU\s*QUOC|\bPQC\b|\bPQ\b/.test(n)) return 'PHU QUOC';
  if(/DA\s*NANG|DANANG/.test(n)) return 'DA NANG';
  if(/NHA\s*TRANG/.test(n)) return 'NHA TRANG';
  return n.replace(/\bHOMESTAY\b/,'').replace(/\bHOTEL\b$/,'').trim();
}
const VN_AIRPORTS=new Set(['HAN','SGN','DAD','PQC','CXR','DLI','VCA','HPH','VII','UIH','VDO','VDH','VXN','VKG','THD','DIN','PXU','UIH','BMV','CAH','VCS','TBB']);
function aptCode(label){
  const n=norm(label);
  if(/\bHAN\b|HA\s*NOI|HANOI/.test(n)) return 'HAN';
  if(/\bSGN\b|SAI\s*GON|SAIGON|HCMC|TAN\s*SON/.test(n)) return 'SGN';
  if(/\bDAD\b|DA\s*NANG|DANANG/.test(n)) return 'DAD';
  if(/\bPQC\b|PHU\s*QUOC|\bPQ\b/.test(n)) return 'PQC';
  if(/\bCXR\b|NHA\s*TRANG|CAM\s*RANH/.test(n)) return 'CXR';
  if(/\bDLI\b|DA\s*LAT|DALAT/.test(n)) return 'DLI';
  const m=n.match(/\b([A-Z]{3})\b/);
  return m?m[1]:'';
}
function isVnAirport(code){ return VN_AIRPORTS.has(String(code||'').toUpperCase()); }
function flightIsDomestic(fl){
  if(!fl) return null;
  const a=aptCode(fl.from), b=aptCode(fl.to);
  if(!a||!b) return null;
  return isVnAirport(a)&&isVnAirport(b);
}
const APT_PICKUP={
  HAN:{name:'Hanoi', intl:{term:'international terminal', pillar:'pillar no.10'}, dom:{term:'domestic terminal', pillar:'pillar no.10'}},
  SGN:{name:'Sai Gon', intl:{term:'international terminal', pillar:'pillar no. 23B'}, dom:{term:'domestic terminal', pillar:'pillar no. 23B'}},
  DAD:{name:'Da Nang', intl:{term:'international terminal', pillar:'the arrivals hall'}, dom:{term:'Domestic terminal', pillar:'the arrivals hall'}},
  PQC:{name:'Phu Quoc', intl:{term:'international terminal', pillar:'the arrivals hall'}, dom:{term:'domestic terminal', pillar:'the arrivals hall'}},
  CXR:{name:'Cam Ranh', intl:{term:'international terminal', pillar:'the arrivals hall'}, dom:{term:'domestic terminal', pillar:'the arrivals hall'}},
  DLI:{name:'Da Lat', intl:{term:'terminal', pillar:'the arrivals hall'}, dom:{term:'terminal', pillar:'the arrivals hall'}}
};
function airportArrivalNote(code, fl, surname){
  const dom=flightIsDomestic(fl);
  const rule=APT_PICKUP[code];
  if(!rule){
    return {note:`The driver will pick you up at the airport arrivals hall with a sign:\n\n GO2/ MR. ${surname.toUpperCase()}`, unknown:true, kind:dom===false?'international':dom===true?'domestic':'unknown'};
  }
  const side=(dom===false)?rule.intl:rule.dom;
  const kind=dom===false?'international':dom===true?'domestic':(dom===null?'unknown':'domestic');
  // Prefer sample-style phrasing for HAN intl / SGN domestic
  if(code==='HAN' && dom===false)
    return {note:`The driver will pick you up at the international terminal (outside the arrivals hall at ${side.pillar}) with a sign:\n\n GO2/ MR. ${surname.toUpperCase()}`, unknown:false, kind:'international'};
  if(code==='SGN' && dom!==false)
    return {note:`The driver will pick you up at the domestic terminal of Sai Gon airport (outside the arrivals hall ${side.pillar}) with a sign:\n\n GO2/ MR. ${surname.toUpperCase()}`, unknown:false, kind:'domestic'};
  const termLabel=side.term;
  const where=(/arrivals hall/i.test(side.pillar) && !/pillar/i.test(side.pillar))
    ? `at the ${termLabel} of ${rule.name} airport (outside the arrivals hall)`
    : `at the ${termLabel} of ${rule.name} airport (outside the arrivals hall at ${side.pillar})`;
  return {note:`The driver will pick you up ${where} with a sign:\n\n GO2/ MR. ${surname.toUpperCase()}`, unknown:dom===null, kind};
}
function airportDropoffNote(code, fl){
  const dom=flightIsDomestic(fl);
  const rule=APT_PICKUP[code];
  const kind=dom===false?'International':dom===true?'Domestic':'';
  if(!rule){
    return {note:`The driver will pick you up at the hotel reception and drop you off at the airport.`, unknown:true, kind:kind||'unknown'};
  }
  // Prefer flight route; if unknown, leave editable (do not assume Domestic)
  const term=(dom===false)?'International terminal':(dom===true)?'Domestic terminal':'terminal (domestic/international — confirm)';
  if(code==='SGN' && dom===false)
    return {note:`The driver will pick you up at the hotel reception and drop you off at the International terminal of Saigon airport.`, unknown:false, kind:'International'};
  if(code==='DAD' && dom===true)
    return {note:`The driver will pick you up at the hotel reception and drop you off at the Domestic terminal of Da Nang airport.`, unknown:false, kind:'Domestic'};
  if(code==='DAD' && dom===false)
    return {note:`The driver will pick you up at the hotel reception and drop you off at the International terminal of Da Nang airport.`, unknown:false, kind:'International'};
  return {note:`The driver will pick you up at the hotel reception and drop you off at the ${term} of ${rule.name} airport.`, unknown:dom===null, kind:kind||term};
}
function dateKey(s){
  const m=String(s||'').match(/(\d{1,2})\s*[.\-\s]*([A-Za-z]{3})/);
  if(!m) return '';
  return m[1].padStart(2,'0')+m[2].toUpperCase().slice(0,3);
}
function parseMeals(blob){
  const s=String(blob||'');
  if(!s.trim()) return 'No meals';
  if(/Meal:\s*Street food/i.test(s)) return 'Street food';
  if(/street\s*food/i.test(s)) return /Meal:/i.test(s)?'Street food':'street food';
  // explicit no
  const hasFoodToken=/breakfast|lunch|dinner|brunch|ăn\s*sáng|an\s*sang|ăn\s*trưa|an\s*trua|ăn\s*tối|an\s*toi|\bB\s*\/\s*L|\bL\s*\/\s*D|\bB\s*&\s*L|meals?\s*:/i.test(s);
  if(/Meal:\s*no\b|no\s*meals?|không\s*(có\s*)?(bữa|ăn)|khong\s*(co\s*)?(bua|an)|Meal:\s*no\b/i.test(s) && !hasFoodToken) return 'No meals';
  // cruise-style counts
  const b=(s.match(/(\d+)\s*breakfasts?/i)||s.match(/(\d+)\s*ăn\s*sáng/i)||[])[1];
  const l=(s.match(/(\d+)\s*lunches?/i)||s.match(/(\d+)\s*ăn\s*trưa/i)||s.match(/Meal:\s*(\d+)\s*LUNCH/i)||[])[1];
  const d=(s.match(/(\d+)\s*dinners?/i)||s.match(/(\d+)\s*ăn\s*tối/i)||[])[1];
  if(b||l||d){
    const parts=[];
    if(b) parts.push(b==='1'?'1 breakfast':`${b} breakfasts`);
    if(l) parts.push(l==='1'?'1 lunch':`${l} lunches`);
    if(d) parts.push(d==='1'?'1 dinner':`${d} dinners`);
    return parts.join(', ');
  }
  if(/\bB\s*\/\s*L\s*\/\s*D\b/i.test(s)) return 'breakfast, lunch, dinner';
  if(/\bB\s*\/\s*L\b/i.test(s)) return 'breakfast, lunch';
  if(/\bL\s*\/\s*D\b/i.test(s)) return 'lunch, dinner';
  const parts=[];
  if(/breakfast|ăn\s*sáng|an\s*sang/i.test(s)) parts.push('breakfast');
  if(/lunch|ăn\s*trưa|an\s*trua/i.test(s)) parts.push('lunch');
  if(/dinner|ăn\s*tối|an\s*toi/i.test(s)) parts.push('dinner');
  if(parts.length) return parts.join(', ');
  // Meal: 1 LUNCH style already handled; bare "Meal: lunch"
  const m=s.match(/Meals?\s*[:\-]\s*([^\n,;]+)/i);
  if(m){
    const v=m[1].trim();
    if(/^no\b/i.test(v)) return 'No meals';
    return v;
  }
  return 'No meals';
}
function sheetMatrix(wb, name){
  const ws=wb.Sheets[name]; if(!ws) return [];
  return XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false});
}
function findRow(mat, pred, start=0){
  for(let i=start;i<mat.length;i++) if(pred(mat[i],i)) return i;
  return -1;
}
function cell(row,i){return String((row&&row[i])||'').trim();}

function loadHotels(){
  return fetch('hotels.json').then(r=>r.json()).then(base=>{
    let extra=[]; try{extra=JSON.parse(localStorage.hotelDir||'[]')}catch(e){}
    return mergeHotels(base, extra);
  });
}
function mergeHotels(a,b){
  const out=a.map(x=>({...x}));
  for(const h of b){
    const i=out.findIndex(x=>norm(x.name)===norm(h.name)||(h.id&&x.id===h.id));
    if(i>=0) out[i]={...out[i],...h}; else out.push(h);
  }
  return out;
}
function saveHotelExtra(h){
  let extra=[]; try{extra=JSON.parse(localStorage.hotelDir||'[]')}catch(e){}
  const i=extra.findIndex(x=>norm(x.name)===norm(h.name));
  if(i>=0) extra[i]=h; else extra.push(h);
  localStorage.hotelDir=JSON.stringify(extra);
}
function matchHotel(hotels, name){
  const n=norm(name); if(!n) return null;
  let best=null, score=0;
  for(const h of hotels){
    const keys=[h.name,...(h.aliases||[])].map(norm);
    for(const k of keys){
      if(!k) continue;
      let s=0;
      if(k===n) s=100;
      else if(n.includes(k)||k.includes(n)) s=80;
      else {
        const at=new Set(k.split(' ')), bt=n.split(' ');
        const inter=bt.filter(t=>at.has(t)&&t.length>2).length;
        s=inter*15;
      }
      if(s>score){score=s; best=h;}
    }
  }
  return score>=50?best:null;
}
function hotelBlock(h){
  if(!h) return '';
  const lines=[h.name];
  const addr=shortenAddress(h.address||'');
  if(addr) lines.push(addr);
  if(h.phone) lines.push((/telefon/i.test(h.phone)?'':'Tel: ')+h.phone.replace(/^Tel(efon)?\s*:\s*/i,''));
  if(h.website) lines.push('Web: '+h.website);
  return lines.join('\n');
}
function resolvePlace(label, hotels, hotelByCity){
  const n=norm(label);
  // "HAN hotel" -> city map
  const cityMap={
    'HAN HOTEL':'HA NOI','HANOI HOTEL':'HA NOI','HA NOI HOTEL':'HA NOI',
    'HA LONG BAY CRUISE':'HA LONG BAY','HALONG BAY CRUISE':'HA LONG BAY',
    'NINH BINH HOTEL':'NINH BINH','HUE HOTEL':'HUE','HOI AN BEACH':'HOI AN','HOI AN HOTEL':'HOI AN',
    'SGN HOTEL':'SAI GON','SAI GON HOTEL':'SAI GON','HCMC HOTEL':'SAI GON',
    'CAT TIEN HOTEL':'CAT TIEN','CAT TIEN':'CAT TIEN','CATTIEN':'CAT TIEN',
    'PHU QUOC HOTEL':'PHU QUOC','PHU QUOC':'PHU QUOC','PQC HOTEL':'PHU QUOC','PQ HOTEL':'PHU QUOC',
    'DA NANG HOTEL':'DA NANG','NHA TRANG HOTEL':'NHA TRANG'
  };
  for(const [k,city] of Object.entries(cityMap)){
    if(n===k||n.includes(k)){const h=hotelByCity[city]; if(h) return h;}
  }
  const cc=canonicalCity(label);
  if(cc && hotelByCity[cc]) return hotelByCity[cc];
  return matchHotel(hotels, label);
}


function headerIndexMap(row){
  const map={};
  (row||[]).forEach((v,i)=>{
    const k=norm(v).replace(/\s+/g,' ');
    if(!k) return;
    if(/^(NO|STT|NR|ROW|#)$/.test(k)) map.no=i;
    else if(/^(NAME|FIRST|FIRST NAME|GIVEN NAME|JMENO|JMENO)$/.test(k)) map.first=i;
    else if(/^(SURNAME|LAST|LAST NAME|FAMILY|PRIJMENI|PRIJMENI)$/.test(k)) map.surname=i;
    else if(/^(SEX|GENDER|TITLE|POHLAVI|POHLAVI)$/.test(k)) map.sex=i;
  });
  return map;
}
function looksLikeRowNo(v){
  return /^\d{1,3}(?:\.0+)?$/.test(String(v||'').trim());
}
function titleFromSex(sex){
  const sx=String(sex||'').trim().toUpperCase().replace(/\./g,'');
  if(sx==='F'||sx==='W'||sx==='MS'||sx==='MISS'||sx==='FEMALE') return 'MS';
  if(sx==='MRS') return 'MRS';
  if(sx==='M'||sx==='MR'||sx==='MALE') return 'MR';
  return 'MR';
}
function formatGuestLine(n,g){
  const title=g.title||titleFromSex(g.sex);
  const sur=String(g.surname||'').trim().toUpperCase();
  const first=String(g.first||'').trim().toUpperCase();
  // Drop accidental numeric tokens (row no leaked into a name field)
  const clean=s=>s.split(/\s+/).filter(tok=>!looksLikeRowNo(tok) && !/^\d+$/.test(tok)).join(' ');
  const name=[clean(sur),clean(first)].filter(Boolean).join(' ');
  return `${n}. ${title} ${name}`.replace(/\s+/g,' ').trim();
}
function parseGuestsFromGt(blob){
  const out=[];
  const re=/(\d+)\s*\.\s*([A-Za-zÁ-žA-Za-z]+)\s*\/\s*([A-Za-zÁ-žA-Za-z]+)(?:\s+(MR|MS|MRS)\.?\b)?/gi;
  let m; const s=String(blob||'');
  while((m=re.exec(s))){
    const title=(m[4]||'').toUpperCase();
    out.push({first:m[3], surname:m[2], sex:title||'', title});
  }
  return out;
}
function parseGuestsFromClients(clients){
  const out=[];
  if(!clients||!clients.length) return out;
  let start=0;
  let cols=headerIndexMap(clients[0]);
  const hasHeader=cols.first!=null||cols.surname!=null||cols.sex!=null||cols.no!=null;
  if(hasHeader){
    start=1;
    if(cols.first==null && cols.no!=null) cols.first=cols.no+1;
    if(cols.surname==null && cols.first!=null) cols.surname=cols.first+1;
    if(cols.sex==null && cols.surname!=null) cols.sex=cols.surname+1;
  } else if(looksLikeRowNo(cell(clients[0],0)) && cell(clients[0],1) && cell(clients[0],2)){
    cols={no:0, first:1, surname:2, sex:3};
  } else {
    cols={first:0, surname:1, sex:2};
  }
  if(cols.first==null) cols.first=0;
  if(cols.surname==null) cols.surname=1;
  if(cols.sex==null) cols.sex=2;

  for(let i=start;i<clients.length;i++){
    const row=clients[i];
    let first=cell(row, cols.first);
    let surname=cell(row, cols.surname);
    let sex=cell(row, cols.sex);
    // Safety: NO column still mapped into first
    if(looksLikeRowNo(first) && cell(row, cols.first+1)){
      first=cell(row, cols.first+1);
      surname=cell(row, cols.first+2);
      sex=cell(row, cols.first+3);
    }
    if(!first && !surname) break;
    if(/platby|payment|^mail$|^notes?$/i.test(first)) break;
    if(/^(name|surname|sex|no|stt)$/i.test(first)) continue;
    let title='';
    const sxU=String(sex||'').trim().toUpperCase().replace(/\./g,'');
    if(/^(MR|MS|MRS)$/.test(sxU)) title=sxU;
    out.push({first:first.trim(), surname:surname.trim(), sex, title});
  }
  return out;
}

function parseSummary(wb, hotels){
  const flags=[];
  const sum=sheetMatrix(wb,'Summary');
  const clients=sheetMatrix(wb,'Clients + Rooming list');
  // guests from Clients sheet (layouts vary: Name|Surname|Sex OR No|Name|Surname|Sex|...)
  const guests=parseGuestsFromClients(clients);
  const gtBlob=[0,1,2].flatMap(r=>[cell(sum[r],0),cell(sum[r],1)]).filter(Boolean).join(' ');
  if(!guests.length){
    for(const g of parseGuestsFromGt(gtBlob)) guests.push(g);
  } else if(gtBlob){
    // Prefer MR/MS/MRS from GT when surname matches (Clients often only has M/F)
    const gt=parseGuestsFromGt(gtBlob);
    for(const g of guests){
      const hit=gt.find(x=>(norm(x.surname)===norm(g.surname)) || (norm(x.first)===norm(g.first)));
      if(hit && hit.title) g.title=hit.title;
    }
  }
  const surname=guests[0]?.surname||'';
  // voucher line: "1. MR SURNAME FIRST" (full name, no stray index)
  const gl=guests.map((g,i)=>formatGuestLine(i+1,g));

  // header fields
  let contact='';
  for(const row of sum.slice(0,15)){
    const all=row.join(' ');
    const m=all.match(/David[^:]*:\s*(\+?[\d\s]+)/i)||all.match(/contact:\s*([^]+)/i);
    if(/Guide contact/i.test(all)||/David:/i.test(all)){
      const mm=all.match(/David:\s*([\+\d\s]+)/i);
      if(mm) contact=`David Vu, tel. ${mm[1].trim()}`;
    }
  }
  if(!contact) contact='David Vu, tel. ';

  const pax=(()=>{for(const r of sum){if(/SOLD PAX/i.test(cell(r,0))) return cell(r,1)||String(guests.length);} return String(guests.length);})();

  // flights
  const flights=[];
  for(let i=0;i<sum.length;i++){
    const a=cell(sum[i],0), dep=cell(sum[i],2), arr=cell(sum[i],4);
    if(/^[A-Z]{2}\s*\d+|^DE\d+/i.test(a) && dep){
      flights.push({carrier:fmtFlight(a), from:cell(sum[i],1), dep, to:cell(sum[i],3), arr, pnr:cell(sum[i],5)});
    }
  }

  // sections — detect by header labels (order-independent)
  const isAccHdr=(r)=>/^accommodation$/i.test(cell(r,0)) || (/^hotel$/i.test(cell(r,1)) && /check\s*in/i.test(cell(r,2)));
  const isActHdr=(r)=>/^activities$/i.test(cell(r,0)) || (/^name$/i.test(cell(r,1)) && /place/i.test(cell(r,2)) && /date/i.test(cell(r,3)));
  const isTrHdr=(r)=>/^transport$/i.test(cell(r,0)) || (/pick-?up\s*TIME/i.test(cell(r,1)) && /pick-?up\s*PLACE/i.test(cell(r,2)));
  const accStart=findRow(sum,isAccHdr);
  const actStart=findRow(sum,isActHdr);
  const trStart=findRow(sum,isTrHdr);
  function sectionEnd(start){
    if(start<0) return -1;
    let end=sum.length;
    for(const s of [accStart,actStart,trStart]){
      if(s>start && s<end) end=s;
    }
    // also stop at other known headers
    for(let i=start+1;i<end;i++){
      const a=cell(sum[i],0);
      if(/^(accommodation|activities|transport|nett rate|international tickets|checklist)/i.test(a) && i!==start) { end=i; break; }
    }
    return end;
  }

  // accommodations
  const acc_rows=[];
  const hotelByCity={};
  if(accStart>=0){
    const accEnd=sectionEnd(accStart);
    for(let i=accStart+1;i<accEnd;i++){
      const city=cell(sum[i],0), hotel=cell(sum[i],1), cin=cell(sum[i],2), cout=cell(sum[i],3);
      const res=cell(sum[i],4), room=cell(sum[i],5);
      if(!city||/^2\s*pax$/i.test(city)||/^accommodation$/i.test(city)||/^hotel$/i.test(city)) continue;
      if(/check\s*in/i.test(cin) && /^hotel$/i.test(hotel)) continue; // header
      if(!hotel&&!cin) continue;
      const isTrain=/train/i.test(city)||/train/i.test(hotel)||/lotus/i.test(hotel);
      if(isTrain){
        // parse SE19 and times from this row notes (cols)
        const blob=[city,hotel,cin,cout,res,room,cell(sum[i],6),cell(sum[i],7)].join(' ');
        let trainNo=(blob.match(/SE\s*#?\s*(\d+)/i)||[])[1];
        trainNo=trainNo?`SE ${trainNo}`:'SE 19';
        const times=[...blob.matchAll(/(\d{1,2})[h:](\d{2})/g)].map(m=>`${m[1].padStart(2,'0')}h${m[2]}`);
        const depT=times[0]||'', arrT=times[1]||'';
        if(!depT||!arrT) flags.push({level:'warn', msg:`Thiếu giờ tàu (không có trong summary) — ${city}/${hotel}. Điền giờ ${trainNo} trên ô vàng trước khi tạo file.`});
        const cinP=parseDayMon(cin), coutP=parseDayMon(cout);
        let train_text=`${cinP.dayMon||cin} - ${coutP.dayMon||cout} ${hotel||'LOTUS TRAIN'} FROM NINH BINH TO HUE (private coupe)`;
        if(/lotus/i.test(hotel)||/night train/i.test(city))
          train_text=`${cinP.dayMon||cin} - ${coutP.dayMon||cout} LOTUS TRAIN FROM NINH BINH TO HUE (private coupe)\n${trainNo}: ${depT||'??'} – ${arrT||'??'}`;
        // Keep ?? placeholders when times absent so the gap is visible & editable
        acc_rows.push({is_train:true, is_hotel:false, train_text, city:train_text, check_in:'', check_out:'', hotel_block:'', note:'', missing:!depT||!arrT, gaps:(!depT||!arrT)?['train_text']:[], hotel, auto_lookup:false, train_no:trainNo, train_dep:depT, train_arr:arrT});
        continue;
      }
      const cinP=parseDayMon(cin), coutP=parseDayMon(cout);
      let check_out=coutP.dayMon||cout;
      if(coutP.late) check_out=(coutP.dayMon||cout.replace(/-.*/,'').trim())+'\n+ Late check out';
      const h=matchHotel(hotels, hotel);
      const cityKey=canonicalCity(city) || norm(city).replace(/ HOMESTAY$/,'');
      if(!h) flags.push({level:'warn', msg:`Chưa có địa chỉ khách sạn: ${hotel||city}`});
      if(h && cityKey) hotelByCity[cityKey]=h;
      // also index common aliases
      if(h){
        for(const alias of [cityKey, canonicalCity(hotel), norm(hotel)]){
          if(alias) hotelByCity[alias]=h;
        }
      }

      const note=[res, room].filter(Boolean).join('\n\n');
      const displayName=(h&&h.name)||hotel||city;
      acc_rows.push({
        is_hotel:true, is_train:false,
        city: cityKey||norm(city),
        check_in: cinP.dayMon||cin,
        check_out,
        hotel_name: hotel||displayName,
        hotel_block: h?hotelBlock(h):(displayName+(hotel&&hotel!==displayName?'':'')+'\n'),
        note, room, reservation:res,
        missing: !h || !h.address,
        phone: h?.phone||'', address: h?.address||'', website: h?.website||'',
        matched_name: displayName
      });
    }
  }

  // activities map by date
  const activities=[];
  if(actStart>=0){
    const actEnd=sectionEnd(actStart);
    for(let i=actStart+1;i<actEnd;i++){
      const name=cell(sum[i],0); if(!name) continue;
      if(/nett rate/i.test(name)||/^name$/i.test(name)) continue;
      activities.push({
        name, place:cell(sum[i],2), when:cell(sum[i],3), res:cell(sum[i],4), notes:cell(sum[i],5),
        dkey: dateKey(cell(sum[i],3))
      });
    }
  }
  function findActivity(date, kind){
    const dk=dateKey(date) || date.replace(/-/g,'');
    let A=activities.find(a=>a.dkey && a.dkey===dk);
    if(!A) A=activities.find(a=>dateKey(a.when)===dk);
    if(!A) A=activities.find(a=>norm(a.when).replace(/\s/g,'').includes(dk));
    if(!A && kind){
      const tokens=norm(kind).split(' ').filter(t=>t.length>2 && !/TOUR|GUIDE|ESG|CSG|TRF/.test(t)).slice(0,3);
      if(tokens.length) A=activities.find(a=>tokens.every(t=>norm(a.name).includes(t)));
    }
    return A||null;
  }

  // transfers / services from transport block
  const trf_rows=[];
  if(trStart>=0){
    const trEnd=sectionEnd(trStart);
    for(let i=trStart+1;i<trEnd;i++){
      const kind=cell(sum[i],0); if(!kind||/^transport$/i.test(kind)||/^pick-?up/i.test(kind)) continue;
      const pickup=cell(sum[i],1), from=cell(sum[i],2), to=cell(sum[i],3), res=cell(sum[i],4), notes=cell(sum[i],5);
      const dp=parseDayMon(pickup);
      const time=parseTimeFrom(pickup);
      const date=dp.dayMon||'';
      const fromH=resolvePlace(from, hotels, hotelByCity);
      const toH=resolvePlace(to, hotels, hotelByCity);
      const k=norm(kind);
      let label='', where_from='', where_to='', note='', missing=false;

      const fmtHotel=(h,fallback)=>h?`${h.name}\n${shortenAddress(h.address||'')}`:(fallback||'');
      const flightNear=(dir)=>{
        // find flight matching date
        const day=date.replace('-','');
        return flights.find(f=>{
          const blob=dir==='arr'?f.arr:f.dep;
          return norm(blob).includes(norm(date).replace('-','').slice(0,5))||norm(blob).includes(day.slice(0,2)+day.slice(2));
        });
      };

      if(/^TRF/i.test(kind)&&/airport/i.test(from)&&!/airport/i.test(to)){
        label='TRANSFER FROM AIRPORT';
        const code=aptCode(from);
        let fll=flights.find(f=>aptCode(f.to)===code && dateKey(f.arr)===dateKey(date||dp.raw));
        if(!fll) fll=flights.find(f=>aptCode(f.to)===code);
        const aptName=code==='HAN'?'HANOI AIRPORT':code==='DAD'?'DAD AIRPORT':code==='SGN'?'SGN AIRPORT':code==='PQC'?'PQC AIRPORT':(code?code+' AIRPORT':from.toUpperCase());
        where_from=`TIME: ${time||''}\n${aptName}`+(fll?`\nFlight: ${fll.carrier}\nArrival: ${time||parseTimeFrom(fll.arr)}`:'');
        where_to=fmtHotel(toH, to);
        const pick=airportArrivalNote(code, fll, surname);
        note=pick.note;
        if(pick.unknown || !code){ missing=true; }
        if(!toH) missing=true;
      } else if(/^TRF/i.test(kind)&&/airport/i.test(to)){
        const code=aptCode(to);
        label=(code==='SGN' && /19:30|23:00/i.test(pickup+time))?'TRANSPORT TO\nAIRPORT':'TRANSFER TO AIRPORT';
        where_from=`TIME: ${time||''}\n`+fmtHotel(fromH, from);
        let fl2=flights.find(f=>aptCode(f.from)===code && dateKey(f.dep)===dateKey(date||dp.raw));
        if(!fl2) fl2=flights.find(f=>aptCode(f.from)===code);
        const aptName=code==='DAD'?'DAD AIRPORT':code==='SGN'?'SGN AIRPORT':code==='PQC'?'PQC AIRPORT':code==='HAN'?'HANOI AIRPORT':(code?code+' AIRPORT':to.toUpperCase());
        where_to=`${aptName}`+(fl2?`\nFlight: ${fl2.carrier}\nDeparture: ${parseTimeFrom(fl2.dep)}`:'');
        const drop=airportDropoffNote(code, fl2);
        note=drop.note;
        if(drop.unknown || !code){ missing=true; }
        if(!fromH) missing=true;
      } else if(/^TRF/i.test(kind)&&/train/i.test(to)){
        label='TRANSFER TO TRAIN';
        where_from=`TIME: ${time||''}\n`+fmtHotel(fromH, from);
        // find train info from acc train row
        const tr=acc_rows.find(r=>r.is_train);
        const tm=(tr&&tr.train_text.match(/SE\s*\d+:\s*(\d{1,2}h\d{2})/i)||[])[1];
        const depClock=tm?tm.replace('h',':'):'';
        where_to=`NINH BINH TRAIN\nTrain: SE 19`+(depClock?`\nDeparture: ${depClock}`:'');
        note='The driver will pick you up at the hotel reception and drop you off at the Ninh Binh train.';
        if(!fromH) missing=true;
      } else if(/^TRF/i.test(kind)&&/train/i.test(from)){
        label='TRANSFER FROM TRAIN';
        const tr=acc_rows.find(r=>r.is_train);
        const tm=(tr&&tr.train_text.match(/–\s*(\d{1,2}h\d{2})/i)||tr&&tr.train_text.match(/-\s*(\d{1,2}h\d{2})/i)||[])[1];
        const arrClock=tm?tm.replace('h',':'):(time||'');
        where_from=`TIME: ${arrClock||time||''}\nHUE TRAIN\nTrain: SE 19\nArrival: ${arrClock||time||''}`;
        where_to=fmtHotel(toH, to);
        note=`The driver will pick you up at the Hue train(outside the arrivals hall) with a sign:\n\n GO2/ MR. ${surname.toUpperCase()}`;
        if(!toH) missing=true;
      } else if(/^TRF/i.test(kind)||/^TRF TO/i.test(kind)){
        // generic transfer
        if(/HOI AN|MM/i.test(kind)){
          label='TRANSFER TO\nHOI AN\n+\nMARBLE MOUNTAIN';
          note='The driver will pick you up at the hotel reception and drop you off at the Hoi An hotel.\n—-----------------------------------\nOn the way you will stop at Marble Mountains for sightseeing, entrance fees included.';
        } else if(/NINH BINH/i.test(to)||/NINH BINH/i.test(kind)){
          label='TRANSFER TO\nNINH BINH';
          note='The driver will pick you up at the cruise harbour and drop you off at the Ninh Binh hotel.';
        } else {
          label=kind;
          note='The driver will pick you up at the hotel reception.';
        }
        where_from=`TIME: ${time||''}\n`+fmtHotel(fromH, from);
        where_to=fmtHotel(toH, to);
        if(!fromH||!toH) missing=true;
      } else {
        // tour / activity — match by date first, then name tokens
        const A=findActivity(date||dp.raw, kind);
        const cz=/CZ|CSG|Czech/i.test(kind+(A?.notes||'')+(A?.name||''));
        label=kind.replace(/\s*\+\s*/g,'\n+');
        if(cz&&!/\+CZ/i.test(label)) label=label+'\n+CZ GUIDE';
        if(/HA NOI CITY|FOOD TOUR/i.test(kind)&&/HAN|HA NOI/i.test(from+to)) label='HAN CITY TOUR & FOOD TOUR\n+CZ GUIDE';
        if(/HA LONG/i.test(kind)) label='HALONG BAY';
        if(/NINH BINH TOUR/i.test(kind)) label='NINH BINH TOUR';
        if(/HUE TOUR/i.test(kind)) label='HUE TOUR';
        if(/HOI AN TOUR/i.test(kind)) label='HOI AN TOUR\n+CZ GUIDE';
        if(/FOOD TOUR BY MOTORBIKE/i.test(kind)) label='FOOD TOUR BY MOTORBIKE';
        if(/CU CHI/i.test(kind)) label='CU CHI TUNNEL TOUR\n+CZ GUIDE';
        if(/CAT\s*TIEN|CATTIEN/i.test(kind)) label=label.includes('CAT')?label:'CAT TIEN TOUR';
        if(/PHU\s*QUOC|\bPQ\b/i.test(kind)) label=label;

        where_from=`TIME: ${time||''}\n`+fmtHotel(fromH, from);
        where_to=fmtHotel(toH, to);
        const an=A?.name||kind;
        const actNotes=A?.notes||'';
        const mealBlob=[actNotes, notes, an, kind].join(' | ');
        const meals=parseMeals(mealBlob);
        const guideLang=/Czech|CSG|\bCZ\b/i.test(actNotes+kind+an)?'Czech speaking':/English|ESG/i.test(actNotes+an+kind)?'English speaking':'English speaking';
        let desc=an;
        if(/HA NOI.*FOOD/i.test(an+kind)) desc='1 Day Hanoi walking tour & Food tour after which you return to the hotel.';
        else if(/HA LONG/i.test(an+kind)) desc='2 Days 1 night HaLong Bay sightseeing cruise, after which you transfer to Ninh Binh hotel.';
        else if(/NINH BINH/i.test(an+kind)) desc='1 Day Ninh Binh sightseeing tour, after which you return to the Ninh Binh hotel.';
        else if(/HUE/i.test(an+kind)) desc='½ Day Hue walking tour, after which you return to the Hue hotel.';
        else if(/HOI AN/i.test(an+kind)) desc='½ Day Hoi An walking tour, after which you return to the Hoi An hotel.';
        else if(/MOTORBIKE|FOOD TOUR BY MOTOR/i.test(an+kind)) desc='Sai Gon food tour by motorbike, after which you return to the Sai Gon hotel.';
        else if(/CU CHI/i.test(an+kind)) desc='½ Day Cu Chi sightseeing on tour, after which you return to the Sai Gon hotel.';
        else if(/CAT\s*TIEN|CATTIEN/i.test(an+kind)) desc='1 Day Cat Tien sightseeing tour, after which you return to the Cat Tien hotel.';
        else if(/PHU\s*QUOC|\bPQ\b/i.test(an+kind)) desc='1 Day Phu Quoc sightseeing tour, after which you return to the Phu Quoc hotel.';
        else {
          // generic from activity name
          desc=(an.replace(/,\s*(ESG|CSG)\s*$/i,'').trim()) + ', after which you return to the hotel.';
        }

        if(/HA LONG/i.test(kind)){
          note=`${desc}\n-------------------------------------\nThe representative of the cruise will pick you up at the hotel reception.\n\nMeals on tour: ${meals}`;
        } else {
          const pay=/food tour by motorbike/i.test(kind)?'pay for street food':'pay for all entry fees'+(/(FOOD|street food)/i.test(actNotes+kind)?' & foods':'');
          note=`${desc}\n-------------------------------------\nThe ${guideLang} guide will pick you up at the hotel reception and ${pay}.\n\nMeals on tour: ${meals}`;
        }
        if(!fromH) missing=true;
      }
      // Append free-text notes from summary transport col (e.g. sim card, extra stops)
      const sumNotes=String(notes||'').trim();
      if(sumNotes){
        // Skip internal ops notes that are not for the guest voucher
        if(!/under\s*100\s*kg|tall\s*\(|:\s*done\b/i.test(sumNotes)){
          const extra=sumNotes.replace(/\s*\(should be in the voucher\)\s*/i,'').trim();
          if(extra){
            if(/stops?/i.test(extra) && /marble|budha|buddha|dragon/i.test(extra)){
              // Expand marble-mountain boilerplate with named stops from summary
              const stopBits=[];
              if(/marble/i.test(extra)) stopBits.push('Marble Mountains');
              if(/budh|buddha/i.test(extra)) stopBits.push('Lady Buddha');
              if(/dragon/i.test(extra)) stopBits.push('Dragon Bridge');
              if(stopBits.length){
                note=note.replace(/On the way you will stop at Marble Mountains for sightseeing, entrance fees included\.?/,
                  'On the way you will stop at '+stopBits.join(', ')+' for sightseeing, entrance fees included.');
                if(!/On the way you will stop/i.test(note)){
                  note=note.replace(/\n?—?-*\n?On the way[\s\S]*/,'') + '\n—-----------------------------------\nOn the way you will stop at '+stopBits.join(', ')+' for sightseeing, entrance fees included.';
                }
              }
            } else if(/sim\s*card/i.test(extra)){
              if(!/sim\s*card/i.test(note)) note = note + '\n\nNote: ' + extra.replace(/^plus\s+/i,'');
            } else if(!note.includes(extra)){
              note = note + (note?'\n':'') + extra;
            }
          }
        }
      }
      const gaps=[];
      if(/TRAIN/i.test(where_to) && !/Departure:/i.test(where_to)){ gaps.push('where_to'); missing=true; }
      if(/TRAIN/i.test(where_from) && !/Arrival:\s*\d/i.test(where_from)){ gaps.push('where_from'); missing=true; }
      if(/Flight:\s*$/m.test(where_from) || (/Flight:/i.test(where_from) && !/Flight:\s*\S/i.test(where_from))){ gaps.push('where_from'); missing=true; }
      if(/Flight:\s*$/m.test(where_to) || (/Flight:/i.test(where_to) && !/Flight:\s*\S/i.test(where_to))){ gaps.push('where_to'); missing=true; }
      if(!time && /TIME:\s*$/m.test(where_from)){ gaps.push('where_from'); missing=true; }
      if(!note) { gaps.push('note'); missing=true; }
      if(/airport arrivals hall with a sign/i.test(note) || /drop you off at the airport\.$/i.test(note)){
        gaps.push('note'); missing=true;
        flags.push({level:'warn', msg:`Thiếu thông tin nhà ga/cổng đón cho "${(label||'').replace(/\n/g,' ')}" — điền ô Note (nội địa/quốc tế + pillar).`});
      }
      trf_rows.push({label, date, where_from, where_to, note, missing, gaps, raw_kind:kind, time, summary_notes:sumNotes});
    }
  }

  // Extra flags for fields absent from summary
  for(const r of trf_rows){
    if(r.gaps && r.gaps.length){
      if(r.gaps.includes('where_to') && /TRAIN/i.test(r.where_to))
        flags.push({level:'warn', msg:`Thiếu giờ tàu ở "${(r.label||'').replace(/\n/g,' ')}" (Departure) — không có trong summary, điền ô vàng.`});
    }
  }
  for(const r of acc_rows){
    if(r.is_train && r.missing)
      flags.push({level:'warn', msg:`Dòng tàu còn ?? — thay bằng giờ thật trước khi in.`});
  }

  return {
    update_date: todayVN(),
    payment: 'ALL FULLY PREPAID BY CK GO2',
    guests: gl,
    guest_lines: gl.join('\n'),
    pax: String(pax).replace(/\.0$/,''),
    contact,
    surname,
    acc_rows,
    trf_rows,
    flags,
    hotels,
    hotelByCity
  };
}

async function renderVoucherDoc(templateName, data){
  const buf=await (await fetch('templates/'+templateName)).arrayBuffer();
  const zip=new PizZip(buf);
  const doc=new docxtemplater(zip,{paragraphLoop:true,linebreaks:true,delimiters:{start:'{',end:'}'},nullGetter:()=>''});
  doc.render(data);
  // Docxtemplater leaves empty conditional table rows (false is_hdr/is_data/is_train). Strip them.
  try{
    const z=doc.getZip();
    const xml=z.file('word/document.xml').asText();
    let cleaned=xml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row)=>{
      const texts=[...row.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map(m=>m[1]).join('');
      const visible=texts.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#0?39;/g,"'").trim();
      return visible ? row : '';
    });
    // Justified guest lines create huge gaps between short words — force left align
    cleaned=cleaned.replace(/<w:jc\b[^>]*w:val="both"[^\/]*\/>/g, '<w:jc w:val="left"/>');
    // Force Arial on every run so placeholders / loop rows / soft-breaks match sample
    const RF='<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/>';
    cleaned=cleaned.replace(/<w:r\b([^>]*)>([\s\S]*?)<\/w:r>/g, (full, attrs, body)=>{
      if(/<w:drawing[\s>]|<w:pict[\s>]|<w:object[\s>]/.test(body)) return full; // keep images
      if(/<w:rPr[\s>]/.test(body)){
        body=body.replace(/<w:rPr\b([^>]*)>([\s\S]*?)<\/w:rPr>/, (m,a,inner)=>{
          if(/<w:rFonts\b/.test(inner))
            inner=inner.replace(/<w:rFonts\b[^\/]*\/>|<w:rFonts\b[^>]*>[\s\S]*?<\/w:rFonts>/, RF);
          else
            inner=RF+inner;
          return '<w:rPr'+a+'>'+inner+'</w:rPr>';
        });
      } else {
        body='<w:rPr>'+RF+'<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'+body;
      }
      return '<w:r'+attrs+'>'+body+'</w:r>';
    });
    // Also normalize styles docDefaults if present in zip
    try{
      const styles=z.file('word/styles.xml') && z.file('word/styles.xml').asText();
      if(styles){
        let st=styles;
        if(/<w:rFonts\b/.test(st)){
          st=st.replace(/<w:rFonts\b[^\/]*\/>/g, RF);
        }
        z.file('word/styles.xml', st);
      }
    }catch(_e){}
    z.file('word/document.xml', cleaned);
    return z.generate({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',compression:'DEFLATE'});
  }catch(e){
    console.warn('row cleanup failed', e);
    return doc.getZip().generate({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',compression:'DEFLATE'});
  }
}



function hotelCacheGet(){ try{return JSON.parse(localStorage.hotelCache||'{}')}catch(e){return {}} }
function hotelCacheSet(key, hotel){
  const c=hotelCacheGet(); c[norm(key)]=hotel; localStorage.hotelCache=JSON.stringify(c);
}
function cacheLookup(query){
  const c=hotelCacheGet(); const n=norm(query);
  if(c[n]) return c[n];
  for(const [k,v] of Object.entries(c)){
    if(n.includes(k)||k.includes(n)||norm(v.name||'').includes(n)||n.includes(norm(v.name||''))) return v;
  }
  return null;
}

async function serperHotelLookup(query, city){
  const cfg=window.DOCGEN_CONFIG||{};
  const key=cfg.SERPER_API_KEY||'';
  if(!key) return {ok:false, reason:'no_key'};
  const q=[query, city, 'Vietnam'].filter(Boolean).join(' ').replace(/\s+/g,' ').trim();
  try{
    const r=await fetch('https://google.serper.dev/places',{
      method:'POST',
      headers:{'X-API-KEY':key,'Content-Type':'application/json'},
      body:JSON.stringify({q, gl:'vn', hl:'en'})
    });
    if(r.status===402||r.status===403||r.status===429) return {ok:false, reason:'quota', status:r.status};
    if(!r.ok) return {ok:false, reason:'http', status:r.status};
    const data=await r.json();
    const places=data.places||[];
    if(!places.length) return {ok:false, reason:'empty'};
    // pick best fuzzy match
    let best=places[0], score=0;
    const nq=norm(query);
    for(const p of places){
      const nt=norm(p.title||'');
      let s=0;
      if(nt===nq) s=100;
      else if(nt.includes(nq)||nq.includes(nt)) s=80;
      else {
        const at=new Set(nq.split(' ').filter(t=>t.length>2));
        s=[...at].filter(t=>nt.includes(t)).length*20;
      }
      if(s>score){score=s; best=p;}
    }
    // Reject weak / unrelated Places hits (e.g. garbage query matching random hotel)
    if(score < 40) return {ok:false, reason:'low_score', score};
    const hotel={
      id:'serper-'+Date.now(),
      aliases:[query, best.title].filter(Boolean),
      name:String(best.title||query).trim(),
      address:shortenAddress(String(best.address||'').trim()),
      phone:String(best.phoneNumber||best.phone||'').trim(),
      website:String(best.website||'').trim(),
      source:'serper'
    };
    if(!hotel.address && !hotel.phone) return {ok:false, reason:'empty'};
    return {ok:true, hotel, score};
  }catch(e){
    console.warn('Serper lookup failed', e);
    // network/CORS -> treat as fail, may fall back
    return {ok:false, reason:'network', error:String(e)};
  }
}

async function geminiHotelLookup(query){
  const cfg=window.DOCGEN_CONFIG||{};
  const key=cfg.GEMINI_API_KEY||'';
  if(!key) return null;
  const model=cfg.GEMINI_MODEL||'gemini-3.5-flash-lite';
  const prompt='Find the official full hotel or cruise ship contact details for: "'+query+'" in Vietnam.\n'
    +'Return ONLY a JSON object with keys: name, address, phone, website.\n'
    +'Use the most official/current public listing. If unknown, use empty string. No markdown.';
  async function call(useSearch){
    const body={contents:[{parts:[{text:prompt}]}], generationConfig:{temperature:0.2}};
    if(useSearch) body.tools=[{google_search:{}}];
    const r=await fetch('https://generativelanguage.googleapis.com/v1beta/models/'+model+':generateContent?key='+encodeURIComponent(key),{
      method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)});
    const textRaw=await r.text();
    if(!r.ok){ const err=new Error('HTTP '+r.status); err.status=r.status; err.body=textRaw; throw err; }
    const j=JSON.parse(textRaw);
    const text=(j.candidates||[]).flatMap(c=>(c.content&&c.content.parts||[]).map(p=>p.text||'')).join('\n');
    const m=text.match(/\{[\s\S]*\}/);
    if(!m) return null;
    const obj=JSON.parse(m[0]);
    if(!obj.name && !obj.address) return null;
    return {id:'auto-'+Date.now(), aliases:[query, obj.name].filter(Boolean),
      name:String(obj.name||query).trim(), address:shortenAddress(String(obj.address||'').trim()),
      phone:String(obj.phone||'').trim(), website:String(obj.website||'').trim(),
      source: useSearch?'gemini-search':'gemini'};
  }
  try{ return await call(true); }
  catch(e){ console.warn('Gemini grounded lookup failed, trying without search', e.status||e);
    try{ return await call(false); } catch(e2){ console.warn('Gemini lookup failed', e2); return null; } }
}


async function tavilyHotelLookup(query, city){
  const cfg=window.DOCGEN_CONFIG||{};
  const key=cfg.TAVILY_API_KEY||'';
  if(!key) return {ok:false, reason:'no_key'};
  const q=[query, city, 'Vietnam', 'address', 'phone'].filter(Boolean).join(' ').replace(/\s+/g,' ').trim();
  try{
    const r=await fetch('https://api.tavily.com/search',{
      method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({api_key:key, query:q, include_answer:true, search_depth:'basic', max_results:5})
    });
    if(r.status===402||r.status===403||r.status===429) return {ok:false, reason:'quota', status:r.status};
    if(!r.ok) return {ok:false, reason:'http', status:r.status};
    const data=await r.json();
    const blob=[data.answer||'', ...((data.results||[]).map(x=>((x.title||'')+'\n'+(x.content||'')+'\n'+(x.url||''))))].join('\n');
    const phone=(blob.match(/(\+?\d[\d\s().-]{7,}\d)/)||[])[1]||'';
    // prefer lines with address-like content
    let address='';
    const addrM=blob.match(/Address\s*[:\-]\s*([^\n]+)/i) || blob.match(/Địa chỉ\s*[:\-]\s*([^\n]+)/i);
    if(addrM) address=addrM[1].trim();
    if(!address){
      const line=(blob.split(/\n/).find(l=>/,.*,/.test(l)&&( /Vietnam|Viet Nam|Hanoi|Hội An|Hoi An|Hue|Saigon|HCMC|Quảng|Street|St\.|Road/i.test(l)))||'');
      address=line.replace(/^[^:]*:\s*/,'').trim().slice(0,180);
    }
    let website='';
    for(const res of (data.results||[])){
      const u=res.url||'';
      if(u && !/tripadvisor|booking\.com|agoda|expedia|google\.|facebook|wikipedia/i.test(u)){ website=u; break; }
    }
    const name=query;
    if(!address && !phone) return {ok:false, reason:'empty'};
    // Reject non-answers / weak matches (query echoed in "no info" text, etc.)
    if(/do not provide|no (information|results|address)|not found|could not find|unknown hotel/i.test(blob)){
      return {ok:false, reason:'low_score'};
    }
    const tokens=norm(query).split(' ').filter(t=>t.length>=4);
    const titles=norm(((data.results||[]).map(x=>(x.title||'')+' '+(x.content||'').slice(0,200))).join(' '));
    const addrN=norm(address);
    // Prefer token hit in address or result titles — not merely in the echoed query/answer
    if(tokens.length){
      const hit=tokens.some(tok=>addrN.includes(tok) || titles.includes(tok));
      if(!hit) return {ok:false, reason:'low_score'};
      // Single long nonsense token (e.g. ZZZNOTAREAL...) that only appears as the query itself
      if(tokens.length===1 && tokens[0].length>=16 && !addrN.includes(tokens[0].slice(0,8))){
        return {ok:false, reason:'low_score'};
      }
    }
    const hotel={id:'tavily-'+Date.now(), aliases:[query], name, address:shortenAddress(address), phone: (phone||'').trim(), website, source:'tavily'};
    return {ok:true, hotel};
  }catch(e){
    console.warn('Tavily lookup failed', e);
    return {ok:false, reason:'network', error:String(e)};
  }
}

function sanitizeHotelCandidate(h){
  if(!h) return null;
  const addr=String(h.address||'').trim();
  if(looksLikeProseAddress(addr)){
    return {...h, address:'', _bad_address:true};
  }
  return {...h, address:shortenAddress(addr)};
}
async function resolveHotelContact(query, city){
  // 1) local cache
  const cached=cacheLookup(query) || (city?cacheLookup(city+' '+query):null);
  if(cached && (cached.address||cached.phone)){
    const h=sanitizeHotelCandidate({...cached, from_cache:true});
    if(h && (h.address||h.phone)) return h;
  }

  // 2) Serper (1 call)
  const s=await serperHotelLookup(query, city);
  if(s.ok){
    const h=sanitizeHotelCandidate(s.hotel);
    if(h && (h.address||h.phone)){
      hotelCacheSet(query, h);
      if(h.name) hotelCacheSet(h.name, h);
      return h;
    }
  }
  // 3) Tavily if Serper failed
  const tv=await tavilyHotelLookup(query, city);
  if(tv.ok){
    const h=sanitizeHotelCandidate(tv.hotel);
    if(h && (h.address||h.phone)){
      hotelCacheSet(query, h);
      if(h.name) hotelCacheSet(h.name, h);
      return h;
    }
  }
  // 4) Gemini last resort — only when search APIs failed (quota/network/auth),
  // not when they returned unrelated/empty hits for a garbage name.
  const searchMiss = (s.reason==='low_score'||s.reason==='empty') && (tv.reason==='low_score'||tv.reason==='empty'||tv.reason==='no_key');
  if(searchMiss) return null;
  const g=sanitizeHotelCandidate(await geminiHotelLookup([query, city, 'Vietnam'].filter(Boolean).join(' ')));
  if(g && (g.address||g.phone)){
    hotelCacheSet(query, g);
    return g;
  }
  return null;
}

async function enrichMissingHotels(data, hotels){
  for(const r of data.acc_rows){
    if(!r.is_hotel || r.is_train || !r.missing) continue;
    const q=r.hotel_name||r.city;
    const found=await resolveHotelContact(q, r.city);
    if(found && (found.address||found.phone)){
      // also persist to hotelDir extras for matching next time
      saveHotelExtra(found);
      hotels.push(found);
      r.hotel_block=hotelBlock(found);
      r.address=found.address; r.phone=found.phone; r.website=found.website;
      r.matched_name=found.name; r.missing=false;
      r.source=found.source||'unknown';
      r.auto_lookup=true;
      // Clear the earlier "missing address" warning for this hotel
      data.flags=(data.flags||[]).filter(f=>!(f.level==='warn' && /Chưa có địa chỉ|No hotel address/i.test(f.msg||'') && norm(f.msg).includes(norm(q).slice(0,12))));
      if(found.source==='serper'){
        data.flags.push({level:'info', msg:'Từ tìm kiếm Google (Serper): '+found.name+(found.from_cache?' (cache)':'')});
      } else if(found.source==='tavily'){
        data.flags.push({level:'info', msg:'Từ tìm kiếm web (Tavily): '+found.name});
      } else if(String(found.source||'').startsWith('gemini')){
        data.flags.push({level:'ai', msg:'⚠ Địa chỉ do AI tạo, có thể không chính xác – vui lòng kiểm tra trước khi tải: '+found.name});
      }
    }
  }
  return data;
}


function parseSummaryFromPdfText(text, hotels){
  // Best-effort rebuild of Summary AOA from a printed GO2 summary PDF.
  const lines=String(text||'').split(/\r?\n/).map(l=>l.replace(/\u00a0/g,' ').trim()).filter(Boolean);
  const sum=[];
  // GT line
  const gt=lines.find(l=>/VESELA|LEKES|\/[A-Z]+\s+M[RS]/i.test(l) && /\d+\./.test(l)) || '';
  sum.push(['GT', gt.replace(/^GT\s*/i,'')]);
  const pax=(text.match(/SOLD PAX\s+(\d+)/i)||[])[1]||'';
  if(pax) sum.push(['SOLD PAX', pax]);
  const contact=(text.match(/David:\s*([\+\d\s]+)/i)||[])[1];
  if(contact) sum.push(['Guide contact', 'David: '+contact.trim()]);

  // Flights: CARRIER + airport codes lines
  sum.push([]);
  sum.push(['carrier','destination','departure','destination','arrival','reservation']);
  const flightRe=/\b([A-Z]{2}\s?\d{2,4})\b.*?\b([A-Z]{3})\b/g;
  // From PDF text content objects we often get pairs; also parse GDS blocks
  for(const m of String(text).matchAll(/\b([A-Z]{2})\s?(\d{2,4})\s+[A-Z]\s+(\d{2}[A-Z]{3})\s+\d\s+([A-Z]{3})([A-Z]{3})\s+HK\d/gi)){
    sum.push([m[1]+' '+m[2], m[4], m[3], m[5], '', '']);
  }
  // Fallback simple: EK 140 PRG etc from layout — skip if already have
  if(sum.filter(r=>/^[A-Z]{2}/i.test(String(r[0]||''))).length<2){
    for(const l of lines){
      const m=l.match(/^([A-Z]{2}\s?\d{2,4})\s+([A-Z]{3})\b/);
      if(m) sum.push([m[1], m[2], '', '', '', '']);
    }
  }

  // Transport block — pair "KIND  DATE" lines with later place lines is unreliable;
  // extract KIND + TIME when on same line (common in pdftotext -layout)
  sum.push([]);
  sum.push(['transport','pick-up TIME','pick-up PLACE','drop-off PLACE','reservation','notes ']);
  for(const l of lines){
    const m=l.match(/^(TRF(?:\s+TO\s+[A-Z ]+)?|TRF\s*\+\s*MM|FOOD TOUR|HA NOI TOUR|HA LONG BAY|NINH BINH TOUR|HUE TOUR|HOI AN TOUR|CAT TIEN TOUR|PHU QUOC TOUR[^\d]*)\s+(\d{1,2}[A-Z]{3}[_\s].+)$/i);
    if(m){
      const kind=m[1].trim();
      const pickup=m[2].trim();
      let from='', to='';
      if(/TRAIN/i.test(kind) && /TO TRAIN/i.test(kind)){ from='NINH BINH hotel'; to='NINH BINH train'; }
      else if(/NINH BINH/i.test(kind) && /TRF/i.test(kind)){ from='HA LONG BAY cruise'; to='NINH BINH hotel'; }
      else if(/MM|HOI AN/i.test(kind) && /TRF/i.test(kind)){ from='HUE hotel'; to='HOI AN hotel'; }
      else if(/HA LONG/i.test(kind)){ from='HAN hotel'; to='HA LONG BAY cruise'; }
      else if(/TOUR|FOOD/i.test(kind)){
        const place=canonicalCity(kind)||'HA NOI';
        const hk=place==='HA NOI'?'HAN hotel':place==='NINH BINH'?'NINH BINH hotel':place==='HUE'?'HUE hotel':place==='HOI AN'?'HOI AN hotel':place+' hotel';
        from=hk; to=hk;
      } else if(/^TRF$/i.test(kind)){
        // Match flight arrival/departure on same day
        const dk=dateKey(pickup);
        const arrFl=sum.find(r=>/^[A-Z]{2}/i.test(String(r[0]||'')) && dateKey(r[2]||r[4]||'')===dk);
        // Also scan full text for "HAN airport" near this date
        if(/HAN airport/i.test(text) && /13:15|airport/i.test(pickup+l)){ from='HAN airport'; to='HAN hotel'; }
        else if(/DAD airport/i.test(text) && /20:00|28AUG/i.test(pickup)){ from='HOI AN hotel'; to='DAD airport'; }
        else if(/HUE train/i.test(text) && /09:10/i.test(pickup)){ from='HUE train'; to='HUE hotel'; }
      }
      sum.push([kind, pickup, from, to, '', '']);
    }
  }

  // Accommodation: CITY + HOTEL on one line
  sum.push([]);
  sum.push(['accommodation','hotel','CHECK IN','CHECK OUT','reservation','room type']);
  const cityHotel=[
    [/HA NOI\s+(ORIENTAL[^\n]*)/i,'HA NOI'],
    [/HA LONG BAY\s+(ASPIRA[^\n]*)/i,'HA LONG BAY'],
    [/NINH BINH\s+(LALITA[^\n]*|TAM COC[^\n]*)/i,'NINH BINH'],
    [/NIGHT TRAIN\s+(LOTUS[^\n]*)/i,'NIGHT TRAIN'],
    [/^\s*HUE\s+(CHERISH[^\n]*)/im,'HUE'],
    [/HOI AN\s+(HOI AN BEACH[^\n]*)/i,'HOI AN'],
    [/DA\s*NANG\s+(CHICLAND[^\n]*)/i,'DANANG'],
    [/CAT\s*TI[EÊ]N\s+([^\n]+)/i,'CAT TIEN'],
    [/PH[UÚ]\s*QU[OỐ]C\s+([^\n]+)/i,'PHU QUOC'],
  ];
  const blob=lines.join('\n');
  for(const [re, city] of cityHotel){
    const m=blob.match(re);
    if(m){
      let hotel=m[1].trim().split(/\s{2,}|\t/)[0].trim();
      hotel=hotel.replace(/\s+\d{2}[A-Z]{3}.*$/,'').trim();
      sum.push([city, hotel, '', '', '', '']);
    }
  }

  // Activities
  sum.push([]);
  sum.push(['activities','name','place','date and times','reservation','notes (meals...)']);
  for(const l of lines){
    const m=l.match(/^(\d*\/?\d*D(?:\d*N)?\s+.+?(?:TOUR|CRUISE|FOOD)[^,]*)(?:,\s*(ESG|CSG))?$/i);
    if(m) sum.push([m[0], '', '', '', '', '']);
  }
  // Meal lines sometimes separate — attach raw meal phrases into a synthetic activity notes via leftover
  const mealLines=lines.filter(l=>/Meal|meals|ăn\s|breakfast|lunch|dinner|street food/i.test(l));
  if(mealLines.length && sum.length){
    // append as notes on last activity-ish row
    for(let i=sum.length-1;i>=0;i--){
      if(/TOUR|CRUISE|FOOD/i.test(String(sum[i][0]||''))){ sum[i][5]=mealLines.join(' | '); break; }
    }
  }

  const clients=[['Name','Surname','Sex']];
  const seen=new Set();
  for(const m of String(text).matchAll(/(\d+)\.([A-Z]+)\/([A-Z]+)(?:\s+(MR|MS|MRS|CHD))?/gi)){
    const key=m[1]+'.'+m[2]+'/'+m[3];
    if(seen.has(key)) continue;
    seen.add(key);
    clients.push([m[3], m[2], /MS|MRS/i.test(m[4]||'')?'F':'M']);
  }
  // Build workbook-like object for parseSummary
  const wb={SheetNames:['Summary','Clients + Rooming list'], Sheets:{}};
  wb.Sheets['Summary']=XLSX.utils.aoa_to_sheet(sum);
  wb.Sheets['Clients + Rooming list']=XLSX.utils.aoa_to_sheet(clients.length>1?clients:[['Name','Surname','Sex'],['GUEST','UNKNOWN','M']]);
  return parseSummary(wb, hotels);
}

global.Voucher={parseSummary,parseSummaryFromPdfText,loadHotels,saveHotelExtra,matchHotel,hotelBlock,todayVN,norm,renderVoucherDoc,mergeHotels,geminiHotelLookup,serperHotelLookup,tavilyHotelLookup,resolveHotelContact,enrichMissingHotels,hotelCacheGet,shortenAddress,canonicalCity,parseMeals,aptCode,dateKey,parseGuestsFromClients,parseGuestsFromGt,formatGuestLine};
})(window);
