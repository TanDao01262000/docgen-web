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
  if(h.address) lines.push(h.address);
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
    'SGN HOTEL':'SAI GON','SAI GON HOTEL':'SAI GON','HCMC HOTEL':'SAI GON'
  };
  for(const [k,city] of Object.entries(cityMap)){
    if(n===k||n.includes(k)){const h=hotelByCity[city]; if(h) return h;}
  }
  return matchHotel(hotels, label);
}

function parseSummary(wb, hotels){
  const flags=[];
  const sum=sheetMatrix(wb,'Summary');
  const clients=sheetMatrix(wb,'Clients + Rooming list');
  // guests from clients
  const guests=[];
  for(let i=1;i<clients.length;i++){
    const name=cell(clients[i],0), sur=cell(clients[i],1);
    if(!name&&!sur) break;
    if(/platby/i.test(name)) break;
    guests.push({first:name.trim(), surname:sur.trim(), sex:cell(clients[i],2)});
  }
  if(!guests.length){
    // fallback GT line
    const gt=cell(sum[0],1);
    const parts=[...gt.matchAll(/(\d+)\.([A-Z]+)\/([A-Z]+)\s+(MR|MS|MRS)?/gi)];
    for(const p of parts) guests.push({first:p[3], surname:p[2], sex:p[4]||''});
  }
  const surname=guests[0]?.surname||'';
  const guest_lines=guests.map((g,i)=>`${i+1}. ${(g.sex||'MR').toUpperCase().replace(/M$/,'MR')} ${norm(g.surname)} ${norm(g.first)}`.replace(/\s+/g,' ')).map(s=>s.replace(/ MR /,' MR ').replace(/^(\d+\.)\s*M\s/,'$1 MR '));
  // nicer guest lines: "1. MR LEKES VIKTOR"
  const gl=guests.map((g,i)=>{
    let title='MR';
    const sx=(g.sex||'').toUpperCase();
    if(sx==='F'||sx==='MS'||sx==='MRS') title='MS';
    return `${i+1}. ${title} ${String(g.surname).trim().toUpperCase()} ${String(g.first).trim().toUpperCase()}`;
  });

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
      flights.push({carrier:a.replace(/\s+/g,''), from:cell(sum[i],1), dep, to:cell(sum[i],3), arr, pnr:cell(sum[i],5)});
    }
  }

  // sections
  const accStart=findRow(sum,r=>/^accommodation$/i.test(cell(r,0)));
  const actStart=findRow(sum,r=>/^activities$/i.test(cell(r,0)));
  const trStart=findRow(sum,r=>/^transport$/i.test(cell(r,0)));

  // accommodations
  const acc_rows=[];
  const hotelByCity={};
  if(accStart>=0){
    for(let i=accStart+1;i<sum.length;i++){
      if(actStart>=0&&i>=actStart) break;
      if(/^activities$/i.test(cell(sum[i],0))) break;
      const city=cell(sum[i],0), hotel=cell(sum[i],1), cin=cell(sum[i],2), cout=cell(sum[i],3);
      const res=cell(sum[i],4), room=cell(sum[i],5);
      if(!city||/^2\s*pax$/i.test(city)||/^accommodation$/i.test(city)) continue;
      if(!hotel&&!cin) continue;
      const isTrain=/train/i.test(city)||/train/i.test(hotel)||/lotus/i.test(hotel);
      if(isTrain){
        // parse SE19 and times from this row notes (cols)
        const blob=[city,hotel,cin,cout,res,room,cell(sum[i],6),cell(sum[i],7)].join(' ');
        let trainNo=(blob.match(/SE\s*#?\s*(\d+)/i)||[])[1];
        trainNo=trainNo?`SE ${trainNo}`:'SE 19';
        const times=[...blob.matchAll(/(\d{1,2})[h:](\d{2})/g)].map(m=>`${m[1].padStart(2,'0')}h${m[2]}`);
        const depT=times[0]||'', arrT=times[1]||'';
        if(!depT||!arrT) flags.push({level:'warn', msg:`Thiếu giờ tàu trên dòng ${city}/${hotel} — để trống, mặc định ${trainNo}`});
        const cinP=parseDayMon(cin), coutP=parseDayMon(cout);
        let train_text=`${cinP.dayMon||cin} - ${coutP.dayMon||cout} ${hotel||'LOTUS TRAIN'} FROM NINH BINH TO HUE (private coupe)`;
        // try detect route from places - keep sample style
        if(/lotus/i.test(hotel)||/night train/i.test(city))
          train_text=`${cinP.dayMon||'16-JUN'} - ${coutP.dayMon||'17-JUN'} LOTUS TRAIN FROM NINH BINH TO HUE (private coupe)\n${trainNo}: ${depT||'??'} – ${arrT||'??'}`;
        if(!depT||!arrT) train_text=train_text.replace('?? – ??','').replace(/: – /,'');
        acc_rows.push({is_train:true, is_hotel:false, train_text, city:train_text, check_in:'', check_out:'', hotel_block:'', note:'', missing:!depT||!arrT, hotel, auto_lookup:false});
        continue;
      }
      const cinP=parseDayMon(cin), coutP=parseDayMon(cout);
      let check_out=coutP.dayMon||cout;
      if(coutP.late) check_out=(coutP.dayMon||cout.replace(/-.*/,'').trim())+'\n+ Late check out';
      const h=matchHotel(hotels, hotel);
      if(!h) flags.push({level:'warn', msg:`Chưa có địa chỉ khách sạn: ${hotel}`});
      else hotelByCity[norm(city).replace(' HOMESTAY','').replace('HA NOI','HA NOI')]=h;
      // map city keys
      const cityKey=city.replace(/ HOMESTAY/i,'').trim().toUpperCase();
      if(h) hotelByCity[cityKey]=h;
      if(/HA NOI/i.test(city)) hotelByCity['HA NOI']=h||hotelByCity['HA NOI'];
      if(/HA LONG/i.test(city)) hotelByCity['HA LONG BAY']=h;
      if(/NINH BINH/i.test(city)) hotelByCity['NINH BINH']=h;
      if(/^HUE/i.test(city)) hotelByCity['HUE']=h;
      if(/HOI AN/i.test(city)) hotelByCity['HOI AN']=h;
      if(/SAI GON|SAIGON/i.test(city)) hotelByCity['SAI GON']=h;

      const note=[res, room].filter(Boolean).join('\n\n');
      acc_rows.push({
        is_hotel:true, is_train:false,
        city: cityKey==='SAI GON'?'SAI GON':cityKey,
        check_in: cinP.dayMon||cin,
        check_out,
        hotel_name: hotel,
        hotel_block: h?hotelBlock(h):(hotel+'\n'),
        note, room, reservation:res,
        missing: !h || !h.address,
        phone: h?.phone||'', address: h?.address||'', website: h?.website||'',
        matched_name: h?.name||hotel
      });
    }
  }

  // activities map by date
  const activities=[];
  if(actStart>=0){
    for(let i=actStart+1;i<sum.length;i++){
      const name=cell(sum[i],0); if(!name) continue;
      if(/nett rate/i.test(name)) break;
      activities.push({
        name, place:cell(sum[i],2), when:cell(sum[i],3), res:cell(sum[i],4), notes:cell(sum[i],5)
      });
    }
  }
  function actFor(datePrefix){
    const d=norm(datePrefix).slice(0,5); // 13JUN
    return activities.find(a=>norm(a.when).includes(norm(datePrefix).replace(/\s/g,'').slice(0,5))||norm(a.when).startsWith(d.replace(' ','')));
  }

  // transfers / services from transport block
  const trf_rows=[];
  if(trStart>=0){
    for(let i=trStart+1;i<sum.length;i++){
      if(accStart>=0&&i>=accStart) break;
      const kind=cell(sum[i],0); if(!kind||/^transport$/i.test(kind)) continue;
      const pickup=cell(sum[i],1), from=cell(sum[i],2), to=cell(sum[i],3), res=cell(sum[i],4), notes=cell(sum[i],5);
      const dp=parseDayMon(pickup);
      const time=parseTimeFrom(pickup);
      const date=dp.dayMon||'';
      const fromH=resolvePlace(from, hotels, hotelByCity);
      const toH=resolvePlace(to, hotels, hotelByCity);
      const k=norm(kind);
      let label='', where_from='', where_to='', note='', missing=false;

      const fmtHotel=(h,fallback)=>h?`${h.name}\n${h.address||''}`:(fallback||'');
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
        const fl=flights.find(f=>/HAN|DAD|SGN/i.test(f.to)&&norm(f.arr).includes(date.replace('-','').slice(0,5)))||flights.find(f=>norm(f.arr).includes(dp.raw.replace(/\s/g,'').slice(0,5)));
        // better: match by place
        let fl2=null;
        if(/HAN/i.test(from)) fl2=flights.find(f=>/HAN/i.test(f.to));
        else if(/SGN|SAI/i.test(from)) fl2=flights.find(f=>/SGN/i.test(f.to));
        else if(/DAD/i.test(from)) fl2=flights.find(f=>/DAD/i.test(f.to));
        const fll=fl2||fl;
        const apt=/HAN/i.test(from)?'HANOI AIRPORT':/DAD/i.test(from)?'DAD AIRPORT':/SGN/i.test(from)?'SGN AIRPORT':from.toUpperCase();
        where_from=`TIME: ${time||''}\n${apt}`+(fll?`\nFlight: ${fll.carrier}\nArrival: ${time||parseTimeFrom(fll.arr)}`:'');
        where_to=fmtHotel(toH, to);
        const pillar=/HAN/i.test(from)?'pillar no.10':/SGN/i.test(from)?'pillar no. 23B':'the arrivals area';
        const term=/HAN/i.test(from)?'international terminal (outside the arrivals hall at '+pillar+')':/SGN/i.test(from)?'domestic terminal of Sai Gon airport (outside the arrivals hall '+pillar+')':'the arrivals hall';
        note=`The driver will pick you up at the ${term} with a sign:\n\n GO2/ MR. ${surname.toUpperCase()}`;
        if(!toH) missing=true;
      } else if(/^TRF/i.test(kind)&&/airport/i.test(to)){
        label=/SGN/i.test(to)&&/19:30|DEP/i.test(pickup+kind)?'TRANSPORT TO\nAIRPORT':'TRANSFER TO AIRPORT';
        if(/DAD/i.test(to)) label='TRANSFER TO AIRPORT';
        where_from=`TIME: ${time||''}\n`+fmtHotel(fromH, from);
        let fl2=null;
        if(/DAD/i.test(to)) fl2=flights.find(f=>/DAD/i.test(f.from));
        else if(/SGN/i.test(to)) fl2=flights.find(f=>/SGN/i.test(f.from)&&/VN\s*031|VN031/i.test(f.carrier+f.from)||true);
        // pick departure flight from that airport matching date
        fl2=flights.find(f=>norm(f.from).includes(norm(to).replace(' AIRPORT','')) && norm(f.dep).includes(date.replace('-','').slice(0,5)))
          ||flights.find(f=>norm(f.from).includes(norm(to).replace('AIRPORT','').trim()));
        const apt=/DAD/i.test(to)?'DAD AIRPORT':/SGN/i.test(to)?'SGN AIRPORT':to.toUpperCase();
        where_to=`${apt}`+(fl2?`\nFlight: ${fl2.carrier}\nDeparture: ${parseTimeFrom(fl2.dep)}`:'');
        if(/DAD/i.test(to)) note='The driver will pick you up at the hotel reception and drop you off at the Domestic terminal of Da Nang airport.';
        else note='The driver will pick you up at the hotel reception and drop you off at the International terminal of Saigon airport.';
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
        // tour / activity
        const act=activities.find(a=>norm(a.when).includes(date.replace('-','')))||activities.find(a=>norm(kind).split(' ').slice(0,2).every(t=>norm(a.name).includes(t)));
        // match by date
        const act2=activities.find(a=>norm(a.when).replace(/\s/g,'').includes(date.replace('-','')));
        const A=act2||act;
        const cz=/CZ|CSG|Czech/i.test(kind+(A?.notes||'')+(A?.name||''));
        label=kind.replace(/\s*\+\s*/g,'\n+');
        if(cz&&!/\+CZ/i.test(label)) label=label+'\n+CZ GUIDE';
        // prettier labels from samples
        if(/HA NOI CITY|FOOD TOUR/i.test(kind)&&/HAN|HA NOI/i.test(from+to)) label='HAN CITY TOUR & FOOD TOUR\n+CZ GUIDE';
        if(/HA LONG/i.test(kind)) label='HALONG BAY';
        if(/NINH BINH TOUR/i.test(kind)) label='NINH BINH TOUR';
        if(/HUE TOUR/i.test(kind)) label='HUE TOUR';
        if(/HOI AN TOUR/i.test(kind)) label='HOI AN TOUR\n+CZ GUIDE';
        if(/FOOD TOUR BY MOTORBIKE/i.test(kind)) label='FOOD TOUR BY MOTORBIKE';
        if(/CU CHI/i.test(kind)) label='CU CHI TUNNEL TOUR\n+CZ GUIDE';

        where_from=`TIME: ${time||''}\n`+fmtHotel(fromH, from);
        where_to=fmtHotel(toH, to);
        // build note from activity
        const an=A?.name||kind;
        const notes=A?.notes||'';
        let meals='No meals';
        if(/street food/i.test(notes)) meals='street food';
        else if(/1 breakfast.*2 lunches.*1 dinner/i.test(notes)) meals='1 breakfast, 2 lunches, 1 dinner';
        else if(/1\s*LUNCH/i.test(notes)) meals='1 lunch';
        else if(/Meal:\s*Street food/i.test(notes)) meals='Street food';
        else if(/no meals/i.test(notes)) meals='No meals';
        const guideLang=/Czech|CSG/i.test(notes+kind+an)?'Czech speaking':/English|ESG/i.test(notes+an)?'English speaking':'';
        let desc=an;
        // humanize
        if(/HA NOI.*FOOD/i.test(an+kind)) desc='1 Day Hanoi walking tour & Food tour after which you return to the hotel.';
        else if(/HA LONG/i.test(an+kind)) desc='2 Days 1 night HaLong Bay sightseeing cruise, after which you transfer to Ninh Binh hotel.';
        else if(/NINH BINH/i.test(an+kind)) desc='1 Day Ninh Binh sightseeing tour, after which you return to the Ninh Binh hotel.';
        else if(/HUE/i.test(an+kind)) desc='½ Day Hue walking tour, after which you return to the Hue hotel.';
        else if(/HOI AN/i.test(an+kind)) desc='½ Day Hoi An walking tour, after which you return to the Hoi An hotel.';
        else if(/MOTORBIKE|FOOD TOUR/i.test(an+kind)) desc='Sai Gon food tour by motorbike, after which you return to the Sai Gon hotel.';
        else if(/CU CHI/i.test(an+kind)) desc='½ Day Cu Chi sightseeing on tour, after which you return to the Sai Gon hotel.';

        if(/HA LONG/i.test(kind)){
          note=`${desc}\n-------------------------------------\nThe representative of the cruise will pick you up at the hotel reception.\n\nMeals on tour: ${meals}`;
        } else {
          const pay=/food tour by motorbike/i.test(kind)?'pay for street food':'pay for all entry fees'+(/(FOOD|street food)/i.test(notes+kind)?' & foods':'');
          note=`${desc}\n-------------------------------------\nThe ${guideLang} guide will pick you up at the hotel reception and ${pay}.\n\nMeals on tour: ${meals}`;
        }
        if(!fromH) missing=true;
      }
      trf_rows.push({label, date, where_from, where_to, note, missing, raw_kind:kind, time});
    }
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
      address:String(best.address||'').trim(),
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
  const key=cfg.GEMINI_API_KEY || (localStorage.prov==='gemini'?localStorage.key:'') || '';
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
      name:String(obj.name||query).trim(), address:String(obj.address||'').trim(),
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
    const hotel={id:'tavily-'+Date.now(), aliases:[query], name, address, phone: (phone||'').trim(), website, source:'tavily'};
    return {ok:true, hotel};
  }catch(e){
    console.warn('Tavily lookup failed', e);
    return {ok:false, reason:'network', error:String(e)};
  }
}

async function resolveHotelContact(query, city){
  // 1) local cache
  const cached=cacheLookup(query) || (city?cacheLookup(city+' '+query):null);
  if(cached && (cached.address||cached.phone)) return {...cached, from_cache:true};

  // 2) Serper (1 call)
  const s=await serperHotelLookup(query, city);
  if(s.ok){
    hotelCacheSet(query, s.hotel);
    if(s.hotel.name) hotelCacheSet(s.hotel.name, s.hotel);
    return s.hotel;
  }
  // 3) Tavily if Serper failed
  const t=await tavilyHotelLookup(query, city);
  if(t.ok){
    hotelCacheSet(query, t.hotel);
    if(t.hotel.name) hotelCacheSet(t.hotel.name, t.hotel);
    return t.hotel;
  }
  // 4) Gemini last resort — only when search APIs failed (quota/network/auth),
  // not when they returned unrelated/empty hits for a garbage name.
  const searchMiss = (s.reason==='low_score'||s.reason==='empty') && (t.reason==='low_score'||t.reason==='empty'||t.reason==='no_key');
  if(searchMiss) return null;
  const g=await geminiHotelLookup([query, city, 'Vietnam'].filter(Boolean).join(' '));
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

global.Voucher={parseSummary,loadHotels,saveHotelExtra,matchHotel,hotelBlock,todayVN,norm,renderVoucherDoc,mergeHotels,geminiHotelLookup,serperHotelLookup,tavilyHotelLookup,resolveHotelContact,enrichMissingHotels,hotelCacheGet};
})(window);
