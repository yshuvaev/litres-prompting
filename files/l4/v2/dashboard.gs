/** @OnlyCurrentDoc */
const LV_OUTPUTS = ['OUT_Checked','OUT_Issues','OUT_ModelAudit','OUT_Dashboard','OUT_Regions','OUT_Books','OUT_Reasons','OUT_Partner','OUT_Rejected'];
function lvRead(name, headers) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Missing sheet: '+name);
  const a=s.getRange(1,1,Math.max(1,s.getLastRow()),headers.length).getValues();
  if(a[0].some((v,i)=>v!==headers[i])) throw new Error('Wrong headers: '+name);
  if(a.length>100001) throw new Error('Training limit: 100000 rows per source');
  return a;
}
function lvWrite(name,a) {
  if(!LV_OUTPUTS.includes(name)) throw new Error('Output is not reserved');
  const ss=SpreadsheetApp.getActiveSpreadsheet(),s=ss.getSheetByName(name)||ss.insertSheet(name);
  if(s.getMaxRows()<a.length)s.insertRowsAfter(s.getMaxRows(),a.length-s.getMaxRows());
  if(s.getMaxColumns()<a[0].length)s.insertColumnsAfter(s.getMaxColumns(),a[0].length-s.getMaxColumns());
  s.getDataRange().clearContent().setBackground(null);
  s.getCharts().forEach(c=>s.removeChart(c));
  // Text copied from input or formula logs must not become an executable formula.
  const safe=a.map(r=>r.map(v=>typeof v==='string'&&v.startsWith('=')?"'"+v:v));
  s.getRange(1,1,a.length,a[0].length).setValues(safe);
  s.getRange(1,1,1,a[0].length).setFontWeight('bold').setBackground('#dcefe4');
  s.setFrozenRows(1);s.autoResizeColumns(1,a[0].length);return s;
}
function lvNum(v){return typeof v==='number'&&Number.isFinite(v);}
function lvRound(v){return Math.round(v*100)/100;}

// Two independent entry points; paste this complete module only once.
const AN_REASONS=['Повреждение','Ошибка комплектации','Отказ покупателя','Излишек запаса','Полиграфический брак'];
function analyseSales(sales,returns,outlets){
  const byID=new Map(),byCode=new Map(),groups=new Map(),books=new Map(),weeks=new Map(),regions=new Map();
  const id=(v)=>{if(typeof v!=='string'||! /^[A-Z0-9_-]+$/.test(v))throw Error('Expected uppercase ID, got: '+v);return v;};
  const week=v=>{if(typeof v!=='string'||! /^2026-W(2[7-9]|3[0-9])$/.test(v))throw Error('Expected week 2026-W27..W39');return v;};
  const num=(v,integer=false)=>{if(!lvNum(v)||v<0||(integer&&!Number.isInteger(v)))throw Error('Expected non-negative '+(integer?'integer':'number'));return v;};
  const each=(rows,name,fn)=>{for(let i=1;i<rows.length;i++){if(rows[i].every(v=>v===''))continue;try{fn(rows[i]);}catch(e){throw Error(name+', row '+(i+1)+': '+e.message);}}};
  each(outlets,'Outlets',r=>{const a=id(r[0]),b=id(r[1]);if(typeof r[2]!=='string'||!r[2].trim())throw Error('Empty region');if(byID.has(a)||byCode.has(b))throw Error('Duplicate outlet key');const o={id:a,region:r[2]};byID.set(a,o);byCode.set(b,o);});
  each(sales,'Sales',r=>{
    const w=week(r[0]),o=byID.get(id(r[1])),b=id(r[2]),q=num(r[4],true),g=num(r[5]);if(!o)throw Error('Unknown OutletID');if(typeof r[3]!=='string'||!r[3].trim())throw Error('Empty book name');
    if(!books.has(b))books.set(b,{name:r[3],sold:0,returned:0,amount:0,reasons:AN_REASONS.map(()=>[0,0])});const book=books.get(b);if(book.name!==r[3])throw Error('BookID has conflicting names');book.sold+=q;
    const key=[w,o.id,b].join('|');if(!groups.has(key))groups.set(key,{week:w,region:o.region,book:b,sold:0,gross:0,returned:0,amount:0});const v=groups.get(key);v.sold+=q;v.gross+=g;
  });
  each(returns,'Returns',r=>{
    const w=week(r[0]),rw=week(r[1]),o=byCode.get(id(r[2])),b=id(r[3]),q=num(r[4],true),a=num(r[5]),reason=AN_REASONS.indexOf(r[6]);if(rw<w)throw Error('Return precedes shipment');if(!o)throw Error('Unknown ReturnPointCode');if(reason<0)throw Error('Unknown return reason');
    const v=groups.get([w,o.id,b].join('|'));if(!v)throw Error('No corresponding shipment group');v.returned+=q;v.amount+=a;const book=books.get(b);book.returned+=q;book.amount+=a;book.reasons[reason][0]+=q;book.reasons[reason][1]+=a;
  });
  for(const v of groups.values()){
    if(v.returned>v.sold||v.amount>v.gross+0.005)throw Error('Returns exceed shipments: '+[v.week,v.region,v.book].join('/'));
    if(!weeks.has(v.week))weeks.set(v.week,[0,0]);if(!regions.has(v.week+'|'+v.region))regions.set(v.week+'|'+v.region,[v.week,v.region,0,0]);
    const w=weeks.get(v.week),r=regions.get(v.week+'|'+v.region);w[0]+=v.gross;w[1]+=v.amount;r[2]+=v.gross;r[3]+=v.amount;
  }
  const weekly=[...weeks].sort(([a],[b])=>a<b?-1:a>b?1:0).map(([w,[g,r]])=>[w,lvRound(g),lvRound(r),lvRound(g-r)]);
  const regional=[...regions.values()].sort((a,b)=>a[0]<b[0]?-1:a[0]>b[0]?1:a[1]<b[1]?-1:a[1]>b[1]?1:0).map(([w,reg,g,r])=>[w,reg,lvRound(g),lvRound(r),lvRound(g-r)]);
  const ranked=[...books].map(([b,v])=>[b,v.name,v.sold,v.returned,v.sold?v.returned/v.sold:0,lvRound(v.amount)]).sort((a,b)=>b[4]-a[4]||b[3]-a[3]||(a[0]<b[0]?-1:a[0]>b[0]?1:0));
  const reasons=ranked.flatMap(row=>AN_REASONS.map((reason,i)=>[row[0],row[1],reason,books.get(row[0]).reasons[i][0],lvRound(books.get(row[0]).reasons[i][1])]).sort((a,b)=>b[3]-a[3]||AN_REASONS.indexOf(a[2])-AN_REASONS.indexOf(b[2])));
  return {weekly,regions:regional,books:ranked,reasons};
}
function readAnalytics(){return analyseSales(lvRead('Sales',['ShipmentWeek','OutletID','BookID','Book','Units','GrossAmount']),lvRead('Returns',['ShipmentWeek','ReturnWeek','ReturnPointCode','BookID','Units','ReturnAmount','Reason']),lvRead('Outlets',['OutletID','ReturnPointCode','Region']));}
function buildDashboard(){
  const a=readAnalytics(); // Validate all inputs before changing any output.
  const s=lvWrite('OUT_Dashboard',[['ShipmentWeek','GrossAmount','ReturnAmount','NetAmount'],...a.weekly]);
  lvWrite('OUT_Regions',[['ShipmentWeek','Region','GrossAmount','ReturnAmount','NetAmount'],...a.regions]);
  if(!a.weekly.length)return;
  s.getRange(2,2,a.weekly.length,3).setNumberFormat('#,##0.00');
  const chart=s.newChart().setChartType(Charts.ChartType.LINE).addRange(s.getRange(1,1,a.weekly.length+1,2)).addRange(s.getRange(1,4,a.weekly.length+1,1)).setNumHeaders(1)
   .setOption('title','Отгрузки и продажи за вычетом известных возвратов').setOption('vAxis',{title:'Стоимость, руб.',viewWindow:{min:0}}).setOption('hAxis',{title:'Неделя исходной отгрузки'})
   .setOption('series',{0:{color:'#83aaf3'},1:{color:'#ff5533',lineWidth:3}}).setOption('width',1000).setOption('height',480).setPosition(2,7,0,0).build();s.insertChart(chart);
}
function analyseReturns(){
  const a=readAnalytics();
  const s=lvWrite('OUT_Books',[['BookID','Book','ShippedUnits','ReturnedUnits','ReturnRate','ReturnAmount'],...a.books]);
  lvWrite('OUT_Reasons',[['BookID','Book','Reason','ReturnedUnits','ReturnAmount'],...a.reasons]);
  if(!a.books.length)return;
  s.getRange(2,5,a.books.length,1).setNumberFormat('0.0%');
  const n=Math.min(10,a.books.length);
  s.insertChart(s.newChart().setChartType(Charts.ChartType.BAR).addRange(s.getRange(1,2,n+1,1)).addRange(s.getRange(1,5,n+1,1)).setNumHeaders(1)
   .setOption('title','10 книг с наибольшей долей возврата').setOption('hAxis',{title:'Возвращено / отгружено, экземпляры',format:'percent',viewWindow:{min:0}}).setOption('colors',['#ff5533']).setOption('width',1000).setOption('height',520).setPosition(2,9,0,0).build());
}
