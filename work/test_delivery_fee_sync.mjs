import assert from 'node:assert/strict';
import {businessDate,selectDeliveryOrders,decideDeliveryFee} from '../agent/delivery_fee_policy.mjs';
import {readProviderPage,extractDeliveryFees} from './delivery_fee_extract.mjs';
const id='3602371202674777015';
const base={order_no:id,date:'2026-11-03',platform:'美团',status:'已完成',current_fee:0};
const providers={maiyatian:{ok:true},cyanclouds:{ok:true}};
const paid=(provider='maiyatian',amount=6.78)=>({provider,order_no:id,state:'complete',amount,delivery_id:'123'});
let cases=0;
const check=(fn)=>{fn();cases++;};
check(()=>assert.equal(businessDate('2026-11-03 14:30'),'2026-11-03'));
check(()=>assert.equal(businessDate(46329.75),'2026-11-03'));
for(const current_fee of [0,null,6.78])check(()=>assert.equal(selectDeliveryOrders([{...base,current_fee,previously_checked:true}],'2026-11-03').length,1));
check(()=>assert.equal(selectDeliveryOrders([{...base,date:'2026-11-02'}],'2026-11-03',[{order_no:id}]).length,1));
check(()=>assert.equal(selectDeliveryOrders([base],'2026-11-02',[{order_no:id}]).length,0));
check(()=>assert.equal(selectDeliveryOrders([{...base,date:'2026-10-31'}],'2026-11-03').length,0));
check(()=>assert.equal(selectDeliveryOrders([{...base,date:'2026-11-02'}],'2026-11-03').length,1));
check(()=>assert.equal(selectDeliveryOrders([{...base,date:'2026-11-02',current_fee:5}],'2026-11-03').length,0));
check(()=>assert.equal(selectDeliveryOrders([{...base,date:'2026-11-02',protected:true}],'2026-11-03').length,0));
check(()=>assert.equal(selectDeliveryOrders([{...base,date:'2026-11-02',confirmed_self:true}],'2026-11-03').length,0));
check(()=>assert.equal(selectDeliveryOrders([{...base,date:'2026-10-31'}],'2026-11-01').length,1));
for(const override of [{intervention:true},{status:'已取消'},{platform:'京东'}])check(()=>assert.equal(selectDeliveryOrders([{...base,...override}],'2026-11-03').length,0));
check(()=>assert.throws(()=>selectDeliveryOrders([base,base],'2026-11-03')));
check(()=>assert.equal(decideDeliveryFee(base,[],providers).state,'default_self'));
check(()=>assert.equal(decideDeliveryFee(base,[{...paid(),amount:0}],providers).state,'default_self'));
check(()=>assert.equal(decideDeliveryFee(base,[{...paid(),state:'pending'}],providers).state,'default_self'));
check(()=>assert.equal(decideDeliveryFee(base,[{...paid(),state:'ambiguous'}],providers).state,'conflict'));
check(()=>assert.equal(decideDeliveryFee(base,[paid()],providers).amount,6.78));
check(()=>assert.equal(decideDeliveryFee(base,[paid('cyanclouds',4.51)],providers).amount,4.51));
check(()=>assert.equal(decideDeliveryFee(base,[{...paid(),state:'pending',amount:null},paid('cyanclouds',4.51)],providers).amount,4.51));
check(()=>assert.equal(decideDeliveryFee(base,[paid(),{...paid('cyanclouds'),state:'pending',amount:null}],providers).amount,6.78));
check(()=>assert.equal(decideDeliveryFee(base,[paid(),paid('cyanclouds')],providers).state,'conflict'));
check(()=>assert.equal(decideDeliveryFee(base,[paid(),paid()],providers).state,'conflict'));
check(()=>assert.equal(decideDeliveryFee({...base,protected:true,current_fee:4.51},[paid('cyanclouds',4.51)],providers).state,'protected'));
check(()=>assert.equal(decideDeliveryFee({...base,protected:true,current_fee:4.51},[paid()],providers).state,'conflict'));
check(()=>assert.equal(decideDeliveryFee({...base,current_fee:6.78},[paid()],providers).changed,false));
check(()=>assert.equal(decideDeliveryFee(base,[paid('cyanclouds',4.51)],{...providers,maiyatian:{ok:false}}).state,'deferred'));
// A zero at 10:00 is never a terminal status; 14:00 can still discover payment.
check(()=>{assert.equal(decideDeliveryFee(base,[],providers).state,'default_self');assert.equal(decideDeliveryFee(base,[paid()],providers).changed,true);});
check(()=>{const d=decideDeliveryFee({...base,current_fee:null},[],providers);assert.equal(d.amount,0);assert.equal(d.changed,true);assert.equal(d.retry,false);});
check(()=>{const d=decideDeliveryFee(base,[],providers);assert.equal(d.changed,false);assert.equal(d.retry,false);});
check(()=>{const d=decideDeliveryFee({...base,current_fee:6.78},[],providers);assert.equal(d.state,'retained');assert.equal(d.amount,undefined);});
check(()=>{const d=decideDeliveryFee({...base,current_fee:6.78},[{...paid(),amount:0}],providers);assert.equal(d.state,'retained');assert.equal(d.changed,undefined);});
check(()=>assert.equal(decideDeliveryFee(base,[{...paid(),state:'pending',amount:null}],providers).retry,true));
check(()=>{const d=decideDeliveryFee({...base,current_fee:6.78},[],{...providers,maiyatian:{ok:false}});assert.equal(d.state,'deferred');assert.equal(d.retry,true);assert.equal(d.amount,undefined);});

// Test the real extractor orchestration with an injected read-only reader.
const calls=[];
const first=await extractDeliveryFees({date:'2026-11-03',orders:[base],pending:[]},async(p,dates,orders)=>{calls.push({p,dates,ids:orders.map(o=>o.order_no)});return [];});
assert.equal(first.pending.length,0);assert.equal(first.decisions[0].state,'default_self');assert.equal(calls.length,2);cases++;
const second=await extractDeliveryFees({date:'2026-11-03',orders:[base],pending:first.pending},async(p)=>p==='cyanclouds'?[paid('cyanclouds',4.51)]:[]);
assert.equal(second.decisions[0].amount,4.51);assert.equal(second.pending.length,0);cases++;
const crossDay=await extractDeliveryFees({date:'2026-11-04',orders:[base],pending:first.pending},async(p,dates)=>{assert.deepEqual(dates,['2026-11-03']);return p==='maiyatian'?[paid()]:[];});
assert.equal(crossDay.decisions[0].amount,6.78);cases++;
const lostQueue=await extractDeliveryFees({date:'2026-11-04',orders:[base],pending:[]},async(p,dates)=>{assert.deepEqual(dates,['2026-11-03']);return p==='maiyatian'?[paid()]:[];});
assert.equal(lostQueue.decisions[0].amount,6.78);assert.equal(lostQueue.pending.length,0);cases++;
const crossDaySelf=await extractDeliveryFees({date:'2026-11-04',orders:[{...base,protected:true,confirmed_self:true}],pending:[{order_no:id,date:'2026-11-03'}]},async()=>{throw Error('Confirmed self delivery must not query');});
assert.equal(crossDaySelf.queried_count,0);assert.equal(crossDaySelf.pending.length,0);cases++;
const self=await extractDeliveryFees({date:'2026-11-03',orders:[{...base,confirmed_self:true}],pending:[]},async()=>{throw Error('Self delivery must not query');});
assert.equal(self.queried_count,0);assert.equal(self.pending.length,0);cases++;
const outage=await extractDeliveryFees({date:'2026-11-03',orders:[base],pending:[]},async()=>{throw Error('offline');});
assert.equal(outage.pending.length,1);assert.equal(outage.decisions[0].state,'deferred');cases++;
const booked=await extractDeliveryFees({date:'2026-11-03',orders:[base],pending:[]},async(p)=>p==='maiyatian'?[{...paid(),state:'pending',amount:null}]:[]);
assert.equal(booked.pending.length,1);assert.equal(booked.decisions[0].state,'default_self');cases++;
const bookedLater=await extractDeliveryFees({date:'2026-11-05',orders:[base],pending:booked.pending},async(p)=>p==='maiyatian'?[paid()]:[]);
assert.equal(bookedLater.decisions[0].amount,6.78);assert.equal(bookedLater.pending.length,0);cases++;
const oldMissing=await extractDeliveryFees({date:'2026-11-05',orders:[base],pending:[{order_no:id,date:base.date}]},async()=>[]);
assert.equal(oldMissing.decisions[0].state,'default_self');assert.equal(oldMissing.pending.length,0);cases++;
const defaultNotLocked=await extractDeliveryFees({date:'2026-11-03',orders:[{...base,note:'默认自配送：未查到骑手配送费（0元，后续继续自动查询）'}],pending:[]},async(p)=>p==='maiyatian'?[paid()]:[]);
assert.equal(defaultNotLocked.queried_count,1);assert.equal(defaultNotLocked.decisions[0].amount,6.78);cases++;

const originalPerformance=Object.getOwnPropertyDescriptor(globalThis,'performance');
const originalXhr=globalThis.XMLHttpRequest;
const originalLocation=globalThis.location;
globalThis.location={href:'https://saas.maiyatian.com/query/'};
Object.defineProperty(globalThis,'performance',{configurable:true,value:{getEntriesByType:()=>[{name:'https://saas.maiyatian.com/query/list/?f=json'}]}});
let respond;
globalThis.XMLHttpRequest=class {
  open(method,url){this.method=method;this.url=new URL(url,globalThis.location.href);}
  setRequestHeader(){}
  send(body){try{this.status=200;this.responseText=JSON.stringify(respond(this.url,body));queueMicrotask(()=>this.onload());}catch{queueMicrotask(()=>this.onerror());}}
};
try{
  const row={id:'101',source_id:id,order_date:'2026-11-03',shop:{name:'宠速星宠物用品店'},delivery:{id:'123',status:50,send_fee:678,tip:0},deliveries:[{id:'123',status:50},{id:'old',status:60,send_fee:999}]};
  respond=u=>u.pathname==='/order/detail/'?{errno:1,data:{source_id:id,delivery:{id:'123',status:50,send_fee:678}}}:{errno:1,data:[row]};
  let result=await readProviderPage('maiyatian',['2026-11-03'],[id]);assert.equal(result[0].amount,6.78);assert.equal(result[0].state,'complete');cases++;
  assert.equal(JSON.stringify(result).includes('token'),false);cases++;
  respond=u=>u.pathname==='/order/detail/'?{errno:1,data:{source_id:id,delivery:{id:'123',status:50,send_fee:678,discount_fee:100}}}:{errno:1,data:[row]};
  result=await readProviderPage('maiyatian',['2026-11-03'],[id]);assert.equal(result[0].state,'complete');assert.equal(result[0].amount,6.78);cases++;
  // Real incident: the displayed delivery charge is 3.87, discount metadata
  // is 0.38. Neither block it nor incorrectly turn the charge into 3.49.
  respond=u=>u.pathname==='/order/detail/'?{errno:1,data:{source_id:id,delivery:{id:'123',status:'50',send_fee:'387',discount_fee:'38',premium_fee:'0',damage_fee:'0',tip:'0'}}}:{errno:1,data:[{...row,delivery:{id:'123',status:50,send_fee:387,tip:0}}]};
  result=await readProviderPage('maiyatian',['2026-11-03'],[id]);assert.equal(result[0].state,'complete');assert.equal(result[0].amount,3.87);cases++;
  for(const extra of [{premium_fee:20},{damage_fee:20},{tip:20},{discount_fee:'unknown'}]){
    respond=u=>u.pathname==='/order/detail/'?{errno:1,data:{source_id:id,delivery:{id:'123',status:50,send_fee:678,...extra}}}:{errno:1,data:[row]};
    result=await readProviderPage('maiyatian',['2026-11-03'],[id]);assert.equal(result[0].state,'ambiguous');cases++;
  }
  respond=()=>({errno:1,data:[{...row,delivery:{id:'123',status:20,send_fee:678}}]});
  result=await readProviderPage('maiyatian',['2026-11-03'],[id]);assert.equal(result[0].state,'pending');assert.equal(result[0].amount,null);cases++;
  respond=()=>({errno:1,data:[{...row,order_date:'2026-11-02'}]});
  await assert.rejects(()=>readProviderPage('maiyatian',['2026-11-03'],[id]),/date mismatch/);cases++;
  respond=()=>({errno:1,data:[{...row,shop:{name:'另外一个店'}}]});
  await assert.rejects(()=>readProviderPage('maiyatian',['2026-11-03'],[id]),/Store identity/);cases++;
  const qyRow={orderId:'201',channelOrderId:id,channelShopName:'宠速星宠物用品店（环常南路店）',channelOrderTime:1793685600,deliveryStatus:50,deliveryId:'321',shipCount:1,orderDeliveryInfo:{actualDeliveryFee:451,deliveryFee:557,discountFee:106,totalExtFee:0,tipFee:0,insuredFee:0}};
  respond=(u,b)=>({code:0,data:{total:1,pageNum:Number(new URLSearchParams(b).get('pageNum')),dataList:[qyRow]}});
  result=await readProviderPage('cyanclouds',['2026-11-03'],[id]);assert.equal(result[0].amount,4.51);assert.equal(result[0].state,'complete');cases++;
  respond=()=>({code:0,data:{total:1,pageNum:1,dataList:[{...qyRow,shipCount:2}]}});
  result=await readProviderPage('cyanclouds',['2026-11-03'],[id]);assert.equal(result[0].state,'ambiguous');cases++;
  respond=()=>({code:0,data:{total:100,pageNum:1,dataList:[qyRow]}});
  await assert.rejects(()=>readProviderPage('cyanclouds',['2026-11-03'],[id]),/Incomplete/);cases++;
  respond=()=>({code:0,data:{total:40,pageNum:1,dataList:Array.from({length:20},(_,i)=>({...qyRow,orderId:String(i)}))}});
  await assert.rejects(()=>readProviderPage('cyanclouds',['2026-11-03'],[id]),/Unexpected page/);cases++;
}finally{
  if(originalPerformance)Object.defineProperty(globalThis,'performance',originalPerformance);
  globalThis.XMLHttpRequest=originalXhr;globalThis.location=originalLocation;
}
process.stdout.write(JSON.stringify({ok:true,cases,external_writes:0}));
