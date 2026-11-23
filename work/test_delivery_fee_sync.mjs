import assert from 'node:assert/strict';
import {businessDate,selectDeliveryOrders,decideDeliveryFee,resolveCyanDispatchHistory} from '../agent/delivery_fee_policy.mjs';
import {readProviderPage,extractDeliveryFees,verifyServiceSettlement} from './delivery_fee_extract.mjs';
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
  for(const cancelFee of [0,200]) {
    respond=u=>u.pathname==='/order/detail/'
      ? {errno:1,data:{source_id:id,delivery:{id:'123',status:50,send_fee:907,tip:100,premium_fee:200,damage_fee:0,pre_tip:0}}}
      : {errno:1,data:[{...row,delivery:{...row.delivery,send_fee:907,tip:100},
        deliveries:[row.deliveries[0],{...row.deliveries[1],cancel_fee:cancelFee}]}]};
    result=await readProviderPage('maiyatian',['2026-11-03'],[id]);
    check(()=>{assert.equal(result[0].state,'ambiguous');assert.equal(result[0].reviewable_tip,cancelFee===0?1:undefined);});
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
  const comparisonRow={...qyRow,waybillId:'paid-waybill'};
  const completedBill={waybillId:'paid-waybill',carrierDeliveryId:'carrier-id',status:50,actualDeliveryFee:451,
    cancelFee:0,totalExtFee:0,tipAmount:0,serviceName:'美团跑腿',waybillTrackInfos:[{ctime:1793686200,status:50}]};
  for (const variant of ['valid','extra','another-bill','wrong-bill','failed']) {
    const waybillList=[{...completedBill,...(variant==='extra'?{cancelFee:200}:variant==='wrong-bill'?{waybillId:'other'}:{})}];
    if (variant==='another-bill') waybillList.push({...completedBill,waybillId:'other'});
    respond=u=>u.pathname==='/api/pc/delivery/progressTrack'
      ? variant==='failed'?{code:1}:{code:0,data:{deliveryList:[{deliveryId:'321',status:50,waybillList}]}}
      : {code:0,data:{total:1,pageNum:1,dataList:[comparisonRow]}};
    result=await readProviderPage('cyanclouds',['2026-11-03'],[id],resolveCyanDispatchHistory,[id]);
    check(()=>{
      assert.equal(result[0].state,'complete');
      assert.deepEqual(result[0].dispatch,variant==='valid'?{single_completed:true,carrier:'美团跑腿',finished_at:1793686200,carrier_waybill:'carrier-id'}:undefined);
    });
  }
  respond=()=>({code:0,data:{total:1,pageNum:1,dataList:[{...qyRow,shipCount:2}]}});
  result=await readProviderPage('cyanclouds',['2026-11-03'],[id]);assert.equal(result[0].state,'ambiguous');cases++;
  // Regression: two send rounds are not two charges. Cancelled waybills still
  // contain positive quotes, but their actual cancellation/additional fees are 0.
  const multi={...qyRow,shipCount:2,waybillId:'paid-waybill',orderDeliveryInfo:{...qyRow.orderDeliveryInfo,actualDeliveryFee:733,deliveryFee:977,discountFee:244}};
  const fees={cancelFee:0,totalExtFee:0,insuredFee:0,tipAmount:0,allowanceFee:0};
  const history={deliveryList:[
    {deliveryId:'321',status:50,waybillList:[{...fees,waybillId:'paid-waybill',status:50,actualDeliveryFee:733},{...fees,waybillId:'cancelled-quote',status:99,actualDeliveryFee:999}]},
    {deliveryId:'320',status:99,waybillList:[{...fees,waybillId:'old-quote',status:99,actualDeliveryFee:548}]}
  ]};
  respond=u=>u.pathname==='/api/pc/delivery/progressTrack'?{code:0,data:history}:{code:0,data:{total:1,pageNum:1,dataList:[multi]}};
  result=await readProviderPage('cyanclouds',['2026-11-03'],[id]);assert.equal(result[0].state,'complete');assert.equal(result[0].amount,7.33);cases++;
  check(()=>assert.deepEqual(result[0].history,{rounds:2,completed:1,cancelled:2,cancel_fee:0}));
  const good=result[0];
  check(()=>assert.equal(decideDeliveryFee(base,[good],providers).amount,7.33));
  check(()=>assert.equal(decideDeliveryFee({...base,current_fee:7.33},[good],providers).changed,false));
  check(()=>assert.equal(decideDeliveryFee({...base,current_fee:7.33,protected:true},[good],providers).state,'protected'));
  check(()=>assert.equal(decideDeliveryFee({...base,current_fee:8,protected:true},[good],providers).state,'conflict'));
  check(()=>assert.equal(decideDeliveryFee(base,[good,paid()],providers).state,'conflict'));
  // The production browser receives the helper explicitly: no Node/module
  // closure, credentials, or filesystem is available inside that context.
  const detached=Function(`return (${readProviderPage.toString()});`)();
  const detachedResolver=Function(`return (${resolveCyanDispatchHistory.toString()});`)();
  result=await detached('cyanclouds',['2026-11-03'],[id],detachedResolver);
  check(()=>assert.equal(result[0].amount,7.33));
  const mutations=[
    h=>{h.deliveryList.pop();},
    h=>{h.deliveryList[1].deliveryId='321';},
    h=>{h.deliveryList[1].waybillList[0].waybillId='paid-waybill';},
    h=>{delete h.deliveryList[1].deliveryId;},
    h=>{h.deliveryList[1].deliveryId=Number.MAX_SAFE_INTEGER+1;},
    h=>{h.deliveryList[1].waybillList=[];},
    h=>{h.deliveryList[1].status=20;},
    h=>{h.deliveryList[1].waybillList[0].status=20;},
    h=>{h.deliveryList[0].waybillList[0].actualDeliveryFee=999;},
    h=>{h.deliveryList[0].waybillList[0].waybillId='wrong-current-waybill';},
    h=>{h.deliveryList[0].waybillList[0].actualDeliveryFee=null;},
    h=>{h.deliveryList[0].waybillList[0].actualDeliveryFee=-1;},
    h=>{h.deliveryList[0].waybillList[0].actualDeliveryFee=733.5;},
    h=>{h.deliveryList[0].waybillList[0].actualDeliveryFee='unknown';},
    h=>{h.deliveryList[1].waybillList[0].status=50;},
    h=>{h.deliveryList[1].status=50;h.deliveryList[1].waybillList[0].status=50;},
    h=>{h.deliveryList[0].waybillList[1].status=50;},
    ...['cancelFee','totalExtFee','insuredFee','tipAmount','allowanceFee'].map(k=>h=>{h.deliveryList[1].waybillList[0][k]=200;}),
    ...['cancelFee','totalExtFee','insuredFee','tipAmount'].map(k=>h=>{delete h.deliveryList[1].waybillList[0][k];}),
    ...[null,'',false,-1,'unknown'].map(v=>h=>{h.deliveryList[1].waybillList[0].cancelFee=v;})
  ];
  for(const mutate of mutations){
    const h=structuredClone(history);mutate(h);
    const resolved=resolveCyanDispatchHistory(multi,h);
    check(()=>{assert.equal(resolved.state,'ambiguous');assert.ok(resolved.reason);});
  }
  for(const override of [{deliveryId:'other-round'},{waybillId:'other-waybill'},{shipCount:3},{deliveryStatus:20},{orderDeliveryInfo:{...multi.orderDeliveryInfo,actualDeliveryFee:734}},{orderDeliveryInfo:{...multi.orderDeliveryInfo,tipFee:null}}]){
    check(()=>assert.equal(resolveCyanDispatchHistory({...multi,...override},history).state,'ambiguous'));
  }
  const stringHistory=structuredClone(history);
  for(const round of stringHistory.deliveryList){round.status=String(round.status);for(const waybill of round.waybillList){waybill.status=String(waybill.status);for(const key of ['cancelFee','totalExtFee','insuredFee','tipAmount'])waybill[key]='0';}}
  check(()=>assert.equal(resolveCyanDispatchHistory(multi,stringHistory).amount,7.33));
  const allCancelled=structuredClone(history);
  for(const round of allCancelled.deliveryList){round.status=99;for(const waybill of round.waybillList)waybill.status=99;}
  check(()=>assert.equal(resolveCyanDispatchHistory({...multi,deliveryStatus:60},allCancelled).state,'missing'));
  check(()=>assert.equal(resolveCyanDispatchHistory(multi,allCancelled).state,'ambiguous'));
  const zero=structuredClone(history);zero.deliveryList[0].waybillList[0].actualDeliveryFee=0;
  check(()=>assert.equal(resolveCyanDispatchHistory({...multi,orderDeliveryInfo:{...multi.orderDeliveryInfo,actualDeliveryFee:0}},zero).amount,0));
  const requests=[];
  respond=(u,b)=>{requests.push(u.pathname);if(u.pathname==='/api/pc/delivery/progressTrack'){assert.equal(new URLSearchParams(b).get('orderId'),'201');return{code:0,data:history};}return{code:0,data:{total:1,pageNum:1,dataList:[multi]}};};
  result=await readProviderPage('cyanclouds',['2026-11-03'],[id]);
  check(()=>assert.deepEqual(requests,['/api/pc/order/manage/list','/api/pc/delivery/progressTrack']));
  const integration=await extractDeliveryFees({date:base.date,orders:[base]},async(provider)=>provider==='cyanclouds'?result:[]);
  check(()=>{assert.equal(integration.decisions[0].state,'confirmed');assert.equal(integration.decisions[0].amount,7.33);assert.equal(integration.pending.length,0);});
  respond=u=>u.pathname==='/api/pc/delivery/progressTrack'?{code:1}:{code:0,data:{total:1,pageNum:1,dataList:[multi]}};
  await assert.rejects(()=>readProviderPage('cyanclouds',['2026-11-03'],[id]),/history unavailable/);cases++;
  const deferred=await extractDeliveryFees({date:base.date,orders:[base]},async(provider)=>provider==='cyanclouds'?await readProviderPage(provider,[base.date],[id]):[]);
  check(()=>{assert.equal(deferred.decisions[0].state,'deferred');assert.equal(deferred.pending.length,1);});
  const charged=structuredClone(history);charged.deliveryList[1].waybillList[0].cancelFee=200;
  respond=u=>u.pathname==='/api/pc/delivery/progressTrack'?{code:0,data:charged}:{code:0,data:{total:1,pageNum:1,dataList:[multi]}};
  result=await readProviderPage('cyanclouds',['2026-11-03'],[id]);
  check(()=>{assert.equal(result[0].state,'ambiguous');assert.match(decideDeliveryFee(base,result,providers).reason,/取消费/);});
  respond=()=>({code:0,data:{total:100,pageNum:1,dataList:[qyRow]}});
  await assert.rejects(()=>readProviderPage('cyanclouds',['2026-11-03'],[id]),/Incomplete/);cases++;
  respond=()=>({code:0,data:{total:40,pageNum:1,dataList:Array.from({length:20},(_,i)=>({...qyRow,orderId:String(i)}))}});
  await assert.rejects(()=>readProviderPage('cyanclouds',['2026-11-03'],[id]),/Unexpected page/);cases++;
}finally{
  if(originalPerformance)Object.defineProperty(globalThis,'performance',originalPerformance);
  globalThis.XMLHttpRequest=originalXhr;globalThis.location=originalLocation;
}
{
  const day='2026-11-20', next='2026-11-21';
  const summary={date:next,description:'技术服务结算(1笔)',amount:'0.2'};
  const row=['技术服务',id,'美团 #1','2026-11-21 00:14:00','已完成','0.20'];
  const tech=verifyServiceSettlement(summary,[row]);
  check(()=>assert.deepEqual(tech,[{order_no:id,settlement_date:next,cents:20}]));
  check(()=>assert.throws(()=>verifyServiceSettlement({...summary,amount:'0.60'},[row]),/totals/));
  check(()=>assert.throws(()=>verifyServiceSettlement({...summary,description:'技术服务结算(2笔)'},[row]),/Incomplete/));
  check(()=>assert.throws(()=>verifyServiceSettlement(summary,[[...row.slice(0,4),'已取消','0.2']]),/status/));
  check(()=>assert.throws(()=>verifyServiceSettlement(summary,[[...row.slice(0,5),'-0.2']]),/amount/));
  const eligible={...base,date:day};
  const fees=async provider=>provider==='maiyatian'?[paid('maiyatian',9.07)]:[];
  const run=(order,reading,providerRead=fees,date=next)=>
    extractDeliveryFees({date,orders:[order],pending:[order]},providerRead,reading);
  const completed=await run(eligible,async (orders, date)=>{
    assert.equal(orders[0].order_no,id);assert.equal(date,next);return tech;
  });
  check(()=>{assert.equal(completed.decisions[0].amount,9.27);assert.equal(completed.decisions[0].technical_service_fee,0.20);assert.equal(completed.pending.length,0);});
  const again=await run({...eligible,current_fee:9.27,has_technical_service_fee:true},async()=>tech);
  check(()=>{assert.equal(again.decisions[0].amount,9.27);assert.equal(again.decisions[0].changed,false);});
  const late=await run(eligible,async()=>[]);
  check(()=>{assert.equal(late.decisions[0].amount,9.07);assert.equal(late.pending.length,1);});
  const failure=await run({...eligible,current_fee:9.27,has_technical_service_fee:true},async()=>{throw Error('timeout');});
  check(()=>{assert.equal(failure.tech_service_settlement,'deferred');assert.equal(failure.decisions[0].state,'retained');});
  const stale=await run({...eligible,current_fee:9.27,has_technical_service_fee:true},async()=>[],fees,'2026-11-29');
  check(()=>{assert.equal(stale.tech_service_settlement,'not_needed');assert.equal(stale.decisions[0].state,'retained');assert.equal(stale.pending.length,0);});
  const duplicate=await run(eligible,async()=>[...tech,...tech]);
  check(()=>assert.equal(duplicate.decisions[0].state,'conflict'));
  const manual=await run({...eligible,protected:true,current_fee:10.27},async()=>tech);
  check(()=>{assert.equal(manual.tech_service_settlement,'not_needed');assert.equal(manual.decisions[0].changed,undefined);});
  const other=async provider=>provider==='cyanclouds'?[paid('cyanclouds',3.21)]:[];
  const switched=await run({...eligible,current_fee:9.27,has_technical_service_fee:true},async()=>tech,other);
  check(()=>assert.equal(switched.decisions[0].state,'conflict'));
  const manualOrder={...eligible,current_fee:10.27,protected:true,confirmed_manual:true};
  const manualSource={...paid('maiyatian',9.07),state:'ambiguous',reviewable_tip:1};
  const reviewed=async provider=>provider==='maiyatian'?[manualSource]:[];
  const matched=await run(manualOrder,async()=>tech,reviewed);
  check(()=>{assert.equal(matched.decisions[0].state,'protected');assert.equal(matched.decisions[0].changed,false);assert.equal(matched.pending.length,0);});
  const newFee=await run(manualOrder,async()=>[{...tech[0],cents:40}],reviewed);
  check(()=>{assert.equal(newFee.decisions[0].state,'conflict');assert.match(newFee.decisions[0].reason,/10\.47.*10\.27/);});
  const missing=await run(manualOrder,async()=>[],reviewed);
  check(()=>{assert.equal(missing.decisions[0].state,'protected');assert.equal(missing.decisions[0].changed,false);assert.equal(missing.pending.length,1);});
  const baseConflict=await run({...manualOrder,current_fee:9.07},async()=>[],reviewed);
  check(()=>{assert.equal(baseConflict.decisions[0].state,'conflict');assert.match(baseConflict.decisions[0].reason,/10\.07.*9\.07/);});
  const duplicateBill=await run(manualOrder,async()=>[...tech,...tech],reviewed);
  check(()=>assert.equal(duplicateBill.decisions[0].state,'conflict'));
  const imported=await run({...manualOrder,confirmed_manual:false},async()=>tech,reviewed);
  check(()=>assert.equal(imported.decisions[0].state,'conflict'));
  const twoProviders=await run(manualOrder,async()=>tech,async provider=>provider==='maiyatian'?[manualSource]:[paid('cyanclouds',3.21)]);
  check(()=>assert.equal(twoProviders.decisions[0].state,'conflict'));

  const mirrorMai={...paid('maiyatian',4.10),dispatch:{carrier:'美团跑腿',finished_at:1795310813,single_completed:true}};
  const mirrorCyan={...paid('cyanclouds',4.10),dispatch:{carrier:'美团跑腿',finished_at:1795310825,single_completed:true,carrier_waybill:'carrier-id'}};
  const compare=(mai=mirrorMai,cyan=mirrorCyan)=>async (provider,_dates,_orders,compareIds)=>{
    if(provider==='maiyatian')return [mai];
    assert.deepEqual(compareIds,mai.state==='complete'?[id]:[]);return [cyan];
  };
  const reconcile=await run({...eligible,current_fee:4.10},async()=>tech,compare());
  check(()=>{const d=reconcile.decisions[0];assert.equal(d.amount,4.30);assert.equal(d.changed,true);assert.equal(d.cross_platform_mirror,true);assert.equal(d.provider,'maiyatian');assert.equal(reconcile.pending.length,0);});
  const noSecondCharge=await run({...eligible,protected:true,current_fee:4.30},async()=>tech,compare());
  check(()=>{assert.equal(noSecondCharge.decisions[0].state,'protected');assert.equal(noSecondCharge.decisions[0].changed,false);assert.equal(noSecondCharge.pending.length,0);});
  const neverLower=await run({...eligible,current_fee:8.20},async()=>tech,compare());
  check(()=>{assert.equal(neverLower.decisions[0].state,'conflict');assert.match(neverLower.decisions[0].reason,/4\.30.*8\.20/);});
  for(const missingProof of [[],[{...tech[0],order_no:'3602371202674777016'}],[...tech,...tech],[{...tech[0],cents:0}]]) {
    const review=await run(eligible,async()=>missingProof,compare());
    check(()=>{assert.equal(review.decisions[0].state,'conflict');assert.match(review.decisions[0].conflict_key,/^[a-f0-9]{64}$/);assert.equal(review.pending.length,1);});
  }
  for(const patch of [{carrier:'顺丰快送'},{finished_at:1795310844},{finished_at:null},{single_completed:false},{separate_charge:true}]) {
    const review=await run(eligible,async()=>{throw Error('Should not read wallet without matching deliveries');},compare(mirrorMai,{...mirrorCyan,dispatch:{...mirrorCyan.dispatch,...patch}}));
    check(()=>{assert.equal(review.decisions[0].state,'conflict');assert.equal(review.tech_service_settlement,'not_needed');});
  }
  const differentCarrierId=await run(eligible,async()=>tech,compare({...mirrorMai,dispatch:{...mirrorMai.dispatch,carrier_waybill:'different-carrier-id'}}));
  check(()=>assert.equal(differentCarrierId.decisions[0].state,'conflict'));
  const sameConflict=await run({...eligible,current_fee:8.20},async()=>tech,compare(),'2026-11-23');
  check(()=>assert.equal(sameConflict.decisions[0].conflict_key,neverLower.decisions[0].conflict_key));
  const newEvidence=await run({...eligible,current_fee:8.20},async()=>tech,compare(mirrorMai,{...mirrorCyan,amount:4.11}));
  check(()=>{assert.equal(newEvidence.decisions[0].state,'conflict');assert.notEqual(newEvidence.decisions[0].conflict_key,neverLower.decisions[0].conflict_key);});
}
process.stdout.write(JSON.stringify({ok:true,cases,external_writes:0}));
