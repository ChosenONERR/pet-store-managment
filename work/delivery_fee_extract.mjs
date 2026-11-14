import fs from 'node:fs';
import { pathToFileURL } from 'node:url';
import { businessDate, selectDeliveryOrders, decideDeliveryFee } from '../agent/delivery_fee_policy.mjs';

const proxyRoot = 'http://localhost:3456';
async function proxy(route, options = {}) {
  const response = await fetch(proxyRoot + route, {...options, signal:AbortSignal.timeout(45000)});
  if (!response.ok) throw Error('Browser request failed');
  const result = await response.json();
  if (result.error) throw Error('Browser evaluation failed');
  return result;
}
async function evaluate(id, code) {
  return (await proxy(`/eval?target=${encodeURIComponent(id)}`, {method:'POST',headers:{'content-type':'text/plain; charset=utf-8'},body:code})).value;
}

// This function is serialized into a NEW background tab. Authentication stays
// inside Chrome. Only allowlisted read endpoints are used; return a projection.
export async function readProviderPage(provider, dates, ids) {
  const wanted = new Set(ids);
  const output = [];
  const yuan = v => {
    if (v === null || v === undefined || v === '' || !Number.isFinite(Number(v)) || Number(v)<0) return null;
    return Math.round(Number(v))/100;
  };
  const store = name => /宠速星宠物用品店/.test(String(name || ''));
  const request = async (url, body) => {
    // Deliberately restrict this browser helper even if a caller passes a URL.
    const parsed=new URL(url,location.href);
    if (parsed.origin!==new URL(location.href).origin || !['/query/list/','/order/detail/','/api/pc/order/manage/list','/api/pc/order/detail'].includes(parsed.pathname)) throw Error('Read endpoint not allowed');
    return await new Promise((resolve,reject) => {
      const x = new XMLHttpRequest();
      x.open(body ? 'POST' : 'GET', url); x.timeout=15000;
      if(body) x.setRequestHeader('Content-Type','application/x-www-form-urlencoded');
      x.onload=()=>{try{if(x.status!==200)throw Error('Read request failed');resolve(JSON.parse(x.responseText));}catch{reject(Error('Invalid read response'));}};
      x.onerror=x.ontimeout=()=>reject(Error('Read request unavailable')); x.send(body || null);
    });
  };
  for (const date of dates) {
    const seen = new Set();
    let complete = false, expected = null;
    for (let page=1;page<=50;page++) {
      let rows;
      if (provider==='maiyatian') {
        const template=performance.getEntriesByType('resource').map(r=>r.name).findLast(n=>new URL(n).pathname==='/query/list/');
        if(!template)throw Error('Order query has not loaded');
        const url=new URL(template);
        for (const [k,v] of Object.entries({page,page_size:20,filter_type:'all',filter_label:'all',filter_gird:'all',filter_time:1,filter_stime:60,filter_date:date,date_type:'order_date',shop_id:0,mode:'list',sort:1,goods_number:0})) url.searchParams.set(k,String(v));
        const result=await request(url.href);
        if(result.errno!==1 || !Array.isArray(result.data))throw Error('Order list unavailable');
        rows=result.data;
      } else {
        const result=await request('/api/pc/order/manage/list', new URLSearchParams({webVersion:'4.9.9-1532',pageNum:String(page),pageSize:'20',startDate:date,endDate:date,dateType:'3',sortRule:'2',goodsTypeFlag:'0'}).toString());
        if(result.code!==0 || !Array.isArray(result.data?.dataList) || !Number.isInteger(result.data.total))throw Error('Order list unavailable');
        if(Number(result.data.pageNum)!==page)throw Error('Unexpected page');
        if(expected!==null && expected!==result.data.total)throw Error('Order list changed while reading');
        expected=result.data.total; rows=result.data.dataList;
      }
      for(const r of rows) {
        const id=String(r.id ?? r.orderId);
        if(!id || id==='undefined' || seen.has(id))throw Error('Duplicate order page');
        seen.add(id);
        const orderNo=String(provider==='maiyatian'?r.source_id:r.channelOrderId);
        if(!wanted.has(orderNo)) continue;
        const correctStore=provider==='maiyatian'?store(r.shop?.name):store(r.channelShopName || r.shopName);
        if(!correctStore)throw Error('Store identity mismatch');
        const actualDate=provider==='maiyatian'?r.order_date:new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date(Number(r.channelOrderTime)*1000));
        if(actualDate!==date)throw Error('Order date mismatch');
        let delivery, amount, deliveryId, ambiguous=false, final=false, cancelled=false;
        if(provider==='maiyatian') {
          // Only the current COMPLETED dispatch is financial evidence. Old
          // cancelled attempts contain quotes, not additional paid amounts.
          const all=Array.isArray(r.deliveries)?r.deliveries:[];
          delivery=r.delivery; deliveryId=String(delivery?.id || '');
          ambiguous=all.filter(d=>Number(d.status)===50).length>1;
          final=Number(delivery?.status)===50;
          cancelled=Number(delivery?.status)===60;
          amount=yuan(delivery?.send_fee);
          if(final && amount>0) {
            const template=performance.getEntriesByType('resource').map(x=>x.name).findLast(n=>new URL(n).pathname==='/query/list/');
            const u=new URL(template);u.pathname='/order/detail/';
            const auth=new URLSearchParams();for(const k of ['f','token'])if(u.searchParams.has(k))auth.set(k,u.searchParams.get(k));
            auth.set('id',id);auth.set('detail','1');u.search=auth.toString();
            const detail=await request(u.href);
            if(detail.errno!==1 || String(detail.data?.source_id)!==orderNo)throw Error('Detail identity mismatch');
            const d=detail.data.delivery;
            ambiguous ||= String(d?.id)!==deliveryId || Number(d?.status)!==50 || yuan(d?.send_fee)!==amount;
            // The merchant page's data-detail-goods-template renders its
            // "配送费用" as delivery.send_fee / 100, even when discount_fee
            // is positive. Use that displayed charge; never subtract the
            // discount again. Other paid extras still require a bill check.
            ambiguous ||= yuan(d?.discount_fee ?? 0)===null ||
              ['premium_fee','damage_fee','tip'].some(k=>Number(d?.[k] || 0)!==0) || Number(delivery?.tip || 0)!==0;
          }
        } else {
          delivery=r.orderDeliveryInfo;deliveryId=String(r.deliveryId || '');
          final=Number(r.deliveryStatus)===50;
          cancelled=Number(r.deliveryStatus)===60;
          amount=yuan(delivery?.actualDeliveryFee);
          // Extra payments or multiple completed deliveries must not replace
          // the first-delivery column without a separate expense decision.
          ambiguous=Number(r.shipCount)>1 || ['totalExtFee','tipFee','insuredFee'].some(k=>Number(delivery?.[k] || 0)!==0);
          if(cancelled && amount>0)ambiguous=true;
        }
        output.push({provider,order_no:orderNo,delivery_id:deliveryId,
          state:ambiguous?'ambiguous':final&&amount>0?'complete':!delivery||amount===0||cancelled?'missing':'pending',
          amount:final?amount:null,date});
      }
      if(provider==='cyanclouds' && seen.size===expected){complete=true;break;}
      if(rows.length<20){if(expected!==null && seen.size!==expected)throw Error('Incomplete order list');complete=true;break;}
    }
    if(!complete)throw Error('Order paging limit reached');
  }
  return output;
}

async function readProvider(provider, dates, orders) {
  const url=provider==='maiyatian'?'https://saas.maiyatian.com/query/':'https://saas.cyanclouds.com/aggregator/saas/order#/historyOrder';
  const opened=await proxy('/new?url='+encodeURIComponent(url));
  if(!opened.targetId)throw Error('Cannot open query tab');
  try {
    let ready=false;
    for(let i=0;i<30;i++){
      ready=await evaluate(opened.targetId,provider==='maiyatian'?`!!document.querySelector('input[name="keyword"]') && performance.getEntriesByType('resource').some(r=>new URL(r.name).pathname==='/query/list/')`:`!!document.querySelector('input[placeholder="订单流水号、订单号、顾客手机号等"]')`);
      if(ready)break;
      await new Promise(r=>setTimeout(r,500));
    }
    if(!ready)throw Error('Query page unavailable or login required');
    const result=await evaluate(opened.targetId,`(async()=>{try{return {ok:true,rows:await (${readProviderPage.toString()})(${JSON.stringify(provider)},${JSON.stringify(dates)},${JSON.stringify(orders.map(r=>r.order_no))})};}catch{return {ok:false};}})()`);
    if(result?.ok!==true || !Array.isArray(result.rows))throw Error('Provider read incomplete');
    return result.rows;
  } finally {await proxy('/close?target='+encodeURIComponent(opened.targetId)).catch(()=>{});}
}

export async function extractDeliveryFees(input, reader = readProvider) {
  const candidates=selectDeliveryOrders(input.orders,input.date,input.pending || []);
  const queried=candidates.filter(r=>!r.confirmed_self);
  const dates=[...new Set(queried.map(r=>businessDate(r.date)))].sort();
  const providers={},observations=[];
  for(const provider of ['maiyatian','cyanclouds']){
    if(!queried.length){providers[provider]={ok:true};continue;}
    try{observations.push(...await reader(provider,dates,queried));providers[provider]={ok:true};}
    catch{providers[provider]={ok:false,error:'查询未完成，请检查该平台登录及订单页面；原费用保持不动'};}
  }
  const decisions=candidates.map(r=>({order_no:r.order_no,date:businessDate(r.date),...(r.confirmed_self?{state:'protected',reason:'已人工确认自配送'}:decideDeliveryFee(r,observations,providers))}));
  return {ok:true,candidate_count:candidates.length,queried_count:queried.length,providers,decisions,
    pending:decisions.filter(r=>r.retry || r.state==='conflict').map(r=>({order_no:r.order_no,date:r.date}))};
}
if(process.argv[1] && import.meta.url===pathToFileURL(process.argv[1]).href){
  try{const input=JSON.parse(fs.readFileSync(0,'utf8').replace(/^\uFEFF/,''));process.stdout.write(JSON.stringify(await extractDeliveryFees(input)));}
  catch{process.stdout.write(JSON.stringify({ok:false,error:'配送费读取输入或浏览器不可用，未写账'}));process.exitCode=1;}
}
