import fs from 'node:fs';
import { pathToFileURL } from 'node:url';
import { businessDate, selectDeliveryOrders, decideDeliveryFee, resolveCyanDispatchHistory,
  wantsMaiyatianServiceFee, includeMaiyatianServiceFee } from '../agent/delivery_fee_policy.mjs';

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
export async function readProviderPage(provider, dates, ids, resolveHistory = resolveCyanDispatchHistory) {
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
    if (parsed.origin!==new URL(location.href).origin || !['/query/list/','/order/detail/','/api/pc/order/manage/list','/api/pc/order/detail','/api/pc/delivery/progressTrack'].includes(parsed.pathname)) throw Error('Read endpoint not allowed');
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
        let delivery, amount, deliveryId, ambiguityReason, historyEvidence, ambiguous=false, final=false, cancelled=false;
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
          ambiguous=['totalExtFee','tipFee','insuredFee'].some(k=>Number(delivery?.[k] || 0)!==0);
          if (Number(r.shipCount)>1) {
            const track=await request('/api/pc/delivery/progressTrack',new URLSearchParams({orderId:id}).toString());
            if (track.code!==0) throw Error('Delivery history unavailable');
            const resolved=resolveHistory(r,track.data);
            if (resolved.state==='ambiguous') { ambiguous=true; ambiguityReason=resolved.reason; }
            else {
              amount=resolved.amount;final=resolved.state==='complete';cancelled=resolved.state==='missing';
              historyEvidence=resolved.history;
            }
          }
          if(cancelled && amount>0)ambiguous=true;
        }
        output.push({provider,order_no:orderNo,delivery_id:deliveryId,
          state:ambiguous?'ambiguous':final&&amount>0?'complete':!delivery||amount===0||cancelled?'missing':'pending',
          amount:final?amount:null,date,
          ...(ambiguityReason?{reason:ambiguityReason}:{}),...(historyEvidence?{history:historyEvidence}:{})});
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
    const result=await evaluate(opened.targetId,`(async()=>{try{return {ok:true,rows:await (${readProviderPage.toString()})(${JSON.stringify(provider)},${JSON.stringify(dates)},${JSON.stringify(orders.map(r=>r.order_no))},(${resolveCyanDispatchHistory.toString()}))};}catch{return {ok:false};}})()`);
    if(result?.ok!==true || !Array.isArray(result.rows))throw Error('Provider read incomplete');
    return result.rows;
  } finally {await proxy('/close?target='+encodeURIComponent(opened.targetId)).catch(()=>{});}
}

// DOM column labels are part of the evidence: a recharge or dispatch quote
// must never be treated as a paid technical service entry.
export function verifyServiceSettlement(summary, detailRows) {
  const cents = value => {
    const s = String(value ?? '').trim();
    if (!/^\d+(?:\.\d{1,2})?$/.test(s)) throw Error('Unexpected settlement amount');
    const n = Math.round(Number(s) * 100);
    if (!Number.isSafeInteger(n) || n <= 0) throw Error('Invalid settlement amount');
    return n;
  };
  const count = /^技术服务结算\s*[（(]\s*(\d+)\s*笔\s*[）)]$/.exec(summary.description)?.[1];
  if (!count || Number(count) !== detailRows.length || detailRows.length === 0) throw Error('Incomplete technical settlement detail');
  const details = detailRows.map(([type, orderNo, , , status, amount]) => {
    if (type !== '技术服务' || status !== '已完成' || !/^\d{16,22}$/.test(orderNo)) throw Error('Unknown technical settlement type or status');
    return {order_no:orderNo, settlement_date:summary.date, cents:cents(amount)};
  });
  if (details.reduce((sum, row) => sum + row.cents, 0) !== cents(summary.amount)) throw Error('Technical settlement totals disagree');
  return details;
}

export async function readMaiyatianServiceFees(orders, date) {
  const earliest = orders.map(order => order.date).sort()[0];
  const wanted = new Set(orders.map(order => order.order_no));
  const target = (await proxy('/new?url='+encodeURIComponent('https://saas.maiyatian.com/service/balance/'))).targetId;
  if (!target) throw Error('Cannot open settlement tab');
  const waitFor = async (code) => {
    for (let i=0; i<18; i++) {
      const value=await evaluate(target, code);
      if (value != null) return value;
      await new Promise(resolve=>setTimeout(resolve,350));
    }
    throw Error('Settlement page or detail unavailable');
  };
  try {
    const summaries = await waitFor(`(()=>{
      const table=[...document.querySelectorAll('table')].find(t=>
        [...t.querySelectorAll('th')].map(th=>th.innerText.trim()).join('|')==='日期|类型|描述|金额/元|操作');
      return table ? [...table.querySelectorAll('tbody tr')].map(r=>
        [...r.cells].map(c=>c.innerText.trim())).filter(r=>r[1]==='技术服务') : null;
    })()`);
    const dates = new Set();
    const results = [];
    for (const [day, , description, amount] of summaries) {
      if (day < earliest || day > date) continue;
      if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || dates.has(day)) throw Error('Technical settlement date is not unique');
      dates.add(day);
      const opened = await evaluate(target, `(() => {
        const rows=[...document.querySelectorAll('tr')].filter(r=>
          r.cells[0]?.innerText.trim()===${JSON.stringify(day)} && r.cells[1]?.innerText.trim()==='技术服务');
        const detail=rows.length===1 && rows[0].querySelector('a.js_detail');
        if (!detail || document.querySelector('#balance-detail table')?.getClientRects().length) return false;
        // The page reopens the previous detail before its asynchronous response arrives.
        globalThis.petstarPreviousSettlement = document.querySelector('#balance-detail table')?.innerHTML;
        detail.click();
        return true;
      })()`);
      if (!opened) throw Error('Cannot open an unambiguous settlement detail');
      const detail = await waitFor(`(() => {
        const table=document.querySelector('#balance-detail table');
        if (!table?.getClientRects().length ||
            table.innerHTML===globalThis.petstarPreviousSettlement ||
            [...table.querySelectorAll('th')].map(th=>th.innerText.trim()).join('|')!=='类型|订单号|订单信息|完单时间|服务状态|金额') return null;
        const rows=[...table.rows].filter(r=>r.cells[0]?.tagName==='TD').map(r=>[...r.cells].map(c=>c.innerText.trim()));
        return rows.length && rows.every(r=>r.length===6) ? rows : null;
      })()`);
      results.push(...verifyServiceSettlement({date:day, description, amount}, detail).filter(r=>wanted.has(r.order_no)));
      const closed = await evaluate(target, `(() => {
        const close=document.querySelector('#balance-detail')?.closest('.full-page')?.querySelector('.close-page');
        if (!close) return false;
        close.click();
        return true;
      })()`);
      if (!closed) throw Error('Cannot close settlement detail');
      await waitFor("document.querySelector('#balance-detail table')?.getClientRects().length ? null : true");
    }
    return results;
  } finally {await proxy('/close?target='+encodeURIComponent(target)).catch(()=>{});}
}

export async function extractDeliveryFees(input, reader = readProvider, serviceReader = readMaiyatianServiceFees) {
  const candidates=selectDeliveryOrders(input.orders,input.date,input.pending || []);
  const queried=candidates.filter(r=>!r.confirmed_self);
  const dates=[...new Set(queried.map(r=>businessDate(r.date)))].sort();
  const providers={},observations=[];
  for(const provider of ['maiyatian','cyanclouds']){
    if(!queried.length){providers[provider]={ok:true};continue;}
    try{observations.push(...await reader(provider,dates,queried));providers[provider]={ok:true};}
    catch{providers[provider]={ok:false,error:'查询未完成，请检查该平台登录及订单页面；原费用保持不动'};}
  }
  const baseDecisions=candidates.map(r=>({order_no:r.order_no,date:businessDate(r.date),...(r.confirmed_self?{state:'protected',reason:'已人工确认自配送'}:decideDeliveryFee(r,observations,providers))}));
  const serviceOrders=baseDecisions.filter(r=>wantsMaiyatianServiceFee(r,input.date));
  let settlements=null, serviceSettlementStatus='not_needed';
  if (serviceOrders.length) {
    try {settlements=await serviceReader(serviceOrders,input.date);serviceSettlementStatus='read';}
    catch {serviceSettlementStatus='deferred';}
  }
  const decisions=baseDecisions.map((row,i)=>includeMaiyatianServiceFee(candidates[i],row,input.date,settlements));
  return {ok:true,candidate_count:candidates.length,queried_count:queried.length,providers,tech_service_settlement:serviceSettlementStatus,decisions,
    pending:decisions.filter(r=>r.retry || r.state==='conflict').map(r=>({order_no:r.order_no,date:r.date}))};
}
if(process.argv[1] && import.meta.url===pathToFileURL(process.argv[1]).href){
  try{const input=JSON.parse(fs.readFileSync(0,'utf8').replace(/^\uFEFF/,''));process.stdout.write(JSON.stringify(await extractDeliveryFees(input)));}
  catch{process.stdout.write(JSON.stringify({ok:false,error:'配送费读取输入或浏览器不可用，未写账'}));process.exitCode=1;}
}
