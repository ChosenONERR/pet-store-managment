// Pure policy: no browser, credentials, files or financial writes.
export function businessDate(value) {
  if (typeof value === 'number' || /^\d+(?:\.\d+)?$/.test(String(value ?? ''))) {
    return new Date(Date.UTC(1899, 11, 30) + Math.floor(Number(value)) * 86400000).toISOString().slice(0, 10);
  }
  return String(value ?? '').slice(0, 10);
}

export function selectDeliveryOrders(orders, date, pending = []) {
  const waiting = new Set(pending.map(r => String(r.order_no)));
  const previousDay = new Date(Date.parse(date+'T00:00:00Z')-86400000).toISOString().slice(0,10);
  const seen = new Set();
  return orders.filter(r => {
    if (!/^\d{16,22}$/.test(r.order_no) || seen.has(r.order_no)) throw Error('Invalid or duplicate order identity');
    seen.add(r.order_no);
    const day=businessDate(r.date);
    const recentUnresolved=day===previousDay && !r.confirmed_self && !r.protected &&
      (r.current_fee == null || Number(r.current_fee)===0);
    return r.platform === '美团' && !r.intervention && !/取消/.test(r.status || '') &&
      (day === date || (day < date && (waiting.has(r.order_no) || recentUnresolved)));
  });
}

// One send round can race several carriers. Cancelled waybills retain quotes;
// only a completed waybill plus explicit zero cancellation/extras is safe here.
// Kept self-contained because the read adapter serializes it into Chrome.
export function resolveCyanDispatchHistory(order, history) {
  const bad = reason => ({state:'ambiguous', reason});
  const cents = value => {
    if (!['number','string'].includes(typeof value) || String(value).trim()==='') return null;
    const number=Number(value);
    return Number.isSafeInteger(number) && number>=0 ? number : null;
  };
  const identity = value => {
    if (typeof value==='number' && (!Number.isSafeInteger(value) || value<0)) return '';
    return ['number','string'].includes(typeof value) ? String(value).trim() : '';
  };
  const rounds=history?.deliveryList, count=cents(order.shipCount);
  if (!Array.isArray(rounds) || !count || rounds.length!==count) return bad('配送历史轮次不完整，保留原费用并继续核对');
  if (!['totalExtFee','tipFee','insuredFee'].every(k=>cents(order.orderDeliveryInfo?.[k])===0)) return bad('配送记录有附加费用或费用缺失，需单独核对');
  const roundIds=new Set(), waybillIds=new Set(), completed=[];
  let cancelledCount=0;
  for (const round of rounds) {
    const roundId=identity(round.deliveryId), status=Number(round.status);
    if (!roundId || roundIds.has(roundId) || ![50,99].includes(status) || !Array.isArray(round.waybillList) || !round.waybillList.length) return bad('配送历史身份、状态或运单不完整，保留原费用');
    roundIds.add(roundId);
    let completedInRound=0;
    for (const waybill of round.waybillList) {
      const waybillId=identity(waybill.waybillId), waybillStatus=Number(waybill.status);
      if (!waybillId || waybillIds.has(waybillId) || ![50,99].includes(waybillStatus)) return bad('配送历史运单重复或状态未结清，保留原费用');
      waybillIds.add(waybillId);
      if (!['cancelFee','totalExtFee','insuredFee','tipAmount'].every(k=>cents(waybill[k])===0) ||
          (waybill.allowanceFee!==undefined && cents(waybill.allowanceFee)!==0)) return bad('配送历史有取消费、附加费或费用缺失，需单独核对');
      if (waybillStatus===50) {
        const amount=cents(waybill.actualDeliveryFee);
        if (status!==50 || amount===null) return bad('已完成配送的实付或状态不一致，保留原费用');
        completed.push({roundId,waybillId,amount});
        completedInRound++;
      } else {
        // actualDeliveryFee on status 99 is a cancelled quote, not a payment.
        cancelledCount++;
      }
    }
    if (status===50 && completedInRound!==1) return bad('同轮存在多笔完成配送或缺失完成记录，需核对');
  }
  if (!roundIds.has(identity(order.deliveryId))) return bad('当前派单与配送历史身份不一致，保留原费用');
  if (completed.length>1) return bad('存在多笔完成配送，不自动合并到首次配送费');
  const evidence={rounds:rounds.length,completed:completed.length,cancelled:cancelledCount,cancel_fee:0};
  if (!completed.length) {
    if (![60,99].includes(Number(order.deliveryStatus))) return bad('订单状态与已取消配送历史不一致，保留原费用');
    return {state:'missing',amount:0,history:evidence};
  }
  const paid=completed[0];
  if (Number(order.deliveryStatus)!==50 || paid.roundId!==identity(order.deliveryId) ||
      paid.waybillId!==identity(order.waybillId) || cents(order.orderDeliveryInfo?.actualDeliveryFee)!==paid.amount) return bad('当前派单实付与完成运单不一致，保留原费用');
  return {state:'complete',amount:paid.amount/100,history:evidence};
}

export function decideDeliveryFee(order, observations, providers) {
  const rows = observations.filter(r => r.order_no === order.order_no);
  const positive = [];
  let awaitingDispatch = false;
  for (const provider of ['maiyatian', 'cyanclouds']) {
    const found = rows.filter(r => r.provider === provider);
    if (!providers[provider]?.ok) return {state:'deferred', retry:true, reason:`${provider}查询失败，保留原费用并后台重试`};
    if (found.length > 1) return {state:'conflict',reason:'配送记录不唯一或有多次派送/附加费用，需核对'};
    const row = found[0];
    if (row?.state==='ambiguous') return {state:'conflict',reason:row.reason || '配送记录不唯一或有多次派送/附加费用，需核对'};
    if (row?.state === 'complete' && Number.isFinite(row.amount) && row.amount > 0) positive.push(row);
    else if (row && row.state !== 'missing') awaitingDispatch = true;
  }
  if (positive.length > 1) return {state:'conflict',reason:'两个平台都有实际配送费，不自动相加或覆盖'};
  const source = positive[0];
  if (order.protected) {
    if (source && Math.abs(Number(order.current_fee) - source.amount) > 0.0049) return {state:'conflict',reason:'平台实付与人工确认费用不一致，保留人工值'};
    return {state:'protected',reason:'保留人工确认费用',source};
  }
  if (!source) {
    // No positive evidence must never erase a previously known charge. Default
    // self delivery is provisional, not a manual lock: today/yesterday are reread.
    if (Number(order.current_fee) > 0) return {state:'retained',retry:awaitingDispatch,reason:'未查到新的正数实付，保留已有配送费'};
    return {state:'default_self',amount:0,changed:order.current_fee == null,retry:awaitingDispatch,
      reason:'未查到正数实付，按店主规则默认自配送0元；后续继续自动查询'};
  }
  return {state:'confirmed', amount:source.amount, provider:source.provider, delivery_id:source.delivery_id,
    changed:order.current_fee == null || Math.abs(Number(order.current_fee)-source.amount)>0.0049};
}
