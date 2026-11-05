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

export function decideDeliveryFee(order, observations, providers) {
  const rows = observations.filter(r => r.order_no === order.order_no);
  const positive = [];
  let awaitingDispatch = false;
  for (const provider of ['maiyatian', 'cyanclouds']) {
    const found = rows.filter(r => r.provider === provider);
    if (!providers[provider]?.ok) return {state:'pending', reason:`${provider}查询失败，保留原费用`};
    if (found.length > 1 || found.some(r => r.state === 'ambiguous')) return {state:'conflict',reason:'配送记录不唯一或有多次派送/附加费用，需核对'};
    const row = found[0];
    if (row?.state === 'complete' && Number.isFinite(row.amount) && row.amount > 0) positive.push(row);
    else if (row && row.state !== 'missing') awaitingDispatch = true;
  }
  if (positive.length > 1) return {state:'conflict',reason:'两个平台都有实际配送费，不自动相加或覆盖'};
  const source = positive[0];
  if (order.protected) {
    if (source && Math.abs(Number(order.current_fee) - source.amount) > 0.0049) return {state:'conflict',reason:'平台实付与人工确认费用不一致，保留人工值'};
    return {state:'protected',reason:'保留人工确认费用',source};
  }
  if (!source) return {state:'pending',reason:awaitingDispatch?'配送尚未完成或实付不明确，下一小时继续查':'两边未查到正数实付，请确认是否自配送；下一小时继续查'};
  return {state:'confirmed', amount:source.amount, provider:source.provider, delivery_id:source.delivery_id,
    changed:order.current_fee == null || Math.abs(Number(order.current_fee)-source.amount)>0.0049};
}
