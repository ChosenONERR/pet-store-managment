import assert from 'node:assert/strict';
import fs from 'node:fs';
import { normalizeOrderStatus } from './meituan_order_extract.mjs';

const cases = [
  ['配送取消', '处理中'], ['取消配送', '处理中'], ['配送已取消', '处理中'],
  ['退款已取消', '处理中'], ['取消申请处理中', '处理中'], ['待用户收货', '处理中'],
  ['已取消', '已取消'], ['订单已取消', '已取消'], ['订单取消', '已取消'],
  ['用户取消订单', '已取消'], ['顾客已取消', '已取消'], ['商家取消订单', '已取消'],
  ['订单已取消 配送取消', '已取消'], ['用户已收货', '已完成'],
  ['订单已完成', '已完成'], ['配送完成', '已完成'], ['已送达', '已完成'], ['', '处理中'],
];
for (const [raw, expected] of cases) assert.equal(normalizeOrderStatus(raw), expected, raw);
// Reproduce the incident and prove delivery cancellation is no longer terminal.
assert.equal(/取消/.test('配送取消'), true);
assert.equal(normalizeOrderStatus('配送取消'), '处理中');
assert.equal(normalizeOrderStatus('用户已收货'), '已完成');
const source = fs.readFileSync(new URL('./meituan_order_extract.mjs', import.meta.url), 'utf8');
assert.ok(source.includes('${normalizeOrderStatus.toString()}'));
assert.ok(source.includes('status: normalizeOrderStatus(rawStatus)'));
process.stdout.write(JSON.stringify({ok:true,cases:cases.length+5,external_writes:0}));
