import assert from 'node:assert/strict';
import { classifyRefundGroup } from './meituan_order_extract.mjs';

const cases = [
  { title: '申请退货退款', expected: 'pending' },
  { title: '退款处理中', expected: 'pending' },
  { title: '商家已同意退货退款', expected: 'completed' },
  { title: '系统自动通过退款', expected: 'completed' },
  { title: '退货退款成功', expected: 'completed' },
  { title: '退款已完成', expected: 'completed' },
  { title: '退款已关闭', expected: 'closed' },
  { title: '商家驳回退款', expected: 'closed' },
  { title: '退款未完成', expected: 'unknown' },
  { title: '售后状态发生变化', expected: 'unknown' },
];

for (const item of cases) {
  const actual = classifyRefundGroup([{ recordId: 'r1', title: item.title }]);
  assert.equal(actual.state, item.expected, `${item.title} should be ${item.expected}`);
  assert.ok(actual.statusTexts.includes(item.title), `${item.title} should remain available as evidence`);
}

const grouped = classifyRefundGroup([
  { recordId: 'r2', title: '申请部分退款' },
  { recordId: 'r2', title: '退款成功' },
]);
assert.equal(grouped.state, 'completed', 'a later completion record must override the earlier request state');

const explicit = classifyRefundGroup([{ recordId: 'r3', statusText: '售后状态', isCompleted: true }]);
assert.equal(explicit.state, 'completed', 'an explicit completion flag must be supported');

process.stdout.write(JSON.stringify({ ok: true, tested_cases: cases.length + 2 }));
