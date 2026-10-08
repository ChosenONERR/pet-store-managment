const PROXY = 'http://localhost:3456';
const ORDER_URL = 'https://shangoue.meituan.com/#/page/orderbusiness#/order/history';
// Chrome/CDP can pause briefly while Meituan renders a large order page.
// Thirty seconds produced repeated false failures even though the next hourly
// run succeeded, so allow one slow render before the outer retry opens a new tab.
const PROXY_TIMEOUT_MS = Number(process.env.MEITUAN_PROXY_TIMEOUT_MS || 60000);

function parseArgs(argv) {
  const result = {};
  for (let i = 2; i < argv.length; i += 1) {
    const key = argv[i];
    if (key === '--start' || key === '--end') result[key.slice(2)] = argv[++i];
  }
  return result;
}

function assertDate(value, name) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) {
    throw new Error(`${name} must use YYYY-MM-DD`);
  }
}

async function proxy(path, options = {}) {
  let response;
  try {
    response = await fetch(`${PROXY}${path}`, {
      ...options,
      signal: options.signal || AbortSignal.timeout(PROXY_TIMEOUT_MS),
    });
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    throw new Error(`CDP proxy request timed out or disconnected at ${path.slice(0, 80)} after ${PROXY_TIMEOUT_MS}ms: ${detail}`);
  }
  const body = await response.text();
  if (!response.ok) throw new Error(`CDP proxy request failed: ${response.status} ${body}`);
  const data = JSON.parse(body);
  if (data.error) throw new Error(data.error);
  return data;
}

async function evaluate(targetId, expression) {
  const result = await proxy(`/eval?target=${encodeURIComponent(targetId)}`, {
    method: 'POST',
    headers: { 'content-type': 'text/plain; charset=utf-8' },
    body: expression,
  });
  return result.value;
}

async function waitFor(targetId, expression, timeoutMs = 45000) {
  const started = Date.now();
  let lastError = null;
  while (Date.now() - started < timeoutMs) {
    try {
      if (await evaluate(targetId, expression)) return;
      lastError = null;
    } catch (error) {
      // Runtime.evaluate can briefly time out while Meituan replaces the iframe.
      // All wait expressions are read-only, so polling them again is safe.
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  const suffix = lastError ? `; last evaluation error: ${lastError.message || String(lastError)}` : '';
  throw new Error(`Timed out waiting for page state: ${expression}${suffix}`);
}

async function getOrderPageState(targetId) {
  return evaluate(targetId, `(() => {
    const iframe = document.querySelector('#hashframe');
    const bodyText = document.body?.innerText || '';
    const safeUrl = (value) => {
      try {
        const parsed = new URL(value, location.href);
        return parsed.origin + parsed.pathname + parsed.hash.split('?')[0];
      } catch {
        return null;
      }
    };
    return {
      topUrl: safeUrl(location.href),
      title: document.title,
      topReadyState: document.readyState,
      iframePresent: Boolean(iframe),
      iframeUrl: iframe?.src ? safeUrl(iframe.src) : null,
      iframeReadyState: iframe?.contentDocument?.readyState || null,
      dateInputPresent: Boolean(iframe?.contentDocument?.querySelector('input[placeholder="开始时间"]')),
      loginHint: /登录|扫码登录|重新登录/.test(bodyText),
    };
  })()`);
}

function isAuthenticationRequired(state) {
  const url = String(state?.topUrl || '');
  return /\/login(?:[_/#?]|$)|login_gw/i.test(url)
    || Boolean(state?.loginHint && !state?.iframePresent);
}

function throwAuthenticationRequired(state) {
  // Stable error code lets the outer scheduler stop retrying and suppress
  // duplicate alerts until a later authenticated order read succeeds.
  throw new Error(`MEITUAN_AUTH_REQUIRED: 美团商家后台登录态已失效，请在当前 Chrome 重新登录；state=${JSON.stringify(state)}`);
}

async function waitForOrderReadyOrAuth(targetId, timeoutMs) {
  const started = Date.now();
  let lastState = null;
  let lastError = null;
  while (Date.now() - started < timeoutMs) {
    try {
      lastState = await getOrderPageState(targetId);
      lastError = null;
      if (isAuthenticationRequired(lastState)) throwAuthenticationRequired(lastState);
      if (lastState?.dateInputPresent) return lastState;
    } catch (error) {
      if (/MEITUAN_AUTH_REQUIRED/.test(error?.message || '')) throw error;
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  const suffix = lastError ? `; last evaluation error: ${lastError.message || String(lastError)}` : '';
  throw new Error(`Timed out waiting for Meituan order page; state=${JSON.stringify(lastState)}${suffix}`);
}

async function waitForOrderPageReady(targetId) {
  try {
    await waitForOrderReadyOrAuth(targetId, 30000);
    return;
  } catch (firstError) {
    if (/MEITUAN_AUTH_REQUIRED/.test(firstError?.message || '')) throw firstError;
    // 商家端偶尔只加载外层壳而不创建订单 iframe。对当前后台 tab 原地重载一次，
    // 比外层任务反复创建新 tab 更能恢复卡住的前端路由；全程只读，不提交页面操作。
    await proxy(`/navigate?target=${encodeURIComponent(targetId)}&url=${encodeURIComponent(ORDER_URL)}`);
    await new Promise((resolve) => setTimeout(resolve, 2500));
    try {
      await waitForOrderReadyOrAuth(targetId, 45000);
      return;
    } catch (secondError) {
      if (/MEITUAN_AUTH_REQUIRED/.test(secondError?.message || '')) throw secondError;
      const state = await getOrderPageState(targetId)
        .catch((error) => ({ diagnosticError: error.message || String(error) }));
      throw new Error(`Meituan order page was not ready after one safe reload; state=${JSON.stringify(state)}; first=${firstError.message}; second=${secondError.message}`);
    }
  }
}

async function expandOrderCards(targetId) {
  await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    if (!root) return 0;
    const buttons = [...root.querySelectorAll('.btn-toggle-record')]
      .filter((button) => /展开/.test(button.innerText || ''));
    for (const button of buttons) {
      button.dispatchEvent(new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        view: root.defaultView,
      }));
    }
    return buttons.length;
  })()`);
  const readyExpression = `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    const bodyText = root?.body?.innerText || '';
    // “0 单”是正常结果，不应继续等待并误报“订单卡片未完整展开”。
    if (/共\\s*0\\s*单/.test(bodyText)) return true;
    const cards = [...(root?.querySelectorAll('[class^="order-card_"]') || [])];
    return cards.length > 0 && cards.every((card) => {
      const expectedText = (card.querySelector('[class^="productCardTitle_"]')?.innerText || card.innerText || '');
      const expected = Number((expectedText.match(/(\\d+)件商品/) || [])[1] || 0);
      const rows = [...card.querySelectorAll('[class^="goods-item-row_"]')];
      const actual = rows.reduce((total, row) => {
        const quantityText = [...row.querySelectorAll('.goods-num')]
          .map((node) => node.innerText.trim())
          .find((value) => /^x\\d+(?:\\.\\d+)?$/i.test(value));
        return total + Number((quantityText || 'x0').replace(/^x/i, ''));
      }, 0);
      const hasSettlement = Boolean(card.querySelector('.settlement-info'));
      return expected > 0 && actual >= expected && hasSettlement;
    });
  })()`;
  try {
    await waitFor(targetId, readyExpression, 45000);
  } catch (error) {
    const states = await evaluate(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      return [...(root?.querySelectorAll('[class^="order-card_"]') || [])].map((card, index) => {
        const text = card.innerText || '';
        const basic = card.querySelector('[class^="orderBasicInfo_"]')?.innerText || '';
        const expectedText = card.querySelector('[class^="productCardTitle_"]')?.innerText || text;
        return {
          index: index + 1,
          orderNo: (basic.match(/订单编号：?(\\d{18,20})/) || [])[1] || null,
          status: card.querySelector('[class^="baseInfoRight_"]')?.innerText.trim() || '',
          expectedItems: Number((expectedText.match(/(\\d+)件商品/) || [])[1] || 0),
          actualRows: card.querySelectorAll('[class^="goods-item-row_"]').length,
          actualItems: [...card.querySelectorAll('[class^="goods-item-row_"]')].reduce((total, row) => {
            const quantityText = [...row.querySelectorAll('.goods-num')]
              .map((node) => node.innerText.trim())
              .find((value) => /^x\\d+(?:\\.\\d+)?$/i.test(value));
            return total + Number((quantityText || 'x0').replace(/^x/i, ''));
          }, 0),
          hasSettlement: Boolean(card.querySelector('.settlement-info')),
          toggleText: card.querySelector('.btn-toggle-record')?.innerText.trim() || '',
        };
      });
    })()`);
    throw new Error(`Order cards did not become complete: ${JSON.stringify(states)}`, { cause: error });
  }
}

async function selectDateRange(targetId, start, end) {
  const startDay = Number(start.slice(-2));
  const endDay = Number(end.slice(-2));
  const sameMonth = start.slice(0, 7) === end.slice(0, 7);
  if (!sameMonth) throw new Error('The current extractor supports one calendar month per run');
  const spanDays = Math.round((Date.parse(`${end}T00:00:00+08:00`) - Date.parse(`${start}T00:00:00+08:00`)) / 86400000);
  if (spanDays < 0 || spanDays > 7) throw new Error('Meituan allows at most 8 inclusive calendar days per query');
  const directChanged = await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    const input = root?.querySelector('#range input');
    if (!input) return false;
    const fiberKey = Object.keys(input).find((key) => key.startsWith('__reactFiber'));
    let fiber = fiberKey ? input[fiberKey] : null;
    while (fiber) {
      const props = fiber.memoizedProps;
      if (props?.value && typeof props.value === 'object' && 'startDate' in props.value && 'endDate' in props.value && typeof props.onChange === 'function') {
        props.onChange({ startDate: '${start}', endDate: '${end}' });
        return true;
      }
      fiber = fiber.return;
    }
    return false;
  })()`);
  if (directChanged) {
    await waitFor(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      const values = [...(root?.querySelectorAll('#range input') || [])].map((node) => node.value);
      return values[0] === '${start}' && values[1] === '${end}';
    })()`);
    return;
  }
  const year = start.slice(0, 4);
  const monthNames = ['一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'];
  const monthName = monthNames[Number(start.slice(5, 7)) - 1];
  const calendarAlreadyOpen = await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    return [...(root?.querySelectorAll('.roo-datepicker-days') || [])].some((panel) => panel.offsetParent !== null);
  })()`);
  if (!calendarAlreadyOpen) {
    await evaluate(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      if (!root) throw new Error('Meituan order iframe not found');
      const startInput = root.querySelector('input[placeholder="开始时间"]');
      if (!startInput) throw new Error('Start date input not found');
      startInput.click();
      return true;
    })()`);
  }
  const targetMonthIndex = Number(year) * 12 + Number(start.slice(5, 7)) - 1;
  let targetMonthVisible = false;
  for (let attempt = 0; attempt < 24; attempt += 1) {
    const navigation = await evaluate(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      const panels = [...(root?.querySelectorAll('.roo-datepicker-days') || [])].filter((panel) => panel.offsetParent !== null);
      const targetText = '${year}${monthName}';
      if (panels.some((panel) => panel.innerText.includes(targetText))) return { found: true };
      if (!panels.length) return { found: false, reason: 'no-visible-panel' };
      const monthNames = ${JSON.stringify(monthNames)};
      const readMonthIndex = (panel) => {
        const values = [...panel.querySelectorAll('.roo-datepicker-header .action-select a')].map((node) => node.textContent.trim());
        const panelYear = Number(values[0]);
        const panelMonth = monthNames.indexOf(values[1]);
        return panelYear * 12 + panelMonth;
      };
      const firstIndex = readMonthIndex(panels[0]);
      const lastIndex = readMonthIndex(panels[panels.length - 1]);
      if (${targetMonthIndex} < firstIndex) {
        const previous = panels[0].querySelector('.action-pre .roo-icon-chevron-left');
        if (!previous) return { found: false, reason: 'previous-button-missing' };
        previous.click();
        return { found: false, moved: 'previous' };
      }
      if (${targetMonthIndex} > lastIndex) {
        const next = panels[panels.length - 1].querySelector('.action-next .roo-icon-chevron-right');
        if (!next) return { found: false, reason: 'next-button-missing' };
        next.click();
        return { found: false, moved: 'next' };
      }
      return { found: false, reason: 'target-between-visible-panels' };
    })()`);
    if (navigation?.found) {
      targetMonthVisible = true;
      break;
    }
    if (navigation?.reason && navigation.reason !== 'target-between-visible-panels') {
      throw new Error(`Calendar navigation failed: ${navigation.reason}`);
    }
    await new Promise((resolve) => setTimeout(resolve, 400));
  }
  if (!targetMonthVisible) throw new Error(`Calendar target month not visible: ${year}${monthName}`);
  await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe').contentDocument;
    const panel = [...root.querySelectorAll('.roo-datepicker-days')].find((node) => node.offsetParent !== null && node.innerText.includes('${year}${monthName}'));
    if (!panel) throw new Error('Calendar panel not found');
    const currentDays = [...panel.querySelectorAll('td:not(.old):not(.new)')];
    const startDay = currentDays.find((el) => el.textContent.trim() === '${startDay}');
    if (!startDay) throw new Error('Start calendar day not found');
    const propsKey = Object.keys(startDay).find((key) => key.startsWith('__reactProps'));
    const handler = propsKey && startDay[propsKey]?.onClick;
    if (typeof handler !== 'function') throw new Error('Start calendar day handler not found');
    handler({ stopPropagation() {} });
    return true;
  })()`);
  await waitFor(targetId, `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    const panel = [...(root?.querySelectorAll('.roo-datepicker-days') || [])].find((node) => node.offsetParent !== null && node.innerText.includes('${year}${monthName}'));
    return [...(panel?.querySelectorAll('td.active') || [])].some((cell) => cell.textContent.trim() === '${startDay}');
  })()`);
  await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe').contentDocument;
    const refreshedPanel = [...root.querySelectorAll('.roo-datepicker-days')].find((node) => node.offsetParent !== null && node.innerText.includes('${year}${monthName}'));
    if (!refreshedPanel) throw new Error('Calendar panel disappeared');
    const endDay = [...refreshedPanel.querySelectorAll('td:not(.old):not(.new)')].find((el) => el.textContent.trim() === '${endDay}');
    if (!endDay) throw new Error('End calendar day not found');
    const propsKey = Object.keys(endDay).find((key) => key.startsWith('__reactProps'));
    const handler = propsKey && endDay[propsKey]?.onClick;
    if (typeof handler !== 'function') throw new Error('End calendar day handler not found');
    handler({ stopPropagation() {} });
    return true;
  })()`);
}

const extractPageExpression = `(() => {
  const root = document.querySelector('#hashframe')?.contentDocument;
  if (!root) throw new Error('Meituan order iframe not found');
  const money = (text) => {
    const match = String(text || '').replace(/,/g, '').match(/-?[￥¥]?\\s*(-?\\d+(?:\\.\\d+)?)/);
    return match ? Number(match[1]) : null;
  };
  const normalizeStatus = (raw) => {
    if (/取消/.test(raw)) return '已取消';
    if (/收货|完成/.test(raw)) return '已完成';
    return raw || '处理中';
  };
  const normalizeProductName = (value) => String(value || '')
    .replace(/\s*[\(（][^\(（\)）]*[\)）]\s*$/, '')
    .replace(/[\s\p{P}\p{S}]+/gu, '')
    .toLowerCase();
  const getOrderModel = (card) => {
    const anchor = card.querySelector('.order-card-refund-info-container') || card;
    const fiberKey = Object.keys(anchor).find((key) => key.startsWith('__reactFiber'));
    let fiber = fiberKey ? anchor[fiberKey] : null;
    for (let depth = 0; fiber && depth < 30; depth += 1, fiber = fiber.return) {
      const props = fiber.memoizedProps;
      if (props?.order?.unifiedBasicInfo || props?.order?.refundInfo) return props.order;
      if (props?.orderData?.unifiedBasicInfo || props?.orderData?.refundInfo) return props.orderData;
    }
    return null;
  };
  const extractRefunds = (orderModel) => {
    const records = orderModel?.refundInfo?.refundInfoList || [];
    if (!records.length) return [];
    const foodRows = (orderModel?.foodInfo?.foodDetails || [])
      .flatMap((group) => group?.details || []);
    const grouped = new Map();
    for (const record of records) {
      if (!record?.recordId) continue;
      if (!grouped.has(String(record.recordId))) grouped.set(String(record.recordId), []);
      grouped.get(String(record.recordId)).push(record);
    }
    return [...grouped.entries()].map(([recordId, group]) => {
      const completedRecord = group.find((record) => /(?:同意|完成|成功).*退款|退款.*(?:完成|成功)/.test(record.title || ''));
      const requestRecord = group.find((record) => record.moneyVO || record.refundFoodInfo)
        || group.find((record) => /申请.*退款/.test(record.title || ''));
      const refundItems = requestRecord?.refundFoodInfo?.refundFoodItemInfoList || [];
      const items = refundItems.map((refundItem) => {
        const refundName = normalizeProductName(refundItem.foodName);
        const matches = foodRows.filter((food) => {
          const foodName = normalizeProductName(food.foodName);
          return foodName === refundName || foodName.startsWith(refundName) || refundName.startsWith(foodName);
        });
        const food = matches.length === 1 ? matches[0] : null;
        return {
          goodsCode: String(food?.skuCode || food?.locatorCode || '').trim().toLowerCase(),
          productId: food?.foodId == null ? '' : String(food.foodId),
          title: String(refundItem.foodName || food?.foodName || '').trim(),
          spec: String(food?.specification || '').trim(),
          quantity: Number(refundItem.foodRefundCount || 0),
          refundAmount: money(refundItem.foodRefundTotalPriceText || refundItem.foodPriceText || '') || 0,
        };
      });
      return {
        recordId,
        completed: Boolean(completedRecord),
        refundType: /部分退款/.test(completedRecord?.title || requestRecord?.title || '') ? '部分退款' : '退款',
        completedAt: completedRecord?.timeFormat || null,
        requestedAt: requestRecord?.timeFormat || null,
        amount: money(requestRecord?.moneyVO?.moneyText || '')
          ?? items.reduce((total, item) => total + item.refundAmount, 0),
        reason: requestRecord?.reasonVO?.reason || '',
        items,
      };
    }).filter((refund) => refund.completed && refund.amount > 0);
  };
  const selectedYear = (root.querySelector('#range input')?.value || '').slice(0, 4) || String(new Date().getFullYear());
  const cards = [...root.querySelectorAll('[class^="order-card_"]')];
  const orders = cards.map((card) => {
    const text = card.innerText;
    const orderModel = getOrderModel(card);
    const basic = card.querySelector('[class^="orderBasicInfo_"]')?.innerText || '';
    // Meituan currently emits both 18 and 19 digit order numbers. Treating all
    // order numbers as 19 digits drops otherwise valid cards and makes the
    // extractor mistake the missing row for another page.
    const orderNo = (basic.match(/订单编号：?(\\d{18,20})/) || [])[1] || null;
    const dailyOrderNoText = (basic.match(/#\\s*(\\d+)/) || text.match(/^#\\s*(\\d+)/m) || [])[1] || '';
    const dailyOrderNo = dailyOrderNoText ? Number(dailyOrderNoText) : null;
    const md = basic.match(/(\\d{2})-(\\d{2})\\s+(\\d{2}:\\d{2})\\s+下单/) || [];
    const pairValue = (label) => {
      const labelNode = [...card.querySelectorAll('.settle-left')].find((node) => node.innerText.trim() === label);
      return money(labelNode?.nextElementSibling?.innerText || '');
    };
    const items = [...card.querySelectorAll('[class^="goods-item-row_"]')].map((row, index) => {
      const rowText = row.innerText;
      const amounts = [...row.querySelectorAll('.goods-num')].map((node) => node.innerText.trim());
      const quantityText = amounts.find((value) => /^x/i.test(value)) || 'x0';
      return {
        lineIndex: index + 1,
        goodsCode: (rowText.match(/货号\\s*([^\\s]+)/) || [])[1] || '',
        name: row.querySelector('.goods-name')?.innerText.trim() || '',
        spec: (rowText.match(/规格\\s*([^\\n]+)/) || [])[1] || '',
        quantity: Number(quantityText.replace(/^x/i, '')),
        unitPrice: money(amounts[0]),
        subtotal: money(amounts[amounts.length - 1]),
      };
    });
    const rawStatus = card.querySelector('[class^="baseInfoRight_"]')?.innerText.trim() || '';
    const expectedItemCount = Number(((card.querySelector('[class^="productCardTitle_"]')?.innerText || text).match(/(\\d+)件商品/) || [])[1] || 0);
    return {
      orderNo,
      dailyOrderNo,
      orderDate: md.length ? selectedYear + '-' + md[1] + '-' + md[2] : null,
      orderTime: md[3] || null,
      status: normalizeStatus(rawStatus),
      rawStatus,
      productSales: pairValue('商品小计'),
      customerPay: money((text.match(/顾客实际支付[￥¥]\\s*\\d+(?:\\.\\d+)?/) || [])[0]),
      expectedIncome: money(card.querySelector('.expect-amount')?.innerText || ''),
      platformFees: Math.abs(pairValue('商家费用合计') || 0),
      deliveryFee: money((text.match(/运费实付[￥¥]\\s*\\d+(?:\\.\\d+)?/) || [])[0]),
      expectedItemCount,
      items,
      refunds: extractRefunds(orderModel),
    };
  }).filter((order) => order.orderNo);
  const totalText = root.body.innerText.match(/共\\s*(\\d+)\\s*单/);
  const next = root.querySelector('a[aria-label="Next"]');
  const activePage = Number(root.querySelector('li.active > a')?.textContent.trim() || 1);
  return {
    total: totalText ? Number(totalText[1]) : orders.length,
    orders,
    activePage,
    nextDisabled: !next
      || next.classList.contains('disabled')
      || next.closest('li')?.classList.contains('disabled')
      || next.getAttribute('aria-disabled') === 'true',
  };
})()`;

async function main() {
  const args = parseArgs(process.argv);
  assertDate(args.start, '--start');
  assertDate(args.end, '--end');
  const health = await proxy('/health');
  if (!health) throw new Error('CDP proxy is unavailable');

  const opened = await proxy(`/new?url=${encodeURIComponent(ORDER_URL)}`);
  const targetId = opened.targetId;
  if (!targetId) throw new Error('Could not create a background Chrome tab');

  try {
    await waitForOrderPageReady(targetId);
    await new Promise((resolve) => setTimeout(resolve, 1800));
    await selectDateRange(targetId, args.start, args.end);
    const selectedDates = await evaluate(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      return root ? [...root.querySelectorAll('#range input')].map((node) => node.value) : [];
    })()`);
    if (selectedDates[0] !== args.start || selectedDates[1] !== args.end) {
      throw new Error(`Calendar selection failed: ${JSON.stringify(selectedDates)}`);
    }
    await waitFor(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      const inputs = root ? [...root.querySelectorAll('#range input')].map((node) => node.value) : [];
      const hasOrders = Boolean(root?.querySelector('[class^="order-card_"]'));
      const isEmpty = /共\\s*0\\s*单/.test(root?.body?.innerText || '');
      return inputs[0] === '${args.start}' && inputs[1] === '${args.end}' && (hasOrders || isEmpty);
    })()`);

    const allOrders = [];
    const seenPageSignatures = new Set();
    let expectedTotal = null;
    for (let page = 1; page <= 50; page += 1) {
      await expandOrderCards(targetId);
      const data = await evaluate(targetId, extractPageExpression);
      if (expectedTotal === null) expectedTotal = data.total;
      if (!data.orders.length) break;
      const pageSignature = data.orders.map((order) => order.orderNo).join('|');
      if (seenPageSignatures.has(pageSignature)) break;
      seenPageSignatures.add(pageSignature);
      allOrders.push(...data.orders);
      if (data.nextDisabled || allOrders.length >= expectedTotal) break;
      await evaluate(targetId, `(() => {
        const root = document.querySelector('#hashframe').contentDocument;
        const next = root.querySelector('a[aria-label="Next"]');
        if (!next) throw new Error('Next page button not found');
        next.scrollIntoView({ block: 'center' });
        next.dispatchEvent(new MouseEvent('click', {
          bubbles: true,
          cancelable: true,
          view: root.defaultView,
        }));
        return true;
      })()`);
      await waitFor(targetId, `(() => {
        const root = document.querySelector('#hashframe')?.contentDocument;
        if (!root) return false;
        const ids = [...root.querySelectorAll('[class^="orderBasicInfo_"]')]
          .map((node) => (node.innerText.match(/\\d{18,20}/) || [])[0] || '')
          .filter(Boolean);
        const signature = ids.join('|');
        const activePage = Number(root.querySelector('li.active > a')?.textContent.trim() || 1);
        return ids.length > 0
          && (signature !== ${JSON.stringify(pageSignature)} || activePage !== ${Number(data.activePage || page)});
      })()`);
    }

    const deduped = [...new Map(allOrders.map((order) => [order.orderNo, order])).values()];
    if (expectedTotal !== null && deduped.length !== expectedTotal) {
      throw new Error(`Expected ${expectedTotal} orders but extracted ${deduped.length}`);
    }
    process.stdout.write(JSON.stringify({
      startDate: args.start,
      endDate: args.end,
      extractedAt: new Date().toISOString(),
      total: deduped.length,
      orders: deduped,
    }));
  } finally {
    await proxy(`/close?target=${encodeURIComponent(targetId)}`).catch(() => {});
  }
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error.message}\n`);
  process.exit(1);
});
