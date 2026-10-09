import fs from 'node:fs';
import path from 'node:path';
import { openPetstarMerchantTab, releasePetstarMerchantTab } from './meituan_browser_context.mjs';

const PROXY = 'http://localhost:3456';
const MERCHANT_URL = 'https://shangoue.meituan.com/';
const PROXY_TIMEOUT_MS = Number(process.env.MEITUAN_PROXY_TIMEOUT_MS || 60000);
const RESOURCE_TIMEOUT_MS = Number(process.env.MEITUAN_PROMOTION_RESOURCE_TIMEOUT_MS || 60000);
const DEBUG = process.env.MEITUAN_PROMOTION_DEBUG === '1';
const debug = (message) => { if (DEBUG) process.stderr.write(`[promotion] ${message}\n`); };

function parseArgs(argv) {
  const result = {};
  for (let i = 2; i < argv.length; i += 1) {
    if (argv[i] === '--date') result.date = argv[++i];
  }
  return result;
}

function getShanghaiDate(offsetDays = 0) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(new Date()).filter((part) => part.type !== 'literal').map((part) => [part.type, part.value]));
  return new Date(Date.UTC(Number(parts.year), Number(parts.month) - 1, Number(parts.day) + offsetDays))
    .toISOString().slice(0, 10);
}

function getShanghaiHour() {
  return Number(new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Shanghai', hour: '2-digit', hourCycle: 'h23',
  }).format(new Date()));
}

function assertDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) throw new Error('--date must use YYYY-MM-DD');
  const today = getShanghaiDate();
  const yesterday = getShanghaiDate(-1);
  const isRecovery = value === yesterday;
  if (value !== today && !isRecovery) {
    throw new Error(`Promotion dashboard can only prove today's spend or yesterday's settled spend (${today}), requested ${value}`);
  }
  return { today, isRecovery };
}

async function proxy(path, options = {}) {
  const response = await fetch(`${PROXY}${path}`, {
    ...options,
    signal: options.signal || AbortSignal.timeout(PROXY_TIMEOUT_MS),
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`CDP proxy request failed: ${response.status} ${body}`);
  const data = JSON.parse(body);
  if (data.error) throw new Error(data.error);
  return data;
}

async function evaluate(targetId, expression) {
  const result = await proxy(`/eval?target=${encodeURIComponent(targetId)}`, {
    method: 'POST', headers: { 'content-type': 'text/plain; charset=utf-8' }, body: expression,
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
    } catch (error) { lastError = error; }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  throw new Error(`Timed out waiting for promotion page${lastError ? `: ${lastError.message}` : ''}`);
}

async function waitForPromotionRoute(targetId, timeoutMs = 60000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const targets = await proxy('/targets');
    const target = targets.find((item) => item.targetId === targetId);
    if (target?.url) {
      let route = null;
      try { route = decodeURIComponent(new URL(target.url).hash.slice(1)); }
      catch { route = null; }
      if (/^https:\/\/waimaieapp\.meituan\.com\/ad\/v1\/(?:pc|rpc)(?:[/?#]|$)/.test(route || '') &&
          /#\/(?:index|account)(?:[/?#]|$)/.test(route || '')) return route;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  throw new Error('Timed out waiting for promotion route');
}

function isAuthenticatedPromotionPortalUrl(value) {
  return /^https:\/\/shangoue\.meituan\.com\/#https:\/\/waimaieapp\.meituan\.com\/ad\/v1\/(?:pc|rpc)(?:[/?#]|$)/.test(String(value || ''));
}

function withPromotionPage(portalUrl, page) {
  if (!isAuthenticatedPromotionPortalUrl(portalUrl)) {
    throw new Error('Authenticated promotion portal URL is unavailable');
  }
  const innerHashIndex = String(portalUrl).lastIndexOf('#/');
  if (innerHashIndex < 0) throw new Error('Could not locate promotion portal inner route');
  return `${String(portalUrl).slice(0, innerHashIndex)}${page}`;
}

async function openAuthenticatedPromotionPortal() {
  const opened = await openPetstarMerchantTab(MERCHANT_URL);
  const targetId = opened.targetId;
  if (!targetId) throw new Error('Could not create Meituan merchant tab');
  try {
    await waitFor(targetId, `!!document.querySelector('li[data-key="Ori.WMAdvert"]')`, RESOURCE_TIMEOUT_MS);
    const clicked = await evaluate(targetId, `(() => {
      const entry = document.querySelector('li[data-key="Ori.WMAdvert"]');
      if (!entry) return false;
      entry.click();
      return true;
    })()`);
    if (!clicked) throw new Error('Could not open Meituan promotion portal');
    await waitForPromotionRoute(targetId, RESOURCE_TIMEOUT_MS);
    const targets = await proxy('/targets');
    const promoted = targets.find((item) => item.targetId === targetId);
    if (!promoted || !isAuthenticatedPromotionPortalUrl(promoted.url)) {
      throw new Error('Meituan promotion portal did not provide an authenticated route');
    }
    return { ...opened, url: promoted.url };
  } catch (error) {
    await releasePetstarMerchantTab(opened);
    throw error;
  }
}

function getChromeWebSocketUrl() {
  const portFile = path.join(process.env.LOCALAPPDATA || '', 'Google', 'Chrome', 'User Data', 'DevToolsActivePort');
  const [port, webSocketPath] = fs.readFileSync(portFile, 'utf8').trim().split(/\r?\n/);
  if (!/^\d+$/.test(port || '') || !webSocketPath?.startsWith('/')) {
    throw new Error('Chrome DevToolsActivePort is invalid');
  }
  return `ws://127.0.0.1:${port}${webSocketPath}`;
}

class CdpBrowser {
  constructor(url) {
    this.url = url;
    this.nextId = 0;
    this.pending = new Map();
    this.listeners = new Set();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener('open', resolve, { once: true });
      this.socket.addEventListener('error', () => reject(new Error('Could not connect to Chrome DevTools')), { once: true });
    });
    this.socket.addEventListener('message', (event) => {
      const message = JSON.parse(String(event.data));
      if (message.id && this.pending.has(message.id)) {
        const pending = this.pending.get(message.id);
        clearTimeout(pending.timer);
        this.pending.delete(message.id);
        if (message.error) pending.reject(new Error(message.error.message || 'CDP command failed'));
        else pending.resolve(message.result || {});
        return;
      }
      for (const listener of this.listeners) listener(message);
    });
  }

  send(method, params = {}, sessionId = null, timeoutMs = 30000) {
    return new Promise((resolve, reject) => {
      const id = ++this.nextId;
      const command = { id, method, params };
      if (sessionId) command.sessionId = sessionId;
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`CDP command timed out: ${method}`));
      }, timeoutMs);
      this.pending.set(id, { resolve, reject, timer });
      this.socket.send(JSON.stringify(command));
    });
  }

  onEvent(listener) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  close() {
    if (this.socket?.readyState === WebSocket.OPEN) this.socket.close();
  }
}

async function captureJsonResource(targetId, pageUrl, pathFragment, timeoutMs = RESOURCE_TIMEOUT_MS) {
  debug(`capture start ${pathFragment}`);
  const browser = new CdpBrowser(getChromeWebSocketUrl());
  await browser.connect();
  let sessionId = null;
  let removeListener = null;
  try {
    const attached = await browser.send('Target.attachToTarget', { targetId, flatten: true });
    sessionId = attached.sessionId;
    if (!sessionId) throw new Error('Could not attach to Petstar merchant tab');
    await browser.send('Network.enable', {}, sessionId);
    await browser.send('Page.enable', {}, sessionId);
    let matched = null;
    let resolveCompleted;
    const completed = new Promise((resolve) => { resolveCompleted = resolve; });
    removeListener = browser.onEvent((message) => {
      if (message.sessionId !== sessionId) return;
      if (message.method === 'Network.responseReceived' && message.params?.response?.url?.includes(pathFragment)) {
        matched = { requestId: message.params.requestId, status: message.params.response.status };
      }
      if (message.method === 'Network.loadingFinished' && matched?.requestId === message.params?.requestId) {
        resolveCompleted();
      }
    });
    await browser.send('Page.navigate', { url: pageUrl }, sessionId, timeoutMs);
    await Promise.race([
      completed,
      new Promise((_, reject) => setTimeout(() => reject(new Error(`Timed out waiting for network resource: ${pathFragment}`)), timeoutMs)),
    ]);
    if (!matched) throw new Error(`Network resource not found: ${pathFragment}`);
    if (matched.status < 200 || matched.status >= 300) {
      throw new Error(`Network resource failed (${matched.status}): ${pathFragment}`);
    }
    const response = await browser.send('Network.getResponseBody', { requestId: matched.requestId }, sessionId);
    const body = response.base64Encoded
      ? Buffer.from(response.body || '', 'base64').toString('utf8')
      : String(response.body || '');
    debug(`capture complete ${pathFragment}`);
    return JSON.parse(body);
  } catch (error) {
    if (error instanceof SyntaxError) throw new Error(`Promotion API did not return JSON: ${pathFragment}`);
    throw error;
  } finally {
    removeListener?.();
    if (sessionId) await browser.send('Target.detachFromTarget', { sessionId }).catch(() => {});
    browser.close();
  }
}

function walk(value, visitor, path = '') {
  if (Array.isArray(value)) {
    value.forEach((item, index) => walk(item, visitor, `${path}[${index}]`));
    return;
  }
  if (!value || typeof value !== 'object') return;
  visitor(value, path);
  for (const [key, child] of Object.entries(value)) walk(child, visitor, path ? `${path}.${key}` : key);
}

function findHomepageTotal(response) {
  const candidates = [];
  walk(response, (object, path) => {
    for (const [key, value] of Object.entries(object)) {
      if (key === 'total' && Number.isFinite(Number(value))) candidates.push({ path: path ? `${path}.${key}` : key, value: Number(value) });
    }
  });
  const preferred = candidates.find((item) => /(^|\.)data\.total$/.test(item.path))
    || candidates.find((item) => /account.*total/i.test(item.path))
    || (candidates.length === 1 ? candidates[0] : null);
  if (!preferred || preferred.value < 0) {
    throw new Error(`Could not uniquely locate homepage promotion total; candidates=${candidates.map((item) => item.path).join(',')}`);
  }
  return preferred.value / 100;
}

function findHistorySpend(response, date, accountTotal) {
  const matches = [];
  walk(response, (object) => {
    const reason = String(object.reason ?? object.bizReason ?? object.remark ?? object.title ?? '');
    const time = String(object.date ?? object.tradeDate ?? object.time ?? object.createTime ?? object.createdAt ?? object.tradeTime ?? object.occurTime ?? '');
    const amountRaw = object.amount ?? object.changeAmount ?? object.tradeAmount ?? object.money;
    const amount = Number(amountRaw);
    if (/推广消费/.test(reason) && Number.isFinite(amount) && time.startsWith(date)) {
      matches.push({ amount: Math.abs(amount), provisional: Boolean(object.unDebit), time });
    }
  });
  const provisional = matches.filter((item) => item.provisional);
  if (provisional.length === 1) return provisional[0];
  if (matches.length === 1) return matches[0];
  if (Math.abs(accountTotal) <= 0.001 && provisional.length === 0) {
    return { amount: 0, provisional: true, time: null };
  }
  throw new Error(`Expected one current promotion history row for ${date}, found ${matches.length}: ${JSON.stringify(matches)}`);
}

async function getHistoryTableSpend(targetId, date, allowMissingZero = false) {
  const result = await evaluate(targetId, `(() => {
    const date = ${JSON.stringify(date)};
    const rows = document.body.innerText.split('\\n').map((line) => line.trim()).filter(Boolean);
    const matches = rows.filter((line) => line.startsWith(date + ' 00:00:00~23:59:59') && line.includes('推广消费'));
    if (matches.length !== 1) return { error: 'expected_one_row', count: matches.length };
    const amountMatch = matches[0].match(/推广消费\\s+([+-]?\\d+(?:\\.\\d+)?)/);
    if (!amountMatch) return { error: 'amount_not_found', count: matches.length };
    return { amount: Math.abs(Number(amountMatch[1])), row_count: matches.length };
  })()`);
  if (allowMissingZero && result?.error === 'expected_one_row' && Number(result?.count) === 0) {
    return { amount: 0, time: date, provisional: true };
  }
  if (result?.error || !Number.isFinite(Number(result?.amount))) {
    throw new Error(`Could not uniquely read promotion history table for ${date}: ${result?.error || 'unknown'}`);
  }
  return { amount: Number(result.amount), time: date, provisional: false };
}

async function main() {
  const args = parseArgs(process.argv);
  const dateMode = assertDate(args.date);
  await proxy('/health');
  const promotionPortal = await openAuthenticatedPromotionPortal();
  try {
    const historyUrl = withPromotionPage(promotionPortal.url, '#/account?scrollConsume=true');
    const historyResponse = await captureJsonResource(promotionPortal.targetId, historyUrl, '/ad/v3/flow/r/history');
    const historySpend = findHistorySpend(historyResponse, args.date, 0);
    if (dateMode.isRecovery && historySpend.provisional) {
      throw new Error(`Promotion history row for ${args.date} is not settled yet`);
    }

    let accountTotal = historySpend.amount;
    if (!dateMode.isRecovery) {
      const homepageUrl = withPromotionPage(promotionPortal.url, '#/index');
      const homepage = await captureJsonResource(promotionPortal.targetId, homepageUrl, '/ad/v4/homepage/account/info');
      accountTotal = findHomepageTotal(homepage);
    }
    if (!dateMode.isRecovery && Math.abs(accountTotal - historySpend.amount) > 0.01) {
      throw new Error(`Promotion totals disagree: homepage=${accountTotal.toFixed(2)}, history=${historySpend.amount.toFixed(2)}`);
    }
    const capturedAmount = dateMode.isRecovery ? historySpend.amount : accountTotal;

    await new Promise((resolve) => process.stdout.write(JSON.stringify({
      ok: true,
      date: args.date,
      amount: Number(capturedAmount.toFixed(2)),
      provisional: !dateMode.isRecovery,
      recovery_capture: dateMode.isRecovery,
      history_un_debited: !dateMode.isRecovery,
      captured_at: new Date().toISOString(),
      evidence: dateMode.isRecovery
        ? ['promotion_history_table']
        : ['homepage_account_total', 'promotion_history_table'],
    }), resolve));
  } finally {
    await releasePetstarMerchantTab(promotionPortal);
  }
}

main().catch((error) => {
  process.exitCode = 1;
  process.stderr.write(`${error.stack || error.message}\n`);
});
