import fs from 'node:fs';
import path from 'node:path';
import { openPetstarMerchantTab, releasePetstarMerchantTab } from './meituan_browser_context.mjs';

const PROXY = 'http://localhost:3456';
const PRODUCT_URL = 'https://shangoue.meituan.com/#/reuse/sc/product/page/list/product/list';
const PAGE_TIMEOUT_MS = Number(process.env.MEITUAN_PAGE_TIMEOUT_MS || 90000);
const EXPORT_TIMEOUT_MS = Number(process.env.MEITUAN_EXPORT_TIMEOUT_MS || 300000);
const POLL_INTERVAL_MS = Number(process.env.MEITUAN_EXPORT_POLL_MS || 2500);
const REFRESH_SETTLE_MS = Number(process.env.MEITUAN_REFRESH_SETTLE_MS || 2000);
const PROXY_TIMEOUT_MS = Number(process.env.MEITUAN_PROXY_TIMEOUT_MS || 30000);
const MAX_ATTEMPTS = Math.max(1, Number(process.env.MEITUAN_EXPORT_ATTEMPTS || 3));
const DIAGNOSTIC_DIR = process.env.MEITUAN_DIAGNOSTIC_DIR
  || path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(?:[A-Za-z]:)/, (value) => value.slice(1))), '..', 'agent', 'state', 'diagnostics');

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
    method: 'POST',
    headers: { 'content-type': 'text/plain; charset=utf-8' },
    body: expression,
  });
  return result.value;
}

async function waitFor(targetId, expression, timeoutMs, description) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (await evaluate(targetId, expression)) return;
    await sleep(800);
  }
  throw new Error(`Timed out waiting for ${description} after ${Math.round(timeoutMs / 1000)} seconds`);
}

function exportModalExpression(innerExpression) {
  return `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    if (!root) return null;
    const visible = (node) => {
      if (!node || node.offsetParent === null) return false;
      const rect = node.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0;
    };
    const modals = [...root.querySelectorAll('.roo-modal')].filter(visible);
    const modal = modals.find((node) => {
      const text = node.innerText || '';
      return /新增下载文件|下载文件|操作时间|生成状态|任务状态/.test(text);
    });
    if (!modal) return null;
    return (${innerExpression});
  })()`;
}

async function captureDiagnostics(targetId, stage, attempt, error) {
  try {
    fs.mkdirSync(DIAGNOSTIC_DIR, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const base = `inventory-${stamp}-a${attempt}-${stage.replace(/[^a-z0-9_-]/gi, '-')}`;
    const screenshotPath = path.join(DIAGNOSTIC_DIR, `${base}.png`);
    const statePath = path.join(DIAGNOSTIC_DIR, `${base}.json`);
    const state = await evaluate(targetId, `(() => {
      const frame = document.querySelector('#hashframe');
      const root = frame?.contentDocument;
      const visible = (node) => node && node.offsetParent !== null;
      const modals = [...(root?.querySelectorAll('.roo-modal') || [])].filter(visible);
      return {
        pageTitle: document.title,
        pageUrl: location.origin + location.pathname + location.hash,
        iframePresent: Boolean(frame),
        iframeLoaded: Boolean(root),
        productLinkPresent: [...(root?.querySelectorAll('a') || [])].some((node) => node.innerText.trim() === '下载全部商品'),
        visibleModalCount: modals.length,
        modalSummaries: modals.map((node) => (node.innerText || '').trim().slice(0, 300)),
        visibleButtons: [...(root?.querySelectorAll('button') || [])].filter(visible).map((node) => node.innerText.trim()).filter(Boolean).slice(0, 40),
        loginHint: /登录|重新登录|扫码登录/.test(root?.body?.innerText || document.body?.innerText || ''),
      };
    })()`).catch(() => null);
    await proxy(`/screenshot?target=${encodeURIComponent(targetId)}&file=${encodeURIComponent(screenshotPath)}`).catch(() => null);
    fs.writeFileSync(statePath, JSON.stringify({
      capturedAt: new Date().toISOString(), stage, attempt,
      error: error?.message || String(error), state,
    }, null, 2), 'utf8');
    return { screenshotPath, statePath, state };
  } catch {
    return null;
  }
}

async function invokeReactFrameClick(targetId, selector, text, modalOnly = false) {
  const result = await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    if (!root) return { called: false, reason: 'iframe-missing' };
    const visible = (node) => {
      if (!node || node.offsetParent === null) return false;
      const rect = node.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0;
    };
    const scope = ${modalOnly ? "[...root.querySelectorAll('.roo-modal')].filter(visible).find((node) => /新增下载文件|下载文件|操作时间|生成状态|任务状态/.test(node.innerText || ''))" : 'root'};
    if (!scope) return { called: false, reason: 'scope-missing' };
    const element = [...(scope.querySelectorAll(${JSON.stringify(selector)}) || [])].find((node) => {
      const rect = node.getBoundingClientRect();
      return node.offsetParent !== null && node.innerText.trim() === ${JSON.stringify(text)}
        && rect.width > 0 && rect.height > 0;
    });
    if (!element) return { called: false, reason: 'element-missing' };
    const propKey = Object.keys(element).find((key) => key.startsWith('__reactProps'));
    const handler = propKey && element[propKey]?.onClick;
    if (typeof handler !== 'function') return { called: false, reason: 'react-handler-missing' };
    handler({
      currentTarget: element,
      target: element,
      preventDefault() {},
      stopPropagation() {},
      nativeEvent: { isTrusted: true },
      isDefaultPrevented: () => false,
      isPropagationStopped: () => false,
    });
    return { called: true };
  })()`);
  if (!result?.called) throw new Error(`Could not invoke ${text}: ${result?.reason || 'unknown'}`);
}

async function trustedClickFrameButton(targetId, buttonText) {
  const markerId = `petstar-click-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const prepared = await evaluate(targetId, `(() => {
    const frame = document.querySelector('#hashframe');
    const root = frame?.contentDocument;
    const modal = [...(root?.querySelectorAll('.roo-modal') || [])]
      .filter((node) => node.offsetParent !== null)
      .find((node) => /新增下载文件|下载文件|操作时间|生成状态|任务状态/.test(node.innerText || ''));
    const button = [...(modal?.querySelectorAll('button') || [])]
      .find((node) => {
        const rect = node.getBoundingClientRect();
        return node.offsetParent !== null && !node.disabled
          && node.innerText.trim() === ${JSON.stringify(buttonText)}
          && rect.width > 0 && rect.height > 0 && rect.right > 0 && rect.bottom > 0
          && rect.left < root.defaultView.innerWidth && rect.top < root.defaultView.innerHeight;
      });
    if (!frame || !root || !modal || !button) return { ready: false };
    const frameRect = frame.getBoundingClientRect();
    const buttonRect = button.getBoundingClientRect();
    const x = frameRect.left + buttonRect.left + buttonRect.width / 2;
    const y = frameRect.top + buttonRect.top + buttonRect.height / 2;
    const marker = document.createElement('div');
    marker.id = ${JSON.stringify(markerId)};
    Object.assign(marker.style, {
      position: 'fixed', left: (x - 2) + 'px', top: (y - 2) + 'px',
      width: '4px', height: '4px', pointerEvents: 'auto', background: 'transparent', zIndex: '2147483647',
    });
    marker.addEventListener('click', () => button.click(), { once: true });
    document.body.appendChild(marker);
    return { ready: true };
  })()`);
  if (!prepared?.ready) throw new Error(`Button is not ready for trusted click: ${buttonText}`);
  try {
    await proxy(`/clickAt?target=${encodeURIComponent(targetId)}`, {
      method: 'POST',
      headers: { 'content-type': 'text/plain; charset=utf-8' },
      body: `#${markerId}`,
    });
  } finally {
    await evaluate(targetId, `document.getElementById(${JSON.stringify(markerId)})?.remove()`).catch(() => {});
  }
}

async function trustedClickFrameLink(targetId, linkText) {
  const markerId = `petstar-link-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const prepared = await evaluate(targetId, `(() => {
    const frame = document.querySelector('#hashframe');
    const root = frame?.contentDocument;
    const link = [...(root?.querySelectorAll('a') || [])].find((node) => {
      const rect = node.getBoundingClientRect();
      return node.offsetParent !== null && node.innerText.trim() === ${JSON.stringify(linkText)}
        && rect.width > 0 && rect.height > 0 && rect.right > 0 && rect.bottom > 0
        && rect.left < root.defaultView.innerWidth && rect.top < root.defaultView.innerHeight;
    });
    if (!frame || !root || !link) return { ready: false };
    link.scrollIntoView({ block: 'center' });
    const frameRect = frame.getBoundingClientRect();
    const linkRect = link.getBoundingClientRect();
    const x = frameRect.left + linkRect.left + linkRect.width / 2;
    const y = frameRect.top + linkRect.top + linkRect.height / 2;
    const marker = document.createElement('div');
    marker.id = ${JSON.stringify(markerId)};
    Object.assign(marker.style, {
      position: 'fixed', left: (x - 2) + 'px', top: (y - 2) + 'px',
      width: '4px', height: '4px', pointerEvents: 'auto', background: 'transparent', zIndex: '2147483647',
    });
    marker.addEventListener('click', () => link.click(), { once: true });
    document.body.appendChild(marker);
    return { ready: true };
  })()`);
  if (!prepared?.ready) throw new Error(`Link is not ready for trusted click: ${linkText}`);
  try {
    await proxy(`/clickAt?target=${encodeURIComponent(targetId)}`, {
      method: 'POST',
      headers: { 'content-type': 'text/plain; charset=utf-8' },
      body: `#${markerId}`,
    });
  } finally {
    await evaluate(targetId, `document.getElementById(${JSON.stringify(markerId)})?.remove()`).catch(() => {});
  }
}

async function getExportTasks(targetId) {
  return await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    const modal = [...(root?.querySelectorAll('.roo-modal') || [])]
      .filter((node) => node.offsetParent !== null)
      .find((node) => /新增下载文件|下载文件|操作时间|生成状态|任务状态/.test(node.innerText || ''));
    return [...(modal?.querySelectorAll('tbody tr') || [])].map((row) => {
      const cells = [...row.querySelectorAll('td')].map((cell) => cell.innerText.trim());
      const operation = cells[0] || '';
      const operationTime = cells[1] || '';
      const status = cells[2] || '';
      return {
        key: operation + '|' + operationTime,
        operation,
        operationTime,
        status,
        href: row.querySelector('a')?.href || '',
        taskText: row.innerText.trim(),
      };
    }).filter((task) => /^[0-9]{4}[.][0-9]{2}[.][0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}$/.test(task.operationTime));
  })()`);
}

function parseMeituanTaskTime(value) {
  const match = /^(\d{4})\.(\d{2})\.(\d{2})\s+(\d{2}):(\d{2}):(\d{2})$/.exec(value || '');
  if (!match) return Number.NaN;
  return new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4]),
    Number(match[5]),
    Number(match[6]),
  ).getTime();
}

async function refreshExportTasks(targetId) {
  const refreshed = await evaluate(targetId, `(() => {
    const root = document.querySelector('#hashframe')?.contentDocument;
    if (!root) return false;
    const modal = [...root.querySelectorAll('.roo-modal')]
      .filter((node) => node.offsetParent !== null)
      .find((node) => /新增下载文件|下载文件|操作时间|生成状态|任务状态/.test(node.innerText || ''));
    const refresh = [...(modal?.querySelectorAll('button') || [])].find((button) => button.innerText.trim() === '刷新');
    if (!refresh) return false;
    refresh.click();
    return true;
  })()`);
  // The task-list API can load later than the dialog. If Refresh is not ready
  // yet, keep polling the table instead of failing the whole export.
  await sleep(refreshed ? REFRESH_SETTLE_MS : POLL_INTERVAL_MS);
  return await getExportTasks(targetId);
}

async function hasExportDialog(targetId) {
  return Boolean(await evaluate(targetId, exportModalExpression('true')));
}

async function ensureExportDialog(targetId) {
  if (await hasExportDialog(targetId)) return;
  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      await invokeReactFrameClick(targetId, 'a', '下载全部商品');
    } catch (error) {
      lastError = error;
      try {
        await trustedClickFrameLink(targetId, '下载全部商品');
      } catch (fallbackError) {
        lastError = fallbackError;
      }
    }
    try {
      await waitFor(targetId, exportModalExpression('true'), 20000, 'the inventory export dialog to open');
      return;
    } catch (error) {
      lastError = error;
      await sleep(1000 * attempt);
    }
  }
  throw lastError || new Error('Could not open the inventory export dialog');
}

function newestTaskSince(tasks, sinceMs, predicate = () => true) {
  return [...tasks]
    .filter((task) => {
      const taskTime = parseMeituanTaskTime(task.operationTime);
      return Number.isFinite(taskTime) && taskTime >= sinceMs && predicate(task);
    })
    .sort((a, b) => parseMeituanTaskTime(b.operationTime) - parseMeituanTaskTime(a.operationTime))[0] || null;
}

function successfulTaskSince(tasks, sinceMs) {
  return newestTaskSince(tasks, sinceMs, (task) => /生成成功/.test(task.status || task.taskText || '') && task.href);
}

async function clickCreateExport(targetId) {
  await waitFor(targetId, exportModalExpression(`(() => {
    const button = [...modal.querySelectorAll('button')].find((node) => node.offsetParent !== null
      && !node.disabled && node.innerText.trim() === '新增下载文件');
    return Boolean(button);
  })()`), 30000, 'the create-export button');
  try {
    await invokeReactFrameClick(targetId, 'button', '新增下载文件', true);
    return 'react';
  } catch (error) {
    await trustedClickFrameButton(targetId, '新增下载文件');
    return 'trusted';
  }
}

async function ensureRequiredExportFields(targetId) {
  const required = [
    'spu_name', 'source_food_code', 'sku_id', 'locator_code',
    'spec_name', 'stock', 'month_sale', 'sell_status', 'product_date',
  ];
  // The export modal becomes visible before Meituan finishes loading its
  // checkbox definitions. Seeing the modal (or the create button) is not
  // enough: wait until every required field actually exists in the DOM.
  await waitFor(targetId, exportModalExpression(`(() => {
    const required = ${JSON.stringify(required)};
    const inputs = [...modal.querySelectorAll('input[type="checkbox"]')];
    return !/正在加载/.test(modal.innerText || '')
      && required.every((value) => inputs.some((input) => input.value === value));
  })()`), 60000, 'the required inventory export fields to finish loading');
  const result = await evaluate(targetId, exportModalExpression(`(() => {
    const required = ${JSON.stringify(required)};
    const inputs = [...modal.querySelectorAll('input[type="checkbox"]')];
    const missing = required.filter((value) => !inputs.some((input) => input.value === value));
    if (missing.length) return { ok: false, missing };
    for (const value of required) {
      const input = inputs.find((candidate) => candidate.value === value);
      if (input && !input.checked) input.click();
    }
    return {
      ok: true,
      selected: required.filter((value) => inputs.find((candidate) => candidate.value === value)?.checked),
    };
  })()`));
  if (!result?.ok) throw new Error(`Required export fields are unavailable: ${(result?.missing || required).join(', ')}`);
  await sleep(500);
  const selected = await evaluate(targetId, exportModalExpression(`(() => {
    const required = ${JSON.stringify(required)};
    const inputs = [...modal.querySelectorAll('input[type="checkbox"]')];
    return required.filter((value) => inputs.find((candidate) => candidate.value === value)?.checked);
  })()`));
  const stillMissing = required.filter((value) => !(selected || []).includes(value));
  if (stillMissing.length) throw new Error(`Could not select required export fields: ${stillMissing.join(', ')}`);
  return selected;
}

async function runAttempt(attempt, globalStartedAt) {
  const opened = await openPetstarMerchantTab(PRODUCT_URL);
  const targetId = opened.targetId;
  if (!targetId) throw new Error('Could not create a background Chrome tab');
  let stage = 'page-open';
  try {
    await waitFor(targetId, `(() => {
      const root = document.querySelector('#hashframe')?.contentDocument;
      return [...(root?.querySelectorAll('a') || [])].some((node) => node.innerText.trim() === '下载全部商品');
    })()`, PAGE_TIMEOUT_MS, 'the Meituan product page');
    await sleep(1200);

    stage = 'dialog-open';
    await ensureExportDialog(targetId);

    // Export field selections are persisted by Meituan and may be changed by
    // manual downloads. Always assert the fields this system requires before
    // creating a task; never reuse an unvalidated recent task.
    stage = 'select-fields';
    const selectedFields = await ensureRequiredExportFields(targetId);
    let tasks = await getExportTasks(targetId);

    const beforeTaskKeys = new Set(tasks.map((task) => task.key));
    const exportStartedAt = Date.now();
    stage = 'create-export';
    const clickMode = await clickCreateExport(targetId);

    stage = 'track-export';
    let trackedTaskKey = '';
    let trackedTask = null;
    while (Date.now() - exportStartedAt < EXPORT_TIMEOUT_MS) {
      await sleep(POLL_INTERVAL_MS);
      if (!(await hasExportDialog(targetId))) await ensureExportDialog(targetId);
      tasks = await refreshExportTasks(targetId);

      const anySuccessful = successfulTaskSince(tasks, exportStartedAt - 60000);
      if (anySuccessful && (!beforeTaskKeys.has(anySuccessful.key) || parseMeituanTaskTime(anySuccessful.operationTime) >= exportStartedAt - 5000)) {
        return {
          downloadUrl: anySuccessful.href,
          taskKey: anySuccessful.key,
          taskText: anySuccessful.taskText,
          waitedSeconds: Math.round((Date.now() - exportStartedAt) / 1000),
          attempt,
          recovered: false,
          clickMode,
          selectedFields,
        };
      }

      if (!trackedTaskKey) {
        const newTask = newestTaskSince(tasks, exportStartedAt - 60000, (task) => task.key && !beforeTaskKeys.has(task.key));
        if (newTask) trackedTaskKey = newTask.key;
      }
      if (!trackedTaskKey) continue;
      trackedTask = tasks.find((task) => task.key === trackedTaskKey) || trackedTask;
      if (!trackedTask) continue;
      if (/失败/.test(trackedTask.status || trackedTask.taskText || '')) {
        throw new Error(`Meituan inventory export task failed: ${trackedTask.taskText}`);
      }
    }

    if (!trackedTaskKey) {
      throw new Error(`The newly created Meituan inventory export task did not appear within ${Math.round(EXPORT_TIMEOUT_MS / 1000)} seconds`);
    }
    const lastState = trackedTask?.taskText || trackedTaskKey;
    throw new Error(`Meituan inventory export task did not finish within ${Math.round(EXPORT_TIMEOUT_MS / 1000)} seconds; last state: ${lastState}`);
  } catch (error) {
    const diagnostics = await captureDiagnostics(targetId, stage, attempt, error);
    error.stage = stage;
    error.diagnostics = diagnostics;
    throw error;
  } finally {
    await releasePetstarMerchantTab(opened);
  }
}

async function main() {
  await proxy('/health');
  const globalStartedAt = Date.now();
  const failures = [];
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt += 1) {
    try {
      const result = await runAttempt(attempt, globalStartedAt);
      process.stdout.write(JSON.stringify({ ...result, totalAttempts: attempt }));
      return;
    } catch (error) {
      failures.push({
        attempt,
        stage: error.stage || 'unknown',
        message: error.message || String(error),
        diagnostics: error.diagnostics?.statePath || null,
      });
      if (attempt < MAX_ATTEMPTS) await sleep(Math.min(10000, attempt * 2500));
    }
  }
  const last = failures[failures.length - 1];
  const failure = new Error(last?.message || 'Meituan inventory export failed');
  failure.code = 'INVENTORY_EXPORT_RETRIES_EXHAUSTED';
  failure.failures = failures;
  throw failure;
}

main().catch((error) => {
  process.stderr.write(`${JSON.stringify({
    ok: false,
    code: error.code || 'INVENTORY_EXPORT_FAILED',
    message: error.message || String(error),
    failures: error.failures || null,
  })}\n`);
  process.exit(1);
});
