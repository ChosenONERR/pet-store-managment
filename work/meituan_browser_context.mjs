import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const PROXY = 'http://localhost:3456';
const DEFAULT_STORE_MARKER = '宠速星宠物用品店';
const DEFAULT_TIMEOUT_MS = Number(process.env.MEITUAN_CONTEXT_TIMEOUT_MS || 90000);
const MODULE_DIR = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_BINDING_PATH = path.resolve(MODULE_DIR, '..', 'agent', 'state', 'meituan_browser_binding.json');
const BINDING_PATH = process.env.MEITUAN_BROWSER_BINDING_PATH || DEFAULT_BINDING_PATH;


async function proxy(requestPath, options = {}) {
  const response = await fetch(`${PROXY}${requestPath}`, {
    ...options,
    signal: options.signal || AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
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


function loadBinding() {
  try {
    const binding = JSON.parse(fs.readFileSync(BINDING_PATH, 'utf8'));
    if (!binding?.profile_directory) return null;
    return binding;
  } catch {
    return null;
  }
}


function saveBinding(binding) {
  const directory = path.dirname(BINDING_PATH);
  fs.mkdirSync(directory, { recursive: true });
  const temporaryPath = `${BINDING_PATH}.${process.pid}.tmp`;
  fs.writeFileSync(temporaryPath, `${JSON.stringify(binding, null, 2)}\n`, 'utf8');
  fs.renameSync(temporaryPath, BINDING_PATH);
}


function findChromeExecutable() {
  const candidates = [
    process.env.CHROME_PATH,
    process.env['PROGRAMFILES(X86)'] && path.join(process.env['PROGRAMFILES(X86)'], 'Google', 'Chrome', 'Application', 'chrome.exe'),
    process.env.PROGRAMFILES && path.join(process.env.PROGRAMFILES, 'Google', 'Chrome', 'Application', 'chrome.exe'),
    process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'Google', 'Chrome', 'Application', 'chrome.exe'),
  ].filter(Boolean);
  const executable = candidates.find((candidate) => fs.existsSync(candidate));
  if (!executable) throw new Error('MEITUAN_BOUND_PROFILE_UNAVAILABLE: 找不到 Google Chrome 可执行文件');
  return executable;
}


function launchInBoundProfile(profileDirectory, url) {
  const child = spawn(findChromeExecutable(), [`--profile-directory=${profileDirectory}`, url], {
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
  });
  child.unref();
}


async function waitForTarget(predicate, timeoutMs) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const targets = await proxy('/targets');
    const target = targets.find(predicate);
    if (target) return target;
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  return null;
}


async function closeTarget(targetId) {
  if (!targetId) return;
  await proxy(`/close?target=${encodeURIComponent(targetId)}`).catch(() => {});
}


async function openProfileProbe(profileDirectory, timeoutMs) {
  const marker = `petstar-profile-probe-${process.pid}-${Date.now()}`;
  const probeUrl = `data:text/html,${marker}`;
  launchInBoundProfile(profileDirectory, probeUrl);
  const target = await waitForTarget(
    (item) => item.type === 'page' && String(item.url || '').includes(marker),
    Math.min(timeoutMs, 15000),
  );
  if (!target?.browserContextId) {
    throw new Error(`MEITUAN_BOUND_PROFILE_UNAVAILABLE: 无法打开已绑定的 Chrome 资料 ${profileDirectory}`);
  }
  return target;
}


async function resolveBoundContext(binding, targets, timeoutMs) {
  const remembered = String(binding.last_browser_context_id || '');
  if (remembered && targets.some((item) => item.type === 'page' && item.browserContextId === remembered)) {
    return { browserContextId: remembered, probeTarget: null };
  }
  const probeTarget = await openProfileProbe(binding.profile_directory, timeoutMs);
  const browserContextId = probeTarget.browserContextId;
  saveBinding({
    ...binding,
    last_browser_context_id: browserContextId,
    context_refreshed_at: new Date().toISOString(),
  });
  return { browserContextId, probeTarget };
}


async function inspectMerchantTarget(targetId, storeMarker) {
  try {
    return await evaluate(targetId, `(() => {
      let frameText = '';
      try { frameText = document.querySelector('#hashframe')?.contentDocument?.body?.innerText || ''; } catch {}
      const text = (document.body?.innerText || '') + '\\n' + frameText;
      return {
        authenticatedStore: text.includes(${JSON.stringify(storeMarker)}),
        loginPage: location.href.includes('/login') || location.href.includes('login_gw')
          || (/登录|扫码登录|重新登录/.test(text) && !document.querySelector('#hashframe')),
      };
    })()`);
  } catch {
    return { authenticatedStore: false, loginPage: false };
  }
}


async function waitForMerchantIdentity(targetId, storeMarker, timeoutMs) {
  const started = Date.now();
  let identity = null;
  while (Date.now() - started < timeoutMs) {
    identity = await inspectMerchantTarget(targetId, storeMarker);
    if (identity?.authenticatedStore || identity?.loginPage) return identity;
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  return identity || { authenticatedStore: false, loginPage: false };
}


async function useExistingBoundMerchantPage(pages, targetUrl, storeMarker, timeoutMs) {
  let sawLoginPage = false;
  for (const target of pages.filter((item) => /meituan\.com/i.test(item.url || ''))) {
    const identity = await waitForMerchantIdentity(target.targetId, storeMarker, timeoutMs);
    if (identity?.loginPage) sawLoginPage = true;
    if (!identity?.authenticatedStore) continue;
    const originalUrl = target.url;
    if (originalUrl !== targetUrl) {
      await proxy(`/navigate?target=${encodeURIComponent(target.targetId)}&url=${encodeURIComponent(targetUrl)}`);
    }
    return {
      targetId: target.targetId,
      browserContextId: target.browserContextId,
      originalUrl,
      openedByAutomation: false,
      source: 'bound-authenticated-existing-page',
    };
  }
  if (sawLoginPage) {
    throw new Error('MEITUAN_AUTH_REQUIRED: 已绑定的 Chrome 资料中，美团商家后台登录态已失效；请只在该资料中重新登录');
  }
  return null;
}


async function openInBoundContext(binding, resolved, targetUrl, storeMarker, timeoutMs) {
  let target = resolved.probeTarget;
  if (!target) {
    const opened = await proxy(`/new?url=${encodeURIComponent(targetUrl)}`);
    target = await waitForTarget(
      (item) => item.type === 'page' && item.targetId === opened.targetId,
      Math.min(timeoutMs, 15000),
    );
    if (!target || target.browserContextId !== resolved.browserContextId) {
      await closeTarget(opened?.targetId);
      throw new Error(`MEITUAN_BOUND_PROFILE_UNAVAILABLE: 无法在已绑定的 Chrome 资料 ${binding.profile_directory} 中打开店铺页`);
    }
  }
  try {
    if (target.url !== targetUrl) {
      await proxy(`/navigate?target=${encodeURIComponent(target.targetId)}&url=${encodeURIComponent(targetUrl)}`);
    }
    const identity = await waitForMerchantIdentity(target.targetId, storeMarker, timeoutMs);
    if (identity?.loginPage) {
      throw new Error('MEITUAN_AUTH_REQUIRED: 已绑定的 Chrome 资料中，美团商家后台登录态已失效；请只在该资料中重新登录');
    }
    if (!identity?.authenticatedStore) {
      throw new Error(`MEITUAN_STORE_IDENTITY_UNVERIFIED: 已绑定 Chrome 资料未能验证店铺“${storeMarker}”`);
    }
    return {
      targetId: target.targetId,
      browserContextId: target.browserContextId,
      originalUrl: '',
      openedByAutomation: true,
      source: 'bound-profile-launched-page',
    };
  } catch (error) {
    await closeTarget(target.targetId);
    throw error;
  }
}


export async function openPetstarMerchantTab(targetUrl, options = {}) {
  const binding = loadBinding();
  const storeMarker = options.storeMarker || binding?.store_marker || process.env.MEITUAN_EXPECTED_STORE || DEFAULT_STORE_MARKER;
  const timeoutMs = Number(options.timeoutMs || DEFAULT_TIMEOUT_MS);
  const targets = await proxy('/targets');
  const pages = targets.filter((item) => item.type === 'page' && item.browserContextId);

  if (binding) {
    const resolved = await resolveBoundContext(binding, pages, timeoutMs);
    const boundPages = (await proxy('/targets')).filter(
      (item) => item.type === 'page' && item.browserContextId === resolved.browserContextId,
    );
    try {
      const existing = await useExistingBoundMerchantPage(boundPages, targetUrl, storeMarker, timeoutMs);
      if (existing) {
        await closeTarget(resolved.probeTarget?.targetId);
        return existing;
      }
      return await openInBoundContext(binding, resolved, targetUrl, storeMarker, timeoutMs);
    } catch (error) {
      await closeTarget(resolved.probeTarget?.targetId);
      throw error;
    }
  }

  // Compatibility path for a new machine before the owner explicitly binds a
  // Chrome profile. It never accepts a page unless the expected store name is
  // visible, and it does not persist credentials.
  for (const target of pages.filter((item) => /meituan\.com/i.test(item.url || ''))) {
    const identity = await waitForMerchantIdentity(target.targetId, storeMarker, timeoutMs);
    if (!identity?.authenticatedStore) continue;
    const originalUrl = target.url;
    if (originalUrl !== targetUrl) {
      await proxy(`/navigate?target=${encodeURIComponent(target.targetId)}&url=${encodeURIComponent(targetUrl)}`);
    }
    return {
      targetId: target.targetId,
      browserContextId: target.browserContextId,
      originalUrl,
      openedByAutomation: false,
      source: 'unbound-authenticated-existing-page',
    };
  }

  throw new Error(
    `MEITUAN_STORE_CONTEXT_REQUIRED: 尚未绑定已登录“${storeMarker}”的 Chrome 资料；`
      + '请在正确资料中打开店铺后台后完成一次绑定。',
  );
}


export async function releasePetstarMerchantTab(opened) {
  if (!opened?.targetId) return;
  if (opened.openedByAutomation) {
    await closeTarget(opened.targetId);
    return;
  }
  if (!opened.originalUrl) return;
  const targets = await proxy('/targets').catch(() => []);
  const current = targets.find((item) => item.targetId === opened.targetId);
  if (!current || current.url === opened.originalUrl) return;
  await proxy(`/navigate?target=${encodeURIComponent(opened.targetId)}&url=${encodeURIComponent(opened.originalUrl)}`)
    .catch(() => {});
}
