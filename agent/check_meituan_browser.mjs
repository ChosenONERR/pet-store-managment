#!/usr/bin/env node

import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const host = '127.0.0.1';
const proxyPort = Number(process.env.CDP_PROXY_PORT || 3456);
const localAppData = process.env.LOCALAPPDATA || '';
const activePortPath = process.env.CDP_ACTIVE_PORT_PATH || path.join(localAppData, 'Google', 'Chrome', 'User Data', 'DevToolsActivePort');
const localStatePath = process.env.CDP_LOCAL_STATE_PATH || path.join(localAppData, 'Google', 'Chrome', 'User Data', 'Local State');
const agentRoot = path.dirname(fileURLToPath(import.meta.url));
const promptStatePath = process.env.CDP_RECOVERY_STATE_PATH || path.join(agentRoot, 'state', 'cdp_authorization_prompt.json');
const watchdogMode = process.argv.includes('--watchdog');
const markCatchUpStartedMode = process.argv.includes('--mark-catch-up-started');
const markSyncSuccessMode = process.argv.includes('--mark-sync-success');
const proxyCandidates = [
  path.join(os.homedir(), '.agents', 'skills', 'web-access', 'scripts', 'cdp-proxy.mjs'),
  path.join(os.homedir(), '.codex', 'skills', 'web-access', 'scripts', 'cdp-proxy.mjs'),
];

function checkPort(port, timeoutMs = 1200) {
  return new Promise((resolve) => {
    const socket = net.createConnection(port, host);
    const timer = setTimeout(() => {
      socket.destroy();
      resolve(false);
    }, timeoutMs);
    socket.once('connect', () => {
      clearTimeout(timer);
      socket.destroy();
      resolve(true);
    });
    socket.once('error', () => {
      clearTimeout(timer);
      resolve(false);
    });
  });
}

async function readChromePort() {
  try {
    const firstLine = fs.readFileSync(activePortPath, 'utf8').trim().split(/\r?\n/)[0];
    const port = Number(firstLine);
    if (Number.isInteger(port) && port > 0 && port < 65536 && await checkPort(port)) return port;
  } catch (_) {}
  return null;
}

function readChromeInstanceFingerprint() {
  try {
    const lines = fs.readFileSync(activePortPath, 'utf8').trim().split(/\r?\n/);
    if (!lines[0] || !lines[1]) return null;
    return createHash('sha256').update(`${lines[0]}\n${lines[1]}`).digest('hex');
  } catch (_) {
    return null;
  }
}

function readCdpState() {
  try {
    const state = JSON.parse(fs.readFileSync(promptStatePath, 'utf8'));
    // Version 1 stored only the instance that had already shown an Allow
    // prompt. Keep that information when upgrading the state in place.
    if (state?.version === 1) {
      return {
        version: 2,
        prompted_instance_fingerprint: state.instance_fingerprint || null,
        connected_instance_fingerprint: null,
        pending_catch_up: false,
      };
    }
    return state && typeof state === 'object' ? state : { version: 2 };
  } catch (_) {
    return { version: 2 };
  }
}

function saveCdpState(changes) {
  const current = readCdpState();
  fs.mkdirSync(path.dirname(promptStatePath), { recursive: true });
  const payload = JSON.stringify({
    ...current,
    ...changes,
    version: 2,
    updated_at: new Date().toISOString(),
  }, null, 2);
  const temporary = `${promptStatePath}.${process.pid}.tmp`;
  fs.writeFileSync(temporary, payload, 'utf8');
  fs.renameSync(temporary, promptStatePath);
}

function markCatchUpStarted() {
  const state = readCdpState();
  saveCdpState({
    pending_catch_up: false,
    catch_up_started_for_instance: state.connected_instance_fingerprint || null,
    catch_up_started_at: new Date().toISOString(),
  });
  console.log('catch-up: marked started');
}

function markSyncSuccess() {
  const state = readCdpState();
  saveCdpState({
    pending_catch_up: false,
    last_successful_sync_instance: state.connected_instance_fingerprint || null,
    last_successful_sync_at: new Date().toISOString(),
  });
  console.log('sync: marked successful');
}

function readRemoteDebuggingConsent() {
  try {
    const localState = JSON.parse(fs.readFileSync(localStatePath, 'utf8'));
    return localState?.devtools?.remote_debugging?.['user-enabled'] === true;
  } catch (_) {
    return null;
  }
}

async function getTargets(timeoutMs = 2500) {
  try {
    const response = await fetch(`http://${host}:${proxyPort}/targets`, {
      signal: AbortSignal.timeout(timeoutMs),
    });
    const body = await response.json();
    return Array.isArray(body) ? body : null;
  } catch (_) {
    return null;
  }
}

async function getProxyHealth(timeoutMs = 1800) {
  try {
    const response = await fetch(`http://${host}:${proxyPort}/health`, {
      signal: AbortSignal.timeout(timeoutMs),
    });
    const body = await response.json();
    return body?.status === 'ok' ? body : null;
  } catch (_) {
    return null;
  }
}

async function waitForProxyTargets(timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const remaining = Math.max(1000, deadline - Date.now());
    const targets = await getTargets(Math.min(5000, remaining));
    if (targets) return targets;
    await new Promise((resolve) => setTimeout(resolve, 1500));
  }
  return null;
}

function startProxy() {
  const proxyScript = proxyCandidates.find((candidate) => fs.existsSync(candidate));
  if (!proxyScript) throw new Error('MEITUAN_CDP_UNAVAILABLE: cdp-proxy.mjs not found');
  const logPath = path.join(os.tmpdir(), 'cdp-proxy.log');
  const logFd = fs.openSync(logPath, 'a');
  const child = spawn(process.execPath, [proxyScript], {
    detached: true,
    windowsHide: true,
    stdio: ['ignore', logFd, logFd],
  });
  child.unref();
  fs.closeSync(logFd);
}

async function main() {
  if (markCatchUpStartedMode) {
    markCatchUpStarted();
    return;
  }
  if (markSyncSuccessMode) {
    markSyncSuccess();
    return;
  }

  const chromePort = await readChromePort();
  if (!chromePort) {
    throw new Error('MEITUAN_CDP_UNAVAILABLE: 已固定的 Chrome 调试连接不可用；请保持当前“宠速星宠物用品店”后台所在 Chrome 开启');
  }

  const cdpState = readCdpState();
  const instanceFingerprint = readChromeInstanceFingerprint();
  if (!instanceFingerprint) {
    throw new Error('MEITUAN_CDP_UNAVAILABLE: 无法识别当前 Chrome 实例；本次未发起授权连接');
  }

  let proxyHealth = await getProxyHealth();
  // A proxy may still be alive while Chrome has restarted. Only a connection
  // to the port recorded by the current DevToolsActivePort proves that this
  // browser instance is authorized and usable.
  const proxyMatchesCurrentInstance = proxyHealth?.connected && Number(proxyHealth.chromePort) === chromePort;
  let targets = proxyMatchesCurrentInstance ? await getTargets() : null;
  if (targets) {
    saveCdpState({
      status: 'ready',
      current_instance_fingerprint: instanceFingerprint,
      connected_instance_fingerprint: instanceFingerprint,
      connected_at: new Date().toISOString(),
      last_error: null,
    });
    console.log(`chrome: ok (port ${chromePort})`);
    console.log('proxy: ready');
    return;
  }

  const remoteDebuggingAllowed = readRemoteDebuggingConsent();
  // Chrome's Local State flag survives a browser restart, but the Allow grant
  // is scoped to one browser instance. The only durable truth is therefore a
  // fingerprint that has actually connected through the proxy before.
  const currentInstanceWasConnected = cdpState.connected_instance_fingerprint === instanceFingerprint;
  const currentInstanceWasPrompted = cdpState.prompted_instance_fingerprint === instanceFingerprint;
  if (!currentInstanceWasConnected && !currentInstanceWasPrompted) {
    saveCdpState({
      status: 'authorization_required',
      current_instance_fingerprint: instanceFingerprint,
      prompted_instance_fingerprint: instanceFingerprint,
      prompted_at: new Date().toISOString(),
      pending_catch_up: true,
      remote_debugging_global_flag: remoteDebuggingAllowed,
      last_error: 'new Chrome instance has not established a real CDP connection',
    });
  }

  if (!proxyHealth) {
    startProxy();
  }
  // A running proxy can be alive while its Chrome WebSocket is temporarily
  // disconnected. Repeated /targets calls make that same proxy rediscover the
  // current DevToolsActivePort and reconnect; starting duplicate proxies cannot
  // repair this state because port 3456 is already occupied.
  const newOrUnverifiedInstance = !currentInstanceWasConnected;
  const watchdogAuthorizationWaitMs = Number(process.env.CDP_WATCHDOG_AUTH_WAIT_MS || 8000);
  const watchdogReconnectWaitMs = Number(process.env.CDP_WATCHDOG_RECONNECT_WAIT_MS || 12000);
  const waitMs = watchdogMode ? (newOrUnverifiedInstance ? watchdogAuthorizationWaitMs : watchdogReconnectWaitMs) : (newOrUnverifiedInstance ? 15000 : 45000);
  targets = await waitForProxyTargets(waitMs);
  if (!targets) {
    if (newOrUnverifiedInstance) {
      saveCdpState({
        status: 'authorization_required',
        current_instance_fingerprint: instanceFingerprint,
        prompted_instance_fingerprint: instanceFingerprint,
        pending_catch_up: true,
        remote_debugging_global_flag: remoteDebuggingAllowed,
        last_error: 'current Chrome instance is waiting for Allow',
      });
      throw new Error(
        'MEITUAN_CDP_AUTH_REQUIRED: 检测到新的 Chrome 实例，但该实例尚未完成远程调试授权。' +
        '系统已主动发起连接以显示一次 Allow；请点击 Allow（或在 chrome://inspect/#remote-debugging 中允许）。' +
        '完成后看门狗会在 5 分钟内自动重连并补跑同步，无需等到下一整点'
      );
    }
    saveCdpState({
      status: 'proxy_unavailable',
      current_instance_fingerprint: instanceFingerprint,
      pending_catch_up: true,
      last_error: `known Chrome instance did not reconnect within ${Math.round(waitMs / 1000)} seconds`,
    });
    throw new Error(`MEITUAN_CDP_UNAVAILABLE: 当前 Chrome 实例曾成功授权，但本地 CDP 代理在 ${Math.round(waitMs / 1000)} 秒自愈窗口内仍未重连；看门狗会继续自动修复`);
  }

  saveCdpState({
    status: 'ready',
    current_instance_fingerprint: instanceFingerprint,
    connected_instance_fingerprint: instanceFingerprint,
    connected_at: new Date().toISOString(),
    pending_catch_up: Boolean(cdpState.pending_catch_up || newOrUnverifiedInstance),
    remote_debugging_global_flag: remoteDebuggingAllowed,
    last_error: null,
  });
  console.log(`chrome: ok (port ${chromePort})`);
  console.log('proxy: ready');
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
