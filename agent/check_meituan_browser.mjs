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
const activePortPath = path.join(localAppData, 'Google', 'Chrome', 'User Data', 'DevToolsActivePort');
const localStatePath = path.join(localAppData, 'Google', 'Chrome', 'User Data', 'Local State');
const agentRoot = path.dirname(fileURLToPath(import.meta.url));
const promptStatePath = path.join(agentRoot, 'state', 'cdp_authorization_prompt.json');
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

function readPromptedFingerprint() {
  try {
    return JSON.parse(fs.readFileSync(promptStatePath, 'utf8'))?.instance_fingerprint || null;
  } catch (_) {
    return null;
  }
}

function savePromptedFingerprint(fingerprint) {
  fs.mkdirSync(path.dirname(promptStatePath), { recursive: true });
  const payload = JSON.stringify({
    version: 1,
    instance_fingerprint: fingerprint,
    prompted_at: new Date().toISOString(),
  });
  const temporary = `${promptStatePath}.${process.pid}.tmp`;
  fs.writeFileSync(temporary, payload, 'utf8');
  fs.renameSync(temporary, promptStatePath);
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
  const chromePort = await readChromePort();
  if (!chromePort) {
    throw new Error('MEITUAN_CDP_UNAVAILABLE: 已固定的 Chrome 调试连接不可用；请保持当前“宠速星宠物用品店”后台所在 Chrome 开启');
  }

  let proxyHealth = await getProxyHealth();
  let targets = proxyHealth?.connected ? await getTargets() : null;
  if (targets) {
    console.log(`chrome: ok (port ${chromePort})`);
    console.log('proxy: ready');
    return;
  }

  const remoteDebuggingAllowed = readRemoteDebuggingConsent();
  const instanceFingerprint = readChromeInstanceFingerprint();
  if (remoteDebuggingAllowed === false) {
    if (!instanceFingerprint) {
      throw new Error('MEITUAN_CDP_UNAVAILABLE: 无法识别当前 Chrome 实例；本次未发起授权连接');
    }
    if (readPromptedFingerprint() === instanceFingerprint) {
      throw new Error(
        'MEITUAN_CDP_AUTH_REQUIRED: 当前 Chrome 实例仍未允许远程调试；' +
        '系统已为这个实例主动弹出过一次 Allow，本次不重复打扰。' +
        '请打开 chrome://inspect/#remote-debugging，勾选“Allow remote debugging for this browser instance”；下一次计划任务会自动恢复同步'
      );
    }
    // A new Chrome instance gets one active connection attempt so the browser
    // can surface its Allow prompt. Persist the fingerprint before connecting;
    // a missed prompt must not be retriggered by every hourly task.
    savePromptedFingerprint(instanceFingerprint);
  } else if (remoteDebuggingAllowed === null) {
    throw new Error('MEITUAN_CDP_UNAVAILABLE: 无法读取 Chrome 远程调试授权状态；本次未发起连接，避免误触发授权弹窗');
  }

  if (!proxyHealth) {
    startProxy();
  }
  // A running proxy can be alive while its Chrome WebSocket is temporarily
  // disconnected. Repeated /targets calls make that same proxy rediscover the
  // current DevToolsActivePort and reconnect; starting duplicate proxies cannot
  // repair this state because port 3456 is already occupied.
  targets = await waitForProxyTargets(remoteDebuggingAllowed === false ? 15000 : 45000);
  if (!targets) {
    if (remoteDebuggingAllowed === false) {
      throw new Error(
        'MEITUAN_CDP_AUTH_REQUIRED: 已为当前 Chrome 实例主动请求一次远程调试授权，但尚未允许；' +
        '为避免每小时重复弹窗，后续定时任务只保留飞书提醒。请在 chrome://inspect/#remote-debugging 完成 Allow'
      );
    }
    throw new Error('MEITUAN_CDP_UNAVAILABLE: Chrome 正常，但本地 CDP 代理在 45 秒自愈窗口内仍未连接；本次安全退出，下一小时自动重试');
  }

  console.log(`chrome: ok (port ${chromePort})`);
  console.log('proxy: ready');
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
