#!/usr/bin/env node

import { spawn } from 'node:child_process';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';

const host = '127.0.0.1';
const proxyPort = Number(process.env.CDP_PROXY_PORT || 3456);
const localAppData = process.env.LOCALAPPDATA || '';
const activePortPath = path.join(localAppData, 'Google', 'Chrome', 'User Data', 'DevToolsActivePort');
const localStatePath = path.join(localAppData, 'Google', 'Chrome', 'User Data', 'Local State');
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
  const remoteDebuggingAllowed = readRemoteDebuggingConsent();
  if (remoteDebuggingAllowed === false) {
    throw new Error(
      'MEITUAN_CDP_AUTH_REQUIRED: 当前 Chrome 实例尚未允许远程调试；' +
      '为避免计划任务每小时重复弹出 Allow，系统不会自动发起连接。' +
      '请打开 chrome://inspect/#remote-debugging，勾选“Allow remote debugging for this browser instance”；下一次计划任务会自动恢复同步'
    );
  }
  if (remoteDebuggingAllowed === null) {
    throw new Error('MEITUAN_CDP_UNAVAILABLE: 无法读取 Chrome 远程调试授权状态；本次未发起连接，避免误触发授权弹窗');
  }
  const chromePort = await readChromePort();
  if (!chromePort) {
    throw new Error('MEITUAN_CDP_UNAVAILABLE: 已固定的 Chrome 调试连接不可用；请保持当前“宠速星宠物用品店”后台所在 Chrome 开启');
  }

  let targets = await getTargets();
  if (!targets) {
    startProxy();
    const deadline = Date.now() + 12000;
    while (Date.now() < deadline) {
      await new Promise((resolve) => setTimeout(resolve, 1000));
      targets = await getTargets(1800);
      if (targets) break;
    }
  }
  if (!targets) {
    throw new Error('MEITUAN_CDP_UNAVAILABLE: Chrome 正常，但本地 CDP 代理在 12 秒内未就绪；本次快速失败，下一小时自动重试');
  }

  console.log(`chrome: ok (port ${chromePort})`);
  console.log('proxy: ready');
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
