#!/usr/bin/env node

import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const projectRoot = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const checkerPath = path.join(projectRoot, 'agent', 'check_meituan_browser.mjs');
const temporaryRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'petstar-cdp-recovery-'));
const activePortPath = path.join(temporaryRoot, 'DevToolsActivePort');
const localStatePath = path.join(temporaryRoot, 'Local State');
const recoveryStatePath = path.join(temporaryRoot, 'cdp_authorization_prompt.json');

function listen(server) {
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolve(server.address().port));
  });
}

function runChecker(argumentsList, environment) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [checkerPath, ...argumentsList], {
      env: { ...process.env, ...environment },
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    let stdout = '';
    let stderr = '';
    child.stdout.on('data', (chunk) => { stdout += chunk.toString(); });
    child.stderr.on('data', (chunk) => { stderr += chunk.toString(); });
    child.once('error', reject);
    child.once('exit', (code) => resolve({ code, stdout, stderr }));
  });
}

const chromeServer = net.createServer((socket) => socket.end());
let proxyConnected = false;
let chromePort = 0;
const proxyServer = http.createServer((request, response) => {
  response.setHeader('content-type', 'application/json');
  if (request.url === '/health') {
    response.end(JSON.stringify({ status: 'ok', connected: proxyConnected, chromePort: proxyConnected ? chromePort : null }));
    return;
  }
  if (request.url === '/targets' && proxyConnected) {
    response.end(JSON.stringify([{ targetId: 'test-page', type: 'page', title: 'test', url: 'https://example.invalid' }]));
    return;
  }
  response.statusCode = 503;
  response.end(JSON.stringify({ error: 'not connected' }));
});

try {
  chromePort = await listen(chromeServer);
  const proxyPort = await listen(proxyServer);
  fs.writeFileSync(activePortPath, `${chromePort}\n/devtools/browser/test-instance\n`, 'utf8');
  fs.writeFileSync(localStatePath, JSON.stringify({ devtools: { remote_debugging: { 'user-enabled': true } } }), 'utf8');
  const environment = {
    CDP_PROXY_PORT: String(proxyPort),
    CDP_ACTIVE_PORT_PATH: activePortPath,
    CDP_LOCAL_STATE_PATH: localStatePath,
    CDP_RECOVERY_STATE_PATH: recoveryStatePath,
    CDP_WATCHDOG_AUTH_WAIT_MS: '100',
    CDP_WATCHDOG_RECONNECT_WAIT_MS: '100',
  };

  const blocked = await runChecker(['--watchdog'], environment);
  assert.equal(blocked.code, 1);
  assert.match(blocked.stderr, /MEITUAN_CDP_AUTH_REQUIRED/);
  const blockedState = JSON.parse(fs.readFileSync(recoveryStatePath, 'utf8'));
  assert.equal(blockedState.status, 'authorization_required');
  assert.equal(blockedState.pending_catch_up, true);
  assert.equal(blockedState.remote_debugging_global_flag, true, 'a stale global Allow flag must not prove this instance is connected');
  assert.equal(blockedState.connected_instance_fingerprint ?? null, null);

  proxyConnected = true;
  const recovered = await runChecker(['--watchdog'], environment);
  assert.equal(recovered.code, 0);
  assert.match(recovered.stdout, /proxy: ready/);
  const recoveredState = JSON.parse(fs.readFileSync(recoveryStatePath, 'utf8'));
  assert.equal(recoveredState.status, 'ready');
  assert.equal(recoveredState.pending_catch_up, true);
  assert.equal(recoveredState.connected_instance_fingerprint, recoveredState.current_instance_fingerprint);

  const marked = await runChecker(['--mark-catch-up-started'], environment);
  assert.equal(marked.code, 0);
  const markedState = JSON.parse(fs.readFileSync(recoveryStatePath, 'utf8'));
  assert.equal(markedState.pending_catch_up, false);

  console.log(JSON.stringify({ ok: true, scenarios: ['stale_global_allow_rejected', 'real_instance_connection_accepted', 'catch_up_claim_is_idempotent'] }));
} finally {
  await new Promise((resolve) => chromeServer.close(resolve));
  await new Promise((resolve) => proxyServer.close(resolve));
  fs.rmSync(temporaryRoot, { recursive: true, force: true });
}
