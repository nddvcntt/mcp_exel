/**
 * excel-proxy.js
 * Quản lý subprocess uvx excel-mcp-server và proxy request tới nó.
 *
 * Cải tiến v2:
 *  - Fix restart counter: reset sau idle timeout (không cạn limit oan)
 *  - Exponential backoff khi chờ subprocess sẵn sàng
 *  - Windows process leak fix: taskkill /T fallback để dọn subprocess tree
 *  - getToolsList(): cache danh sách 25 tools để health endpoint trả về
 *  - Tách startSubprocess / stopSubprocess / ensureRunning rõ ràng hơn
 */
'use strict';
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const { spawn, spawnSync } = require('child_process');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const INTERNAL_PORT     = parseInt(process.env.EXCEL_MCP_INTERNAL_PORT || '5013', 10);
const EXCEL_FILES_PATH  = process.env.EXCEL_FILES_PATH || path.join(__dirname, '..', 'excel_files');
const IDLE_TIMEOUT_MS   = parseInt(process.env.IDLE_TIMEOUT_MS || '900000', 10);
const INTERNAL_URL      = `http://localhost:${INTERNAL_PORT}`;
const EXCEL_MCP_COMMAND = process.env.EXCEL_MCP_COMMAND || 'uvx';
const EXCEL_MCP_ARGS    = (process.env.EXCEL_MCP_ARGS || 'excel-mcp-server,streamable-http')
  .split(',').map(s => s.trim()).filter(Boolean);
const EXCEL_MCP_SHELL   = process.env.EXCEL_MCP_SHELL
  ? process.env.EXCEL_MCP_SHELL !== 'false'
  : false;

let subprocess    = null;
let isStarting    = false;
let idleTimer     = null;
let restartCount  = 0;
const MAX_RESTARTS = 3;
let ready         = false;
let cachedTools   = null;   // Cache danh sách tools
const intentionallyStopped = new WeakSet();

// Đảm bảo thư mục excel_files tồn tại
fs.mkdirSync(EXCEL_FILES_PATH, { recursive: true });

// ── Idle timer ────────────────────────────────────────────

function resetIdleTimer() {
  if (idleTimer) clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    console.log('[ExcelProxy] Idle timeout — stopping subprocess.');
    stopSubprocess();
  }, IDLE_TIMEOUT_MS);
}

// ── Stop subprocess ───────────────────────────────────────

function terminateProcessTree(proc, forceAfterMs = 3000) {
  if (!proc) return;

  if (process.platform === 'win32' && proc.pid) {
    try {
      spawnSync('taskkill', ['/PID', String(proc.pid), '/T', '/F'], {
        stdio: 'ignore',
        windowsHide: true,
      });
    } catch {}
    return;
  }

  try { proc.kill('SIGTERM'); } catch {}

  if (proc.pid) {
    const fallback = setTimeout(() => {
      if (proc.exitCode !== null) return;
      try { proc.kill('SIGKILL'); } catch {}
    }, forceAfterMs);
    if (typeof fallback.unref === 'function') fallback.unref();
    proc.once('exit', () => clearTimeout(fallback));
  }
}

function stopSubprocess() {
  ready        = false;
  cachedTools  = null;
  restartCount = 0; // FIX: reset sau idle stop để tránh cạn restart limit oan

  if (subprocess) {
    const proc = subprocess;
    intentionallyStopped.add(proc);
    subprocess  = null;
    terminateProcessTree(proc);
  }

  if (idleTimer) { clearTimeout(idleTimer); idleTimer = null; }
}

// ── Wait for subprocess ready (exponential backoff) ───────

async function waitForReady(maxWaitMs = 240000) {
  const start   = Date.now();
  let   delayMs = 1000;
  console.log(`[ExcelProxy] Chờ subprocess sẵn sàng (tối đa ${maxWaitMs / 1000}s)...`);

  while (Date.now() - start < maxWaitMs) {
    try {
      const r = await axios.post(`${INTERNAL_URL}/mcp`, {
        jsonrpc: '2.0', id: 1, method: 'initialize',
        params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'probe', version: '1' } }
      }, { timeout: 3000, headers: { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' } });
      if (r.status >= 200 && r.status < 500) return true;
    } catch (e) {
      // 400/405 → server đang chạy
      if (e.response && e.response.status >= 400) return true;
    }
    await new Promise(r => setTimeout(r, delayMs));
    delayMs = Math.min(delayMs * 1.5, 8000); // exponential backoff, max 8s
  }
  return false;
}

// ── Start subprocess ──────────────────────────────────────

async function startSubprocess() {
  if (subprocess || isStarting) return;
  isStarting = true;
  ready      = false;
  console.log(`[ExcelProxy] Khởi động ${EXCEL_MCP_COMMAND} ${EXCEL_MCP_ARGS.join(' ')} trên port ${INTERNAL_PORT}...`);

  const env = {
    ...process.env,
    EXCEL_FILES_PATH,
    FASTMCP_PORT: String(INTERNAL_PORT),
  };

  const child = spawn(EXCEL_MCP_COMMAND, EXCEL_MCP_ARGS, {
    env,
    stdio: ['ignore', 'pipe', 'pipe'],
    shell: EXCEL_MCP_SHELL,
    windowsHide: true,
  });
  subprocess = child;

  child.stdout.on('data', d => process.stdout.write(`[ExcelMCP] ${d}`));
  child.stderr.on('data', d => process.stderr.write(`[ExcelMCP][ERR] ${d}`));

  child.on('exit', (code, signal) => {
    console.log(`[ExcelProxy] Subprocess exit code=${code} signal=${signal}`);
    if (subprocess !== child) return;

    subprocess = null;
    ready = false;
    cachedTools = null;
    isStarting = false;

    if (intentionallyStopped.has(child)) {
      restartCount = 0;
      return;
    }

    if (restartCount < MAX_RESTARTS) {
      restartCount++;
      console.log(`[ExcelProxy] Auto-restart (${restartCount}/${MAX_RESTARTS}) sau 2s...`);
      setTimeout(startSubprocess, 2000);
    } else {
      console.error('[ExcelProxy] Đã thử restart tối đa. Dừng hẳn.');
    }
  });

  const ok = await waitForReady();
  isStarting = false;

  if (ok) {
    ready        = true;
    restartCount = 0;
    console.log(`[ExcelProxy] ✅ uvx excel-mcp-server sẵn sàng trên port ${INTERNAL_PORT}`);
    // Warm-up cache tool list
    getToolsList().catch(() => {});
  } else {
    console.error('[ExcelProxy] ❌ Subprocess không sẵn sàng sau thời gian chờ.');
    terminateProcessTree(child, 0);
  }
}

// ── Ensure running ────────────────────────────────────────

async function ensureRunning() {
  if (ready) { resetIdleTimer(); return true; }
  if (!subprocess && !isStarting) await startSubprocess();

  // Đợi thêm nếu đang trong quá trình start
  if (isStarting) {
    for (let i = 0; i < 480; i++) {
      await new Promise(r => setTimeout(r, 500));
      if (ready) break;
    }
  }
  if (ready) resetIdleTimer();
  return ready;
}

// ── Get tools list (cached) ───────────────────────────────

async function getToolsList() {
  if (cachedTools) return cachedTools;
  if (!ready) return [];

  try {
    // Cần initialize session trước
    const r1 = await axios.post(`${INTERNAL_URL}/mcp`, {
      jsonrpc: '2.0', id: 100, method: 'initialize',
      params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'health-probe', version: '1' } }
    }, { timeout: 5000, headers: { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' } });

    const sid = r1.headers['mcp-session-id'];

    const r2 = await axios.post(`${INTERNAL_URL}/mcp`, {
      jsonrpc: '2.0', id: 101, method: 'tools/list', params: {}
    }, {
      timeout: 10000,
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json, text/event-stream',
        ...(sid ? { 'mcp-session-id': sid } : {}),
      },
    });

    // Parse SSE or plain JSON
    let tools = [];
    const rawData = typeof r2.data === 'string' ? r2.data : JSON.stringify(r2.data);
    const sseMatch = rawData.match(/^data:\s*(.+)$/m);
    const parsed = sseMatch ? JSON.parse(sseMatch[1]) : (typeof r2.data === 'object' ? r2.data : null);
    tools = parsed?.result?.tools?.map(t => t.name) || [];

    cachedTools = tools;
    return tools;
  } catch (err) {
    console.warn('[ExcelProxy] getToolsList failed:', err.message);
    return [];
  }
}

// ── Probe ready ───────────────────────────────────────────

async function probeReady() {
  try {
    const r = await axios.post(`${INTERNAL_URL}/mcp`, {
      jsonrpc: '2.0', id: 99, method: 'initialize',
      params: { protocolVersion: '2024-11-05', capabilities: {}, clientInfo: { name: 'probe', version: '1' } }
    }, { timeout: 3000, headers: { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' } });
    if (r.status >= 200 && r.status < 500) { ready = true; return true; }
  } catch (e) {
    if (e.response && e.response.status >= 400) { ready = true; return true; }
  }
  return false;
}

// ── Getters ───────────────────────────────────────────────

function isReady()          { return ready; }
function getInternalUrl()   { return INTERNAL_URL; }
function getRuntimeConfig() {
  return {
    command: EXCEL_MCP_COMMAND,
    args: EXCEL_MCP_ARGS,
    shell: EXCEL_MCP_SHELL,
    idle_timeout_ms: IDLE_TIMEOUT_MS,
    restart_count: restartCount,
    max_restarts: MAX_RESTARTS,
  };
}

module.exports = {
  ensureRunning,
  stopSubprocess,
  startSubprocess,
  probeReady,
  isReady,
  getInternalUrl,
  getRuntimeConfig,
  getToolsList,
};
