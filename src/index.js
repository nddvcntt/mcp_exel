/**
 * index.js — MCP Excel Server v2
 *
 * Endpoints:
 *   GET  /health          → trạng thái server + subprocess + tool count
 *   GET  /files           → danh sách file Excel (excel_files/ + output/)
 *   GET  /files/:name     → info 1 file
 *   DELETE /files/:name   → xóa file
 *   POST /upload          → upload file .xlsx vào excel_files/ (multipart)
 *   GET  /download/:name  → tải file Excel về (stream)
 *   GET  /tools           → danh sách 25 tools từ subprocess (cached)
 *   POST /mcp             → proxy tới uvx excel-mcp-server (Streamable HTTP MCP)
 *   GET  /mcp             → MCP discovery (405)
 */
'use strict';
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });

const express = require('express');
const cors    = require('cors');
const axios   = require('axios');
const path    = require('path');
const fs      = require('fs');
const multer  = require('multer');

const excelProxy = require('./excel-proxy');
const fileMgr    = require('./file-manager');

const PORT         = parseInt(process.env.PORT || '5003', 10);
const MAX_FILE_SIZE = parseInt(process.env.MAX_FILE_SIZE || '52428800', 10); // 50MB

const app = express();

// ── Multer — upload Excel files ───────────────────────────

const storage = multer.memoryStorage();
const upload  = multer({
  storage,
  limits: { fileSize: MAX_FILE_SIZE },
  fileFilter: (req, file, cb) => {
    const allowed = ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xlam'];
    const ext     = path.extname(file.originalname).toLowerCase();
    if (!allowed.includes(ext)) {
      return cb(new Error(`Extension không được hỗ trợ: ${ext}. Chỉ chấp nhận: ${allowed.join(', ')}`));
    }
    cb(null, true);
  },
});

// ── Middleware ────────────────────────────────────────────

app.use(cors({
  origin: '*',
  exposedHeaders: ['mcp-session-id', 'content-type', 'content-disposition'],
}));

// Raw body chỉ cho /mcp để proxy nguyên vẹn
app.use('/mcp', express.raw({ type: '*/*', limit: `${Math.round(MAX_FILE_SIZE / 1024 / 1024)}mb` }));
app.use(express.json({ limit: '10mb' }));

// ── Logging ───────────────────────────────────────────────

app.use((req, res, next) => {
  const start = Date.now();
  res.on('finish', () => {
    const ms = Date.now() - start;
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.path} ${res.statusCode} ${ms}ms`);
  });
  next();
});

// ── Simple rate limiter cho /mcp (50 req/s per IP) ────────

const rlMap = new Map();
function rateLimitMcp(req, res, next) {
  const ip  = req.ip || req.socket.remoteAddress || 'unknown';
  const now = Date.now();
  const bucket = rlMap.get(ip) || { count: 0, reset: now + 1000 };
  if (now > bucket.reset) { bucket.count = 0; bucket.reset = now + 1000; }
  bucket.count++;
  rlMap.set(ip, bucket);
  if (bucket.count > 50) {
    return res.status(429).json({ error: 'Too Many Requests', retry_after: Math.ceil((bucket.reset - now) / 1000) });
  }
  next();
}
// Cleanup rate limit map mỗi 60s để tránh memory leak
setInterval(() => {
  const now = Date.now();
  for (const [ip, b] of rlMap.entries()) { if (b.reset < now) rlMap.delete(ip); }
}, 60000);

// ── GET /health ───────────────────────────────────────────

app.get('/health', async (req, res) => {
  const quick = req.query.quick === '1' || req.query.quick === 'true';
  if (!excelProxy.isReady()) await excelProxy.probeReady();
  if (!quick && !excelProxy.isReady()) await excelProxy.ensureRunning();
  const subprocessReady = excelProxy.isReady();

  // Lấy tool list sau warm-up để health phản ánh khả năng dùng thật của MCP.
  const tools = subprocessReady ? await excelProxy.getToolsList() : [];
  const files  = fileMgr.listFiles();

  res.json({
    status    : 'ok',
    service   : 'mcp_excel',
    version   : '2.0.0',
    port      : PORT,
    subprocess: {
      ready       : subprocessReady,
      internal_url: excelProxy.getInternalUrl(),
      runtime     : excelProxy.getRuntimeConfig(),
      tool_count  : tools.length,
    },
    excel_files_path: fileMgr.EXCEL_FILES_PATH,
    output_dir      : fileMgr.OUTPUT_DIR,
    total_files     : files.length,
    timestamp       : new Date().toISOString(),
  });
});

// ── GET /tools ────────────────────────────────────────────
// Trả về danh sách tools từ subprocess (cached)

app.get('/tools', async (req, res) => {
  if (!excelProxy.isReady()) await excelProxy.probeReady();
  if (!excelProxy.isReady()) {
    // Thử warm-up
    await excelProxy.ensureRunning();
  }
  const tools = await excelProxy.getToolsList();
  res.json({
    success   : true,
    count     : tools.length,
    tools     : tools,
    subprocess: excelProxy.isReady() ? 'ready' : 'not_ready',
  });
});

// ── GET /files ────────────────────────────────────────────

app.get('/files', (req, res) => {
  try {
    const files = fileMgr.listFiles();
    res.json({ success: true, count: files.length, files });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ── GET /files/:name ──────────────────────────────────────

app.get('/files/:name', (req, res) => {
  try {
    const info = fileMgr.getFileInfo(req.params.name);
    res.json({ success: true, file: info });
  } catch (err) {
    res.status(404).json({ success: false, error: err.message });
  }
});

// ── DELETE /files/:name ───────────────────────────────────

app.delete('/files/:name', (req, res) => {
  try {
    const result = fileMgr.deleteFile(req.params.name);
    res.json({ success: true, ...result });
  } catch (err) {
    res.status(404).json({ success: false, error: err.message });
  }
});

// ── POST /upload ──────────────────────────────────────────
// Upload file Excel vào excel_files/

app.post('/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: 'Không có file được upload. Dùng field name "file".' });
    }

    // Tên file: ưu tiên query param, sau đó originalname
    const filename = (req.query.filename || req.file.originalname || 'upload.xlsx').replace(/[^a-zA-Z0-9_\-. ]/g, '_');
    const overwrite = req.query.overwrite === 'true';

    const info = fileMgr.uploadFile(filename, req.file.buffer, { overwrite });
    res.json({ success: true, message: `Upload thành công: ${filename}`, file: info });
  } catch (err) {
    const status = err.message?.includes('giới hạn') ? 413 : 400;
    res.status(status).json({ success: false, error: err.message });
  }
});

// Multer error handler (file too large, wrong type)
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError || err.message?.includes('Extension')) {
    return res.status(400).json({ success: false, error: err.message });
  }
  next(err);
});

// ── GET /download/:name ───────────────────────────────────
// Stream file Excel về client

app.get('/download/:name', (req, res) => {
  try {
    const filePath = fileMgr.getFilePath(req.params.name);
    const filename = path.basename(filePath);
    const stat     = fs.statSync(filePath);

    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Length', stat.size);

    const stream = fs.createReadStream(filePath);
    stream.on('error', err => {
      if (!res.headersSent) res.status(500).json({ error: 'Lỗi đọc file: ' + err.message });
    });
    stream.pipe(res);
  } catch (err) {
    res.status(404).json({ success: false, error: err.message });
  }
});

// ── GET /mcp — MCP discovery ──────────────────────────────

app.get('/mcp', (req, res) => {
  res.status(405).json({
    error  : 'Method Not Allowed',
    hint   : 'Use POST /mcp with MCP JSON-RPC or SSE protocol',
    service: 'mcp_excel',
    version: '2.0.0',
  });
});

// ── POST /mcp — proxy tới uvx subprocess ─────────────────

app.post('/mcp', rateLimitMcp, async (req, res) => {
  try {
    const ready = await excelProxy.ensureRunning();
    if (!ready) {
      return res.status(503).json({
        jsonrpc: '2.0', id: null,
        error: { code: -32000, message: 'excel-mcp-server subprocess không khởi động được' }
      });
    }

    const body = Buffer.isBuffer(req.body) ? req.body : Buffer.from(JSON.stringify(req.body));

    const forwardHeaders = {};
    if (req.headers['mcp-session-id']) forwardHeaders['mcp-session-id'] = req.headers['mcp-session-id'];
    if (req.headers['accept'])         forwardHeaders['accept']          = req.headers['accept'];
    if (req.headers['last-event-id'])  forwardHeaders['last-event-id']   = req.headers['last-event-id'];

    const resp = await axios.post(excelProxy.getInternalUrl() + '/mcp', body, {
      headers: {
        'Content-Type': req.headers['content-type'] || 'application/json',
        'Accept': req.headers['accept'] || 'application/json, text/event-stream',
        ...forwardHeaders,
      },
      responseType : 'stream',
      timeout      : 120000,
      validateStatus: () => true,
    });

    res.status(resp.status);
    const forwardResponseHeaders = ['content-type', 'mcp-session-id', 'transfer-encoding'];
    for (const h of forwardResponseHeaders) {
      if (resp.headers[h]) res.setHeader(h, resp.headers[h]);
    }

    resp.data.pipe(res);
    resp.data.on('error', (err) => {
      if (!res.headersSent) res.status(500).end();
      console.error('[mcp_excel] Proxy stream error:', err.message);
    });

  } catch (err) {
    console.error('[mcp_excel] POST /mcp error:', err.message);
    if (!res.headersSent) {
      res.status(502).json({
        jsonrpc: '2.0', id: null,
        error: { code: -32000, message: `Proxy error: ${err.message}` }
      });
    }
  }
});

// ── 404 handler ───────────────────────────────────────────

app.use((req, res) => {
  res.status(404).json({ error: 'Not found', path: req.path });
});

// ── Global error handler ──────────────────────────────────

app.use((err, req, res, next) => {
  console.error('[mcp_excel] Unhandled error:', err.message);
  if (!res.headersSent) res.status(500).json({ error: 'Internal Server Error', message: err.message });
});

// ── Start server ──────────────────────────────────────────

app.listen(PORT, () => {
  console.log(`
╔══════════════════════════════════════════════════╗
║         MCP Excel Server  v2.0.0                ║
╠══════════════════════════════════════════════════╣
║  Port     : ${PORT}                              ║
║  MCP URL  : http://localhost:${PORT}/mcp          ║
║  Health   : http://localhost:${PORT}/health       ║
║  Files    : http://localhost:${PORT}/files        ║
║  Upload   : POST http://localhost:${PORT}/upload  ║
║  Download : GET  http://localhost:${PORT}/download/:name  ║
║  Tools    : GET  http://localhost:${PORT}/tools   ║
╚══════════════════════════════════════════════════╝
`);
  // Warm-up subprocess ngay khi server bật
  excelProxy.startSubprocess().catch(e => console.error('[mcp_excel] Warm-up error:', e.message));
});

// Graceful shutdown
process.on('SIGTERM', () => { excelProxy.stopSubprocess(); process.exit(0); });
process.on('SIGINT',  () => { excelProxy.stopSubprocess(); process.exit(0); });
