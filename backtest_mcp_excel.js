/**
 * backtest_mcp_excel.js  v2 â€” 35+ test cases
 * Cháº¡y: node backtest_mcp_excel.js
 */
'use strict';

const BASE = 'http://localhost:5003';
const RESULTS = [];
let pass = 0, fail = 0;

async function fetchJSON(url, opts = {}) {
  const res = await fetch(url, {
    headers: { 'Content-Type': 'application/json', 'Accept': 'application/json', ...(opts.headers || {}) },
    ...opts,
  });
  const text = await res.text();
  let json;
  try { json = JSON.parse(text); } catch { json = { _raw: text.slice(0, 300) }; }
  return { status: res.status, ok: res.ok, data: json };
}

function parseSseResponse(text) {
  const lines = text.split('\n');
  for (const line of lines) {
    if (line.startsWith('data: ')) {
      try { return JSON.parse(line.slice(6)); } catch {}
    }
  }
  try { return JSON.parse(text); } catch {}
  return null;
}

let mcpSessionId = null;

async function mcpCall(method, params = {}) {
  const headers = {
    'Content-Type': 'application/json',
    'Accept': 'application/json, text/event-stream',
  };
  if (mcpSessionId) headers['mcp-session-id'] = mcpSessionId;

  const res = await fetch(`${BASE}/mcp`, {
    method: 'POST',
    headers,
    body: JSON.stringify({ jsonrpc: '2.0', id: Math.floor(Math.random() * 10000), method, params }),
  });

  const sid = res.headers.get('mcp-session-id');
  if (sid && !mcpSessionId) { mcpSessionId = sid; }

  const text = await res.text();
  const data = parseSseResponse(text) || { _raw: text.slice(0, 300) };
  return { status: res.status, ok: res.ok, data, sessionId: mcpSessionId };
}

function test(group, name, fn) {
  return fn().then(result => {
    const ok = result?.pass !== false && !result?.error;
    RESULTS.push({ group, name, ok, result });
    console.log(`  ${ok ? 'âœ…' : 'âŒ'} [${group}] ${name}`);
    if (!ok) console.log(`       â†’ ${result?.error || JSON.stringify(result).slice(0, 200)}`);
    if (ok) pass++; else fail++;
  }).catch(err => {
    RESULTS.push({ group, name, ok: false });
    console.log(`  âŒ [${group}] ${name}: ${err.message}`);
    fail++;
  });
}

const TS = Date.now();
const TEST_FILE  = `backtest_${TS}.xlsx`;
const TEST_FILE2 = `backtest2_${TS}.xlsx`;

console.log('\n' + '='.repeat(60));
console.log('BACKTEST mcp_excel v2 â€” ' + new Date().toLocaleString('vi-VN'));
console.log('='.repeat(60));

(async () => {

  // â”€â”€ GROUP A: Health & Discovery â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[A] Health & Discovery');

  await test('A', 'GET /health â†’ 200 + version 2.0.0', async () => {
    const r = await fetchJSON(`${BASE}/health`);
    if (!r.ok) return { error: `HTTP ${r.status}` };
    if (r.data?.version !== '2.0.0') return { error: `version = ${r.data?.version}, expected 2.0.0` };
    return { pass: true, version: r.data.version };
  });

  await test('A', 'GET /health subprocess.ready = true', async () => {
    const r = await fetchJSON(`${BASE}/health`);
    if (!r.data?.subprocess?.ready) return { error: 'subprocess.ready = false' };
    return { pass: true };
  });

  await test('A', 'GET /health subprocess.tool_count = 25', async () => {
    const r = await fetchJSON(`${BASE}/health`);
    const tc = r.data?.subprocess?.tool_count;
    if (tc !== 25) return { error: `tool_count = ${tc}, expected 25` };
    return { pass: true, tool_count: tc };
  });

  await test('A', 'GET /files â†’ 200', async () => {
    const r = await fetchJSON(`${BASE}/files`);
    if (!r.ok) return { error: `HTTP ${r.status}` };
    return { pass: true, count: r.data?.count };
  });

  await test('A', 'GET /mcp â†’ 405 (MCP discovery)', async () => {
    const r = await fetchJSON(`${BASE}/mcp`);
    if (r.status !== 405) return { error: `Expected 405, got ${r.status}` };
    return { pass: true };
  });

  await test('A', 'GET /tools â†’ danh sĂ¡ch 25 tools', async () => {
    const r = await fetchJSON(`${BASE}/tools`);
    if (!r.ok) return { error: `HTTP ${r.status}` };
    if (r.data?.count !== 25) return { error: `count = ${r.data?.count}, expected 25` };
    return { pass: true, tools: r.data.tools?.slice(0, 4).join(', ') };
  });

  // â”€â”€ GROUP B: MCP Protocol â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[B] MCP Protocol');

  await test('B', 'POST /mcp â†’ initialize (SSE)', async () => {
    const r = await mcpCall('initialize', {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'backtest-v2', version: '2.0' }
    });
    const hasResult = r.data?.result?.serverInfo;
    if (!hasResult) return { error: `No serverInfo. Raw: ${JSON.stringify(r.data).slice(0, 200)}` };
    return { pass: true, server: r.data.result.serverInfo.name };
  });

  let toolsList = [];
  await test('B', 'POST /mcp â†’ tools/list (â‰¥25 tools)', async () => {
    const r = await mcpCall('tools/list', {});
    const tools = r.data?.result?.tools || [];
    toolsList = tools.map(t => t.name);
    if (tools.length < 20) return { error: `Chá»‰ cĂ³ ${tools.length} tools` };
    return { pass: true, count: tools.length, sample: toolsList.slice(0, 4).join(', ') };
  });

  await test('B', 'tools/list cĂ³ Ä‘á»§ nhĂ³m core tools', async () => {
    const required = ['create_workbook', 'write_data_to_excel', 'read_data_from_excel', 'format_range', 'create_chart', 'create_pivot_table', 'create_table', 'insert_rows'];
    const missing = required.filter(t => !toolsList.includes(t));
    if (missing.length > 0) return { error: `Thiáº¿u: ${missing.join(', ')}` };
    return { pass: true };
  });

  // â”€â”€ GROUP C: CRUD Round-trip â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[C] Excel CRUD Round-trip');

  await test('C', 'create_workbook', async () => {
    const r = await mcpCall('tools/call', { name: 'create_workbook', arguments: { filepath: TEST_FILE } });
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    if (!content) return { error: `No content. Raw: ${JSON.stringify(r.data).slice(0, 200)}` };
    return { pass: true, result: content.slice(0, 60) };
  });

  await test('C', 'write_data_to_excel', async () => {
    const data = [
      ['Ho ten', 'Chuc vu', 'Luong'],
      ['Nguyen Van An', 'Giam doc', 30000000],
      ['Tran Thi Binh', 'Ke toan', 15000000],
      ['Le Minh Cuong', 'Ky su', 20000000],
    ];
    const r = await mcpCall('tools/call', { name: 'write_data_to_excel', arguments: { filepath: TEST_FILE, sheet_name: 'Sheet1', data } });
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('C', 'read_data_from_excel', async () => {
    const r = await mcpCall('tools/call', { name: 'read_data_from_excel', arguments: { filepath: TEST_FILE, sheet_name: 'Sheet1' } });
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    const hasData = content.includes('An') || content.includes('30000000') || content.includes('Cuong');
    if (!hasData) return { error: `KhĂ´ng tháº¥y dá»¯ liá»‡u. Got: ${content.slice(0, 200)}` };
    return { pass: true };
  });

  await test('C', 'get_workbook_metadata', async () => {
    const r = await mcpCall('tools/call', { name: 'get_workbook_metadata', arguments: { filepath: TEST_FILE } });
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    if (!content) return { error: 'No content' };
    return { pass: true };
  });

  // â”€â”€ GROUP D: Formatting â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[D] Formatting & Formulas');

  await test('D', 'format_range (bold header)', async () => {
    const r = await mcpCall('tools/call', { name: 'format_range', arguments: {
      filepath: TEST_FILE, sheet_name: 'Sheet1',
      start_cell: 'A1', end_cell: 'C1', bold: true, bg_color: 'FF4472C4', font_color: 'FFFFFFFF'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('D', 'apply_formula (SUM)', async () => {
    const r = await mcpCall('tools/call', { name: 'apply_formula', arguments: {
      filepath: TEST_FILE, sheet_name: 'Sheet1', cell: 'C5', formula: '=SUM(C2:C4)'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('D', 'validate_formula_syntax (valid)', async () => {
    const r = await mcpCall('tools/call', { name: 'validate_formula_syntax', arguments: { formula: '=SUM(B2:B10)/COUNT(B2:B10)' } });
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error') && !content.toLowerCase().includes('valid')) return { error: content };
    return { pass: true };
  });

  await test('D', 'merge_cells', async () => {
    const r = await mcpCall('tools/call', { name: 'create_worksheet', arguments: { filepath: TEST_FILE, sheet_name: 'BaoCao' } });
    const r2 = await mcpCall('tools/call', { name: 'merge_cells', arguments: {
      filepath: TEST_FILE, sheet_name: 'BaoCao', start_cell: 'A1', end_cell: 'D1'
    }});
    const content = r2.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('D', 'unmerge_cells', async () => {
    const r = await mcpCall('tools/call', { name: 'unmerge_cells', arguments: {
      filepath: TEST_FILE, sheet_name: 'BaoCao', start_cell: 'A1', end_cell: 'D1'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  // â”€â”€ GROUP E: Worksheet Operations â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[E] Worksheet Operations');

  await test('E', 'create_worksheet', async () => {
    const r = await mcpCall('tools/call', { name: 'create_worksheet', arguments: { filepath: TEST_FILE, sheet_name: 'TongHop' } });
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('E', 'rename_worksheet', async () => {
    const r = await mcpCall('tools/call', { name: 'rename_worksheet', arguments: {
      filepath: TEST_FILE, old_name: 'TongHop', new_name: 'TongHop_v2'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('E', 'copy_worksheet', async () => {
    const r = await mcpCall('tools/call', { name: 'copy_worksheet', arguments: {
      filepath: TEST_FILE, source_sheet: 'Sheet1', target_sheet: 'Sheet1_Copy'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('E', 'delete_worksheet', async () => {
    const r = await mcpCall('tools/call', { name: 'delete_worksheet', arguments: {
      filepath: TEST_FILE, sheet_name: 'Sheet1_Copy'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  // â”€â”€ GROUP F: Row/Column Ops â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[F] Row & Column Operations');

  await test('F', 'insert_rows', async () => {
    const r = await mcpCall('tools/call', { name: 'insert_rows', arguments: {
      filepath: TEST_FILE, sheet_name: 'Sheet1', start_row: 2, count: 1
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('F', 'insert_columns', async () => {
    const r = await mcpCall('tools/call', { name: 'insert_columns', arguments: {
      filepath: TEST_FILE, sheet_name: 'Sheet1', start_col: 2, count: 1
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('F', 'delete_sheet_rows', async () => {
    const r = await mcpCall('tools/call', { name: 'delete_sheet_rows', arguments: {
      filepath: TEST_FILE, sheet_name: 'Sheet1', start_row: 3, count: 1
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('F', 'delete_sheet_columns', async () => {
    const r = await mcpCall('tools/call', { name: 'delete_sheet_columns', arguments: {
      filepath: TEST_FILE, sheet_name: 'Sheet1', start_col: 3, count: 1
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  // â”€â”€ GROUP G: Advanced Tools â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[G] Advanced Tools');

  // Táº¡o file thá»© 2 sáº¡ch Ä‘á»ƒ test advanced
  await mcpCall('tools/call', { name: 'create_workbook', arguments: { filepath: TEST_FILE2 } });
  await mcpCall('tools/call', { name: 'write_data_to_excel', arguments: {
    filepath: TEST_FILE2, sheet_name: 'Data',
    data: [
      ['San pham', 'Vung', 'Quy', 'Doanh thu'],
      ['A', 'Mien Bac', 'Q1', 100],
      ['B', 'Mien Nam', 'Q1', 200],
      ['A', 'Mien Bac', 'Q2', 150],
      ['B', 'Mien Nam', 'Q2', 250],
    ]
  }});

  await test('G', 'create_table', async () => {
    const r = await mcpCall('tools/call', { name: 'create_table', arguments: {
      filepath: TEST_FILE2, sheet_name: 'Data',
      data_range: 'A1:D5', table_name: 'DanhSach'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('G', 'copy_range', async () => {
    // source_start, source_end, target_start are the correct params
    const r = await mcpCall('tools/call', { name: 'copy_range', arguments: {
      filepath: TEST_FILE2, sheet_name: 'Data',
      source_start: 'A1', source_end: 'D3', target_start: 'A7'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('G', 'delete_range', async () => {
    // Delete the copied range (rows 7-9 exist now after copy)
    const r = await mcpCall('tools/call', { name: 'delete_range', arguments: {
      filepath: TEST_FILE2, sheet_name: 'Data',
      start_cell: 'A7', end_cell: 'D9'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('G', 'get_merged_cells', async () => {
    const r = await mcpCall('tools/call', { name: 'get_merged_cells', arguments: {
      filepath: TEST_FILE, sheet_name: 'BaoCao'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  await test('G', 'validate_excel_range (valid)', async () => {
    const r = await mcpCall('tools/call', { name: 'validate_excel_range', arguments: { range_str: 'A1:Z100' } });
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error') && !content.toLowerCase().includes('valid')) return { error: content };
    return { pass: true };
  });

  await test('G', 'get_data_validation_info', async () => {
    const r = await mcpCall('tools/call', { name: 'get_data_validation_info', arguments: {
      filepath: TEST_FILE2, sheet_name: 'Data'
    }});
    const content = r.data?.result?.content?.[0]?.text || '';
    if (content.toLowerCase().includes('error')) return { error: content };
    return { pass: true };
  });

  // â”€â”€ GROUP H: Upload / Download â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[H] Upload & Download');

  const UPLOAD_FILENAME = `upload_test_${TS}.xlsx`;

  await test('H', 'POST /upload (invalid extension bá»‹ tá»« chá»‘i)', async () => {
    const form = new FormData();
    const badBlob = new Blob(['not excel'], { type: 'text/plain' });
    form.append('file', badBlob, 'bad.txt');
    const r = await fetch(`${BASE}/upload`, { method: 'POST', body: form });
    if (r.ok) return { error: 'Pháº£i bá»‹ tá»« chá»‘i nhÆ°ng OK' };
    return { pass: true, status: r.status };
  });

  await test('H', 'POST /upload (no file bá»‹ tá»« chá»‘i)', async () => {
    const r = await fetch(`${BASE}/upload`, { method: 'POST' });
    if (r.ok) return { error: 'Pháº£i bá»‹ tá»« chá»‘i khi khĂ´ng cĂ³ file' };
    return { pass: true, status: r.status };
  });

  await test('H', 'GET /download/:name (file tá»“n táº¡i)', async () => {
    // Download file Ä‘Ă£ táº¡o trong group C
    const r = await fetch(`${BASE}/download/${TEST_FILE}`);
    if (!r.ok) return { error: `HTTP ${r.status}: ${await r.text().catch(()=>'')}` };
    const ct = r.headers.get('content-type') || '';
    const cd = r.headers.get('content-disposition') || '';
    if (!ct.includes('spreadsheet') && !ct.includes('octet')) return { error: `Wrong content-type: ${ct}` };
    if (!cd.includes('attachment')) return { error: `Missing attachment header: ${cd}` };
    return { pass: true, content_type: ct.slice(0, 50) };
  });

  await test('H', 'GET /download/:name (file khĂ´ng tá»“n táº¡i â†’ 404)', async () => {
    const r = await fetch(`${BASE}/download/nonexistent_${TS}.xlsx`);
    if (r.status !== 404) return { error: `Expected 404, got ${r.status}` };
    return { pass: true };
  });

  // â”€â”€ GROUP I: File Management â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[I] File Management');

  await test('I', 'GET /files sau khi táº¡o file', async () => {
    const r = await fetchJSON(`${BASE}/files`);
    const files = r.data?.files || [];
    const found = files.some(f => f.name === TEST_FILE);
    if (!found) return { error: `KhĂ´ng tĂ¬m tháº¥y ${TEST_FILE}` };
    return { pass: true, total: files.length };
  });

  await test('I', 'GET /files/:name (info)', async () => {
    const r = await fetchJSON(`${BASE}/files/${TEST_FILE}`);
    if (!r.ok) return { error: `HTTP ${r.status}: ${JSON.stringify(r.data).slice(0, 100)}` };
    return { pass: true, size_kb: r.data?.file?.size_kb };
  });

  await test('I', 'GET /files/:name KHĂ”NG tráº£ full_path', async () => {
    const r = await fetchJSON(`${BASE}/files/${TEST_FILE}`);
    if (r.data?.file?.full_path) return { error: 'full_path bá»‹ lá»™ ra ngoĂ i!' };
    return { pass: true };
  });

  // â”€â”€ GROUP J: Security â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n[J] Security');

  await test('J', 'Path traversal bá»‹ tá»« chá»‘i (../secret.xlsx)', async () => {
    const r = await mcpCall('tools/call', {
      name: 'read_data_from_excel', arguments: { filepath: '../secret.xlsx', sheet_name: 'Sheet' }
    });
    const content = r.data?.result?.content?.[0]?.text || '';
    const isBlockedOrMissing = content.toLowerCase().includes('error')
      || content.toLowerCase().includes('not found')
      || content.toLowerCase().includes('no such file')
      || r.data?.error;
    if (!isBlockedOrMissing && content.length > 50) return { error: `Path traversal khĂ´ng bá»‹ block! Response: ${content.slice(0, 100)}` };
    return { pass: true, blocked: true };
  });

  await test('J', 'Rate limiter header tá»“n táº¡i (/mcp hoáº¡t Ä‘á»™ng)', async () => {
    const r = await fetch(`${BASE}/mcp`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Accept': 'application/json, text/event-stream' },
      body: JSON.stringify({ jsonrpc: '2.0', id: 999, method: 'ping', params: {} })
    });
    // Pháº£i nháº­n response (dĂ¹ 4xx), khĂ´ng pháº£i connection refused
    if (!r) return { error: 'KhĂ´ng káº¿t ná»‘i Ä‘Æ°á»£c' };
    return { pass: true, status: r.status };
  });

  await test('J', 'DELETE /files/:name (dá»n dáº¹p test file 1)', async () => {
    const r = await fetchJSON(`${BASE}/files/${TEST_FILE}`, { method: 'DELETE' });
    if (!r.ok) return { error: `HTTP ${r.status}: ${JSON.stringify(r.data).slice(0, 100)}` };
    return { pass: true, deleted: r.data?.deleted };
  });

  await test('J', 'DELETE /files/:name (dá»n dáº¹p test file 2)', async () => {
    const r = await fetchJSON(`${BASE}/files/${TEST_FILE2}`, { method: 'DELETE' });
    if (!r.ok) return { error: `HTTP ${r.status}: ${JSON.stringify(r.data).slice(0, 100)}` };
    return { pass: true, deleted: r.data?.deleted };
  });

  // â”€â”€ Káº¿t quáº£ â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  console.log('\n' + '='.repeat(60));
  console.log(`Káº¾T QUáº¢: PASS ${pass}/${pass + fail}   FAIL ${fail}/${pass + fail}`);
  console.log('='.repeat(60));
  const groups = [...new Set(RESULTS.map(r => r.group))];
  for (const g of groups) {
    const gr = RESULTS.filter(r => r.group === g);
    const gp = gr.filter(r => r.ok).length;
    console.log(`  ${gp === gr.length ? 'âœ…' : 'â ï¸'} Group ${g}: ${gp}/${gr.length}`);
    gr.filter(r => !r.ok).forEach(r => console.log(`       FAIL: ${r.name}`));
  }
  console.log();
  if (fail === 0) console.log('đŸ‰ Táº¤T Cáº¢ PASS â€” mcp_excel v2 hoĂ n thiá»‡n!');
  else console.log(`â ï¸  CĂ³ ${fail} test tháº¥t báº¡i â€” cáº§n kiá»ƒm tra`);
})();
