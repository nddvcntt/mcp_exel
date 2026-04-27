/**
 * file-manager.js
 * Quản lý thư mục excel_files/ và output/: liệt kê, info, upload, download, xóa, auto-cleanup.
 *
 * Cải tiến v2:
 *  - uploadFile(filename, buffer): lưu file Excel từ buffer vào excel_files/
 *  - downloadFile(filename): trả về path an toàn để stream về client
 *  - Hỗ trợ .csv trong ALLOWED_EXTENSIONS (read-only, không upload)
 *  - getFilePath(): expose safe path cho stream download
 */
'use strict';
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const fs   = require('fs');
const path = require('path');

const EXCEL_FILES_PATH   = process.env.EXCEL_FILES_PATH || path.join(__dirname, '..', 'excel_files');
const OUTPUT_DIR         = process.env.OUTPUT_DIR       || path.join(__dirname, '..', 'output');
// [MCP Contract §4] TTL 5 ngày (thống nhất với mcp_powerpoint)
const MAX_FILE_AGE_DAYS  = parseInt(process.env.FILE_TTL_DAYS || '5', 10);
const MAX_UPLOAD_BYTES   = parseInt(process.env.MAX_FILE_SIZE || '52428800', 10); // 50MB

const ALLOWED_EXTENSIONS        = ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xlam'];
const ALLOWED_READONLY_EXTS     = [...ALLOWED_EXTENSIONS, '.csv'];

// Đảm bảo thư mục tồn tại
fs.mkdirSync(EXCEL_FILES_PATH, { recursive: true });
fs.mkdirSync(OUTPUT_DIR,       { recursive: true });

// ── Path safety ───────────────────────────────────────────

/**
 * Sanitize đường dẫn — ngăn path traversal.
 * [P1] Chỉ chấp nhận tên file đơn giản (không có /\\)
 */
function safePath(baseDir, relativePath) {
  if (!relativePath) throw new Error('Đường dẫn không được để trống');
  if (path.isAbsolute(relativePath)) throw new Error('Chỉ chấp nhận đường dẫn tương đối');
  const base     = path.resolve(baseDir);
  const resolved = path.resolve(base, relativePath);
  const rel      = path.relative(base, resolved);
  if (rel.startsWith('..') || path.isAbsolute(rel)) {
    throw new Error('Path traversal không được phép');
  }
  return resolved;
}

/**
 * Kiểm tra extension cho phép (upload)
 */
function validateExtension(filename) {
  const ext = path.extname(filename).toLowerCase();
  if (!ALLOWED_EXTENSIONS.includes(ext)) {
    throw new Error(`Extension không được hỗ trợ: ${ext}. Chỉ chấp nhận: ${ALLOWED_EXTENSIONS.join(', ')}`);
  }
  return ext;
}

/**
 * Kiểm tra tên file được quản lý (không có đường dẫn, chỉ tên file)
 */
function validateManagedFilename(filename) {
  if (!filename || /[/\\]/.test(filename) || path.isAbsolute(filename) || path.basename(filename) !== filename) {
    throw new Error('Tên file không hợp lệ (không được có đường dẫn)');
  }
  const ext = path.extname(filename).toLowerCase();
  if (!ALLOWED_READONLY_EXTS.includes(ext)) {
    throw new Error(`Extension không được hỗ trợ: ${ext}`);
  }
}

// ── List files ────────────────────────────────────────────

/**
 * Liệt kê toàn bộ file Excel trong excel_files + output
 */
function listFiles() {
  const result = [];
  for (const dir of [EXCEL_FILES_PATH, OUTPUT_DIR]) {
    const label = dir === OUTPUT_DIR ? 'output' : 'excel_files';
    if (!fs.existsSync(dir)) continue;
    const files = fs.readdirSync(dir, { withFileTypes: true });
    for (const f of files) {
      if (!f.isFile()) continue;
      const ext = path.extname(f.name).toLowerCase();
      if (!ALLOWED_READONLY_EXTS.includes(ext)) continue;
      const fullPath = safePath(dir, f.name);
      const stat     = fs.statSync(fullPath);
      result.push({
        name         : f.name,
        relative_path: f.name,
        // full_path intentionally omitted (MCP Contract §3)
        size_bytes   : stat.size,
        size_kb      : Math.round(stat.size / 1024),
        created      : stat.birthtime.toISOString(),
        modified     : stat.mtime.toISOString(),
        location     : label,
      });
    }
  }
  return result.sort((a, b) => new Date(b.modified) - new Date(a.modified));
}

// ── Get file info ─────────────────────────────────────────

/**
 * Lấy thông tin chi tiết 1 file.
 * [MCP Contract §3] Không trả full_path ra ngoài.
 */
function getFileInfo(filename) {
  validateManagedFilename(filename);
  for (const dir of [EXCEL_FILES_PATH, OUTPUT_DIR]) {
    let fullPath;
    try { fullPath = safePath(dir, filename); } catch { continue; }
    if (fs.existsSync(fullPath)) {
      const stat = fs.statSync(fullPath);
      return {
        name      : filename,
        size_bytes: stat.size,
        size_kb   : Math.round(stat.size / 1024),
        created   : stat.birthtime.toISOString(),
        modified  : stat.mtime.toISOString(),
        location  : dir === OUTPUT_DIR ? 'output' : 'excel_files',
      };
    }
  }
  throw new Error(`File không tồn tại: ${filename}`);
}

// ── Get file path (for download streaming) ────────────────

/**
 * Trả về đường dẫn tuyệt đối an toàn để stream download.
 * Chỉ dùng nội bộ — không expose ra API response.
 */
function getFilePath(filename) {
  validateManagedFilename(filename);
  for (const dir of [EXCEL_FILES_PATH, OUTPUT_DIR]) {
    let fullPath;
    try { fullPath = safePath(dir, filename); } catch { continue; }
    if (fs.existsSync(fullPath)) return fullPath;
  }
  throw new Error(`File không tồn tại: ${filename}`);
}

// ── Upload file ───────────────────────────────────────────

/**
 * Lưu file Excel từ buffer vào excel_files/.
 * @param {string} filename   - tên file (chỉ basename)
 * @param {Buffer} buffer     - nội dung file
 * @param {object} [opts]     - { overwrite: bool }
 */
function uploadFile(filename, buffer, opts = {}) {
  // Validate filename & extension
  if (!filename || /[/\\]/.test(filename) || path.basename(filename) !== filename) {
    throw new Error('Tên file không hợp lệ');
  }
  validateExtension(filename); // chỉ cho phép ALLOWED_EXTENSIONS (không bao gồm .csv)

  // Validate kích thước
  if (buffer.length > MAX_UPLOAD_BYTES) {
    throw new Error(`File quá lớn: ${Math.round(buffer.length / 1024 / 1024)}MB > ${Math.round(MAX_UPLOAD_BYTES / 1024 / 1024)}MB giới hạn`);
  }

  const destPath = safePath(EXCEL_FILES_PATH, filename);

  // Kiểm tra overwrite
  if (!opts.overwrite && fs.existsSync(destPath)) {
    throw new Error(`File đã tồn tại: ${filename}. Dùng overwrite=true để ghi đè.`);
  }

  fs.writeFileSync(destPath, buffer);
  const stat = fs.statSync(destPath);
  return {
    name      : filename,
    size_bytes: stat.size,
    size_kb   : Math.round(stat.size / 1024),
    location  : 'excel_files',
    uploaded  : new Date().toISOString(),
  };
}

// ── Delete file ───────────────────────────────────────────

/**
 * Xóa file.
 * [P1] Dùng safePath() — chặn path traversal.
 */
function deleteFile(filename) {
  validateManagedFilename(filename);
  for (const dir of [EXCEL_FILES_PATH, OUTPUT_DIR]) {
    let fullPath;
    try { fullPath = safePath(dir, filename); } catch { continue; }
    if (fs.existsSync(fullPath)) {
      fs.unlinkSync(fullPath);
      return { deleted: filename, from: dir === OUTPUT_DIR ? 'output' : 'excel_files' };
    }
  }
  throw new Error(`File không tồn tại: ${filename}`);
}

// ── Auto-cleanup ──────────────────────────────────────────

/**
 * Auto-cleanup: xóa file cũ hơn MAX_FILE_AGE_DAYS ngày trong output/
 */
function autoCleanup() {
  if (!fs.existsSync(OUTPUT_DIR)) return;
  const cutoff = Date.now() - MAX_FILE_AGE_DAYS * 24 * 60 * 60 * 1000;
  const files  = fs.readdirSync(OUTPUT_DIR, { withFileTypes: true });
  let deleted  = 0;
  for (const f of files) {
    if (!f.isFile()) continue;
    if (!ALLOWED_READONLY_EXTS.includes(path.extname(f.name).toLowerCase())) continue;
    const fullPath = safePath(OUTPUT_DIR, f.name);
    const stat     = fs.statSync(fullPath);
    if (stat.mtime.getTime() < cutoff) {
      fs.unlinkSync(fullPath);
      deleted++;
      console.log(`[FileManager] Auto-cleanup: xóa ${f.name}`);
    }
  }
  if (deleted > 0) console.log(`[FileManager] Auto-cleanup: đã xóa ${deleted} file cũ`);
}

// Chạy auto-cleanup mỗi giờ
const cleanupTimer = setInterval(autoCleanup, 60 * 60 * 1000);
if (typeof cleanupTimer.unref === 'function') cleanupTimer.unref();

module.exports = {
  listFiles,
  getFileInfo,
  getFilePath,
  uploadFile,
  deleteFile,
  safePath,
  validateExtension,
  validateManagedFilename,
  EXCEL_FILES_PATH,
  OUTPUT_DIR,
};
