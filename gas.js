// Google Apps Script — Quản lý đèn tắt CSCC
// Deploy: Extensions → Apps Script → Deploy as Web App
// Execute as: Me | Who has access: Anyone

const SHEET_NAME  = 'DanhSachDen';  // ← Tab chứa dữ liệu đèn
const USERS_SHEET = 'TaiKhoan';     // ← Tab tài khoản: tenDangNhap | matKhau | hoTen | vaiTro

// Map camelCase JS → tên cột Sheet chính xác
const FIELD_MAP = {
  'id':            'ID',
  'soTru':         'Số trụ',
  'tenTu':         'Tên tủ',
  'lat':           'latitude',
  'latitude':      'latitude',
  'lon':           'lontitude',
  'longitude':     'lontitude',
  'lontitude':     'lontitude',
  'loaiDen':       'Loại đèn',
  'congSuat':      'Công suất',
  'trangThai':     'Trang thai',
  'duong':         'Đường',
  'phuong':        'Phường',
  'ngayPhatHien':  'Ngày phát hiện',
  'nguoiPhatHien': 'Người phát hiện',
  'ngaySua':       'Ngày sửa',
  'nguoiSua':      'Người sửa',
  'vatTuSua':      'Vật tư sửa',
  'hinhAnh':       'HÌnh ảnh',
  'ghiChu':        'Ghi chú',
  'vn2000x':       'VN2000-X',
  'vn2000y':       'VN2000-Y',
};

// ── UTILS ──────────────────────────────────────────────────────────────────

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getUsersSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
}

// Chuẩn hóa chuỗi để so sánh (NFC + trim + lowercase)
function norm(s) {
  return String(s || '').normalize('NFC').trim().toLowerCase();
}

// Xây dựng index: norm(header) → col index (0-based)
function buildHeaderIndex(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[norm(h)] = i; });
  return idx;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LOGIN ──────────────────────────────────────────────────────────────────

function handleLogin(username, password) {
  const sheet = getUsersSheet();
  if (!sheet) {
    return jsonResponse({ status: 'error', message: 'Sheet "TaiKhoan" chưa được tạo. Admin cần tạo tab này trong Google Sheet.' });
  }

  const allData = sheet.getDataRange().getValues();
  if (allData.length < 2) {
    return jsonResponse({ status: 'error', message: 'Chưa có tài khoản nào trong Sheet TaiKhoan.' });
  }

  const headers = allData[0].map(h => norm(String(h)));
  const colUser = headers.indexOf(norm('tenDangNhap'));
  const colPass = headers.indexOf(norm('matKhau'));
  const colName = headers.indexOf(norm('hoTen'));
  const colRole = headers.indexOf(norm('vaiTro'));

  if (colUser === -1 || colPass === -1) {
    return jsonResponse({ status: 'error', message: 'Sheet TaiKhoan thiếu cột "tenDangNhap" hoặc "matKhau".' });
  }

  const usernameNorm = norm(username);
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const rowUser = norm(String(row[colUser] || ''));
    const rowPass = String(row[colPass] || '');
    if (rowUser === usernameNorm && rowPass === String(password)) {
      return jsonResponse({
        status: 'ok',
        user: {
          username:    String(row[colUser]).trim(),
          displayName: colName >= 0 ? String(row[colName] || '').trim() || username : username,
          role:        colRole >= 0 ? norm(String(row[colRole] || ''))             : 'user',
        }
      });
    }
  }

  return jsonResponse({ status: 'error', message: 'Sai tên đăng nhập hoặc mật khẩu.' });
}

// ── GITHUB IMAGE UPLOAD ────────────────────────────────────────────────────

function handleImageUpload(imageBase64, soTru, ext) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) {
    return jsonResponse({ status: 'error', message: 'GITHUB_TOKEN chưa được cài trong Script Properties của GAS.' });
  }

  const ts = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH-mm-ss") + 'Z';
  const safeName = (soTru || 'img').replace(/[^a-zA-Z0-9]/g, '-').toLowerCase();
  const fileName = safeName + '-' + ts + '.' + (ext || 'jpg');
  const filePath = 'images/' + fileName;

  const apiUrl = 'https://api.github.com/repos/neo-era/dentat/contents/'
    + filePath.split('/').map(encodeURIComponent).join('/');

  const res = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: {
      'Authorization': 'token ' + token,
      'Accept': 'application/vnd.github+json',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      message: 'Upload ảnh đèn tắt: ' + fileName,
      content: imageBase64,
      branch: 'main'
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  if (code !== 200 && code !== 201) {
    return jsonResponse({ status: 'error', message: 'GitHub API lỗi ' + code + ': ' + res.getContentText() });
  }
  return jsonResponse({ status: 'ok', path: 'images/' + fileName });
}

// ── GITHUB FILE WRITE (dùng chung cho Excel + ảnh từ index.html) ───────────

function handleGithubWriteFile(filePath, content, sha, message) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) return jsonResponse({ status: 'error', message: 'GITHUB_TOKEN chưa được cài trong Script Properties.' });
  const encodedPath = filePath.split('/').map(encodeURIComponent).join('/');
  const apiUrl = 'https://api.github.com/repos/neo-era/dentat/contents/' + encodedPath;
  const body = { message: message || 'Update file', content: content, branch: 'main' };
  if (sha) body.sha = sha;
  const res = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: { 'Authorization': 'token ' + token, 'Accept': 'application/vnd.github+json', 'Content-Type': 'application/json' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code !== 200 && code !== 201) {
    return jsonResponse({ status: 'error', message: 'GitHub API lỗi ' + code + ': ' + res.getContentText().slice(0, 300) });
  }
  return jsonResponse({ status: 'ok' });
}

// ── MAIN HANDLER ───────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // Login
    if (data.action === 'login') {
      return handleLogin(data.username || '', data.password || '');
    }

    // Upload ảnh đèn tắt (tự tạo tên file)
    if (data.action === 'upload_image') {
      return handleImageUpload(data.imageBase64 || '', data.soTru || '', data.ext || 'jpg');
    }

    // Ghi file tuỳ chỉnh lên GitHub (Excel + ảnh từ index.html)
    if (data.action === 'github_write_file') {
      return handleGithubWriteFile(data.path || '', data.content || '', data.sha || '', data.message || '');
    }

    // Ghi dữ liệu đèn
    const sheet = getSheet();
    if (!sheet) throw new Error('Không tìm thấy sheet: ' + SHEET_NAME);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hIdx = buildHeaderIndex(headers);

    if (data.action === 'full_update') {
      const rowNum = findRowNum(sheet, headers, hIdx, data);
      if (rowNum > 0) {
        updateRow(sheet, hIdx, rowNum, data);
      } else {
        appendRow(sheet, headers, hIdx, data);
      }
    } else {
      // GPS-only: chỉ cập nhật lat/lon
      const rowNum = findRowNum(sheet, headers, hIdx, data);
      if (rowNum > 0) {
        const updates = {};
        const latVal = data.lat || data.latitude || '';
        const lonVal = data.lon || data.longitude || data.lontitude || '';
        if (latVal !== '') updates['latitude']  = latVal;
        if (lonVal !== '') updates['lontitude'] = lonVal;
        updateRowFields(sheet, hIdx, rowNum, updates);
      }
    }

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── MARKER CRUD ────────────────────────────────────────────────────────────

// Tìm số hàng theo ID (ưu tiên) hoặc Số trụ
function findRowNum(sheet, headers, hIdx, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const idColIdx    = hIdx[norm('ID')];
  const soTruColIdx = hIdx[norm('Số trụ')];

  const searchId    = norm(data.id    || data['ID']    || '');
  const searchSoTru = norm(data.soTru || data['Số trụ'] || '');
  if (!searchId && !searchSoTru) return -1;

  const allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = 0; i < allData.length; i++) {
    const rowId    = idColIdx    !== undefined ? norm(allData[i][idColIdx])    : '';
    const rowSoTru = soTruColIdx !== undefined ? norm(allData[i][soTruColIdx]) : '';
    if ((searchId    && rowId    && rowId    === searchId)    ||
        (searchSoTru && rowSoTru && rowSoTru === searchSoTru)) {
      return i + 2; // i+2: bỏ header (hàng 1) + offset 0-based
    }
  }
  return -1;
}

// Cập nhật toàn bộ fields của 1 hàng
function updateRow(sheet, hIdx, rowNum, data) {
  updateRowFields(sheet, hIdx, rowNum, buildFieldValues(data));
}

// Ghi { 'Tên cột Sheet': value } vào đúng cột bằng hIdx (normalized)
function updateRowFields(sheet, hIdx, rowNum, fieldValues) {
  for (const [header, value] of Object.entries(fieldValues)) {
    const col = hIdx[norm(header)];
    if (col !== undefined && value !== null && value !== undefined) {
      sheet.getRange(rowNum, col + 1).setValue(value);
    }
  }
}

// Thêm hàng mới theo đúng thứ tự cột
function appendRow(sheet, headers, hIdx, data) {
  const fieldValues = buildFieldValues(data);
  const row = headers.map(h => {
    return fieldValues[h] !== undefined ? fieldValues[h] : '';
  });
  sheet.appendRow(row);
}

// Chuyển payload camelCase → { 'Tên cột Sheet': giá trị }
function buildFieldValues(data) {
  const result = {};
  for (const [key, value] of Object.entries(data)) {
    if (key === 'action') continue;
    const header = FIELD_MAP[key] || key;
    if (value !== undefined && value !== null && value !== '') {
      result[header] = value;
    }
  }
  return result;
}

// ── HEALTH CHECK ───────────────────────────────────────────────────────────

function doGet(e) {
  return jsonResponse({ status: 'ok', message: 'Den tat GAS v3 — login ready' });
}
