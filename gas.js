// Google Apps Script — Quản lý Camera Trụ CSCC
// Deploy: Extensions → Apps Script → Deploy as Web App
// Execute as: Me | Who has access: Anyone

const SHEET_NAME  = 'DanhSachCamera'; // ← Tab chứa dữ liệu camera (đổi cho đúng tên tab)
const USERS_SHEET = 'TaiKhoan';       // ← Tab tài khoản: tenDangNhap | matKhau | hoTen | vaiTro

// Map camelCase JS → tên cột Sheet chính xác (phải khớp với header row trong Sheet)
const FIELD_MAP = {
  'id':              'STT',
  'lat':             'Lat',
  'latitude':        'Lat',
  'lon':             'Long',
  'longitude':       'Long',
  'lontitude':       'Long',
  'duong':           'Tuyến đường',
  'tenTu':           'Tên tủ điều khiển',
  'tongSoTru':       'Tổng số trụ có camera đang lắp trên trụ',
  'soTru':           'Vị trí trụ (Số trụ CS)',
  'loaiTru':         'Loại trụ',
  'soNha':           'Địa chỉ: Số nhà (nếu có)',
  'phuong':          'Địa chỉ: Phường (Xã)',
  'chuDauTu':        'Chủ đầu tư',
  'ngayPhatHien':    'Ngày phát hiện',
  'bienBan':         'Biên bản báo cáo sự cố/Biên bản ghi nhận hiện trường (có/không)',
  'anToanCachLy':    'Mức độ an toàn: Camera đã được cách ly với trụ đèn',
  'anToanCapNguon':  'Mức độ an toàn: Cáp nguồn 2 lớp vỏ bọc',
  'anToanMoiNoi':    'Mức độ an toàn: Có mối nối hở',
  'anToanKepTreo':   'Mức độ an toàn: Có kẹp treo, kẹp dừng cố định cáp',
  'nguyenNhanKhac':  'Các nguyên nhân gây mất an toàn khác',
  'hinhAnh':         'Hình ảnh',
  'ghiChu':          'Ghi chú',
};

// ── UTILS ──────────────────────────────────────────────────────────────────

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function getUsersSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
}

// Chuẩn hóa chuỗi: NFC + chuẩn hóa khoảng trắng (bao gồm xuống dòng) + trim + lowercase
function norm(s) {
  return String(s || '').normalize('NFC').replace(/\s+/g, ' ').trim().toLowerCase();
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
      message: 'Upload ảnh camera CSCC: ' + fileName,
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

    // Upload ảnh camera (tự tạo tên file)
    if (data.action === 'upload_image') {
      return handleImageUpload(data.imageBase64 || '', data.soTru || '', data.ext || 'jpg');
    }

    // Ghi file tuỳ chỉnh lên GitHub
    if (data.action === 'github_write_file') {
      return handleGithubWriteFile(data.path || '', data.content || '', data.sha || '', data.message || '');
    }

    // Ghi dữ liệu camera vào Sheet
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
        if (latVal !== '') updates['Lat']  = latVal;
        if (lonVal !== '') updates['Long'] = lonVal;
        updateRowFields(sheet, hIdx, rowNum, updates);
      }
    }

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

// ── MARKER CRUD ────────────────────────────────────────────────────────────

// Tìm số hàng theo STT (ưu tiên) hoặc Số trụ CS
function findRowNum(sheet, headers, hIdx, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const sttColIdx   = hIdx[norm('STT')];
  const soTruColIdx = hIdx[norm('Vị trí trụ (Số trụ CS)')];

  const searchId    = norm(data.id    || data['STT']                      || '');
  const searchSoTru = norm(data.soTru || data['Vị trí trụ (Số trụ CS)'] || '');
  if (!searchId && !searchSoTru) return -1;

  const allData = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  for (let i = 0; i < allData.length; i++) {
    const rowId    = sttColIdx   !== undefined ? norm(allData[i][sttColIdx])   : '';
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
  // Build normalized lookup: norm(fieldKey) → value
  const normFV = {};
  for (const [k, v] of Object.entries(fieldValues)) normFV[norm(k)] = v;
  const row = headers.map(h => {
    const v = normFV[norm(h)];
    return v !== undefined ? v : '';
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
  return jsonResponse({ status: 'ok', message: 'Camera CSCC GAS v1 — login ready' });
}
