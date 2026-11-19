/**
 * NHÂN VIÊN & MẬT KHẨU – anh Đạt tự quản lý trong app (Điều hành → Nhân viên).
 * Mỗi nhân viên: tên, mã PIN 4 số (không trùng), vai trò, các chức năng được dùng trong Sổ Quán, nghỉ làm.
 * Mật khẩu của anh Đạt: tạo lần đầu bằng mã quản lý (QL_KEY), sau đó đổi trong app.
 */
var SQ_CHUC_NANG = { hd: 'Chụp hóa đơn', dicho: 'Đi chợ', td: 'Hóa đơn quen', tn: 'Thu ngân cuối ca', sapo: 'Gửi số Sapo', huy: 'Hủy hàng', kho: 'Xem kho (cần đặt, tủ bếp, kiểm tủ)', kiemchi: 'Kiểm chi phí ngày (dò hóa đơn, đi chợ, chi tại quầy)' };
var SQ_VAI_MAC_DINH = {
  'ke toan': ['hd', 'dicho', 'td', 'huy', 'kiemchi'], 'di cho': ['dicho', 'hd', 'td', 'huy'], 'thu ngan': ['tn', 'sapo', 'hd', 'td', 'huy'],
  'quan ly': ['tn', 'sapo', 'kho', 'hd', 'dicho', 'td', 'huy'], 'bep': ['huy', 'hd', 'td', 'dicho'], 'phuc vu': ['huy']
};
function sqTachCN_(s, vai) {
  var ds = String(s || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return SQ_CHUC_NANG[x]; });
  ds = ds.length ? ds : (SQ_VAI_MAC_DINH[khongDau_(vai)] || ['hd', 'dicho', 'td', 'huy']).slice();
  if (khongDau_(vai) === 'thu ngan' && ds.indexOf('tn') < 0) ds.unshift('tn'); // thu ngân luôn có ô Thu ngân cuối ca
  if (/^(thu ngan|quan ly)$/.test(khongDau_(vai)) && ds.indexOf('sapo') < 0) ds.push('sapo'); // thu ngân + quản lý gửi số Sapo mỗi tối
  return ds;
}
function sqTaoPass_(p) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CHU_PASS')) throw new Error('Đã có mật khẩu rồi. Muốn đổi thì vào Điều hành → Nhân viên.');
  var ql = props.getProperty('QL_KEY');
  if (!ql || String(p.qlKey || '').trim() !== ql) throw new Error('Mã quản lý không đúng.');
  var moi = String(p.moi || '');
  if (moi.length < 6) throw new Error('Mật khẩu ít nhất 6 ký tự.');
  if (moi === ql) throw new Error('Mật khẩu phải khác mã quản lý.');
  props.setProperty('CHU_PASS', moi);
  return { qlKey: ql };
}
function sqDoiPass_(p) {
  var moi = String(p.moi || ''), props = PropertiesService.getScriptProperties();
  if (moi.length < 6) throw new Error('Mật khẩu ít nhất 6 ký tự.');
  if (moi === props.getProperty('QL_KEY')) throw new Error('Mật khẩu phải khác mã quản lý.');
  props.setProperty('CHU_PASS', moi);
  return {};
}
function sqNgaySinh_(v) { var s = (v instanceof Date) ? chuanNgay_(v) : chuanNgay_(String(v || '').trim()); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
function sqNVDs_() {
  var v = layFile_().getSheetByName(SQ_NV).getDataRange().getValues().slice(1);
  return {
    chucNang: SQ_CHUC_NANG, vaiMacDinh: SQ_VAI_MAC_DINH,
    ds: v.map(function (r, i) { return { row: i + 2, ten: String(r[0]).trim(), pin: String(r[1]).trim(), vai: String(r[2]).trim(), nghi: /x/i.test(String(r[3])), chucNang: sqTachCN_(r[4], r[2]), suaLuc: String(r[5] || ''), sdt: String(r[6] || '').trim(), sinh: sqNgaySinh_(r[7]) }; })
      .filter(function (x) { return x.ten; })
  };
}
function sqNVLuu_(p) {
  var ten = chu_(p.ten, 40), pin = String(p.pin || '').trim(), vai = chu_(p.vai, 20);
  if (!ten) throw new Error('Chưa ghi tên.');
  if (!/^\d{4}$/.test(pin)) throw new Error('Mã PIN phải đúng 4 số.');
  var cn = (p.chucNang || []).filter(function (x) { return SQ_CHUC_NANG[x]; });
  if (!cn.length) throw new Error('Chọn ít nhất 1 chức năng cho ' + ten + '.');
  var sdt = String(p.sdt || '').replace(/[^\d+ .\-]/g, '').trim().slice(0, 16), sinh = String(p.sinh || '').trim();
  if (sinh && !/^\d{4}-\d{2}-\d{2}$/.test(sinh)) throw new Error('Ngày sinh chưa đúng (cần chọn đủ ngày/tháng/năm).');
  return khoa_(function () {
    var sh = layFile_().getSheetByName(SQ_NV), ds = sqNVDs_().ds, row = so_(p.row);
    ds.forEach(function (x) {
      if (x.row === row) return;
      if (x.pin === pin && !x.nghi) throw new Error('Mã PIN ' + pin + ' đang là của ' + x.ten + ', chọn số khác.');
      if (khongDau_(x.ten) === khongDau_(ten) && !x.nghi) throw new Error('Đã có nhân viên tên ' + x.ten + '.');
    });
    if (sh.getRange(1, 7).getValue() !== 'Số điện thoại') { sh.getRange(1, 1, 1, H_SQ_NV.length).setValues([H_SQ_NV]); sh.getRange('G:H').setNumberFormat('@'); }
    var vals = [[ten, pin, vai, p.nghi ? 'x' : '', cn.join(','), bayGio_(), sdt, sinh]];
    if (row >= 2 && ds.some(function (x) { return x.row === row; })) sh.getRange(row, 1, 1, H_SQ_NV.length).setValues(vals);
    else { row = sh.getLastRow() + 1; sh.getRange(row, 1, 1, H_SQ_NV.length).setValues(vals); }
    sqCacheXoa_('sq_nv');
    return { row: row };
  });
}

// sinh nhật nhân viên: báo ở Tổng quan từ 3 ngày trước đến đúng ngày
function sqSinhNhatViec_() {
  var hn = homNay_(), y = +hn.slice(0, 4), out = [];
  sqNVDs_().ds.forEach(function (x) {
    if (x.nghi || !x.sinh) return;
    var m = +x.sinh.slice(5, 7), d = +x.sinh.slice(8, 10), pad = function (n) { return ('0' + n).slice(-2); };
    var ngay = y + '-' + pad(m) + '-' + pad(d), con = sqKhoSoNgay_(hn, ngay);
    if (con < 0) { ngay = (y + 1) + '-' + pad(m) + '-' + pad(d); con = sqKhoSoNgay_(hn, ngay); }
    if (con < 0 || con > 3) return;
    var tuoi = +ngay.slice(0, 4) - (+x.sinh.slice(0, 4)), dm = pad(d) + '/' + pad(m);
    out.push({ muc: 'i', loai: '🎂 Sinh nhật', tab: 'nv', ngay: ngay, con: con, text: (con === 0 ? 'Hôm nay là sinh nhật ' : con === 1 ? 'Ngày mai sinh nhật ' : 'Còn ' + con + ' ngày nữa là sinh nhật ') + x.ten + ' (' + dm + (tuoi > 0 && tuoi < 80 ? ', tròn ' + tuoi + ' tuổi' : '') + ')' + (x.sdt ? ' · ' + x.sdt : '') });
  });
  return out.sort(function (a, b) { return a.con - b.con; });
}

/* ================= BẢNG LƯƠNG (chuyển từ trang Phiếu lương Gặm vào Điều hành) =================
 * Mỗi "tài liệu" 1 dòng ở tab SQ_Luong: Đường dẫn | JSON | Sửa lúc
 *   months/2026-10            → tháng (doanh thu ngày, ngày lễ, mốc thưởng nóng…)
 *   staff-2026-10/<mã-nv>     → 1 nhân viên trong tháng (lương/giờ, ca làm, ứng…)
 * Nút "Lưu tổng vào Google Sheets" ghi file BẢNG_LƯƠNG qua Sheets API (cần file được chia sẻ cho tài khoản chạy Apps Script). */
var SQ_LUONG = 'SQ_Luong', H_SQ_LUONG = ['Đường dẫn', 'JSON', 'Sửa lúc'];
var LUONG_SHEET_ID = '1JQGGRceNehTqxxIZNy8-MhVCby4It407yCRNaYS2mXM';
function sqLuongTab_() { return sqTab_(SQ_LUONG, H_SQ_LUONG, function (sh) { sh.getRange('A:C').setNumberFormat('@'); }); }
function sqLuongDuong_(p) { p = String(p || ''); if (!/^(months\/\d{4}-\d{2}|staff-\d{4}-\d{2}\/[a-z0-9-]{1,60})$/.test(p)) throw new Error('Đường dẫn bảng lương không hợp lệ: ' + p); return p; }
function sqLuongDs_() {
  var sh = sqLuongTab_(), n = sh.getLastRow(), docs = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { var k = String(r[0] || ''); if (k) { var o = sqJ_(String(r[1] || ''), null); if (o) docs[k] = o; } });
  return { docs: docs };
}
function sqLuongGhi_(docs) {
  var keys = Object.keys(docs || {}); if (!keys.length) return { ok: true, n: 0 };
  keys.forEach(sqLuongDuong_);
  return khoa_(function () {
    var sh = sqLuongTab_(), n = sh.getLastRow(), cot = n >= 2 ? sh.getRange(2, 1, n - 1, 1).getValues() : [], hang = {}, luc = bayGio_(), moi = [];
    cot.forEach(function (r, i) { if (r[0]) hang[String(r[0])] = i + 2; });
    keys.forEach(function (k) {
      var s = JSON.stringify(docs[k]); if (s.length > 48000) throw new Error('Dữ liệu ' + k + ' quá lớn.');
      if (hang[k]) sh.getRange(hang[k], 2, 1, 2).setValues([[s, luc]]); else moi.push([k, s, luc]);
    });
    if (moi.length) sh.getRange(sh.getLastRow() + 1, 1, moi.length, 3).setValues(moi);
    return { ok: true, n: keys.length };
  });
}
function sqLuongLuu_(p) { var d = {}; d[sqLuongDuong_(p.path)] = p.body || {}; return sqLuongGhi_(d); }
// Google Sheets cho nút "Lưu tổng": cùng tên công cụ như trang cũ. Dùng dịch vụ nâng cao "Sheets" (Google Sheets API) nếu đã bật, không thì gọi thẳng API.
function sqLuongGS_(p) {
  var id = LUONG_SHEET_ID, inp = p.input || {}, t = String(p.tool || '');
  if (inp.spreadsheetId && inp.spreadsheetId !== id) throw new Error('Chỉ ghi được file BẢNG_LƯƠNG.');
  var loi = function (e) { var m = String(e && e.message || e); if (/permission|PERMISSION_DENIED|403|404|not found|caller does not have/i.test(m)) return new Error('Tài khoản chạy app (tiemangam2015) chưa sửa được file BẢNG_LƯƠNG: ' + m.slice(0, 160)); return new Error('Google Sheets báo lỗi: ' + m.slice(0, 200)); };
  var sach = function (o) { return JSON.parse(JSON.stringify(o || {})); };
  if (typeof Sheets !== 'undefined') {
    try {
      if (t === 'get_spreadsheet') return { payload: sach(Sheets.Spreadsheets.get(id, { fields: (inp.fields || []).join(',') })) };
      if (t === 'update_spreadsheet') return { payload: sach(Sheets.Spreadsheets.batchUpdate({ requests: inp.requests || [] }, id)) };
      if (t === 'batch_clear_values') return { payload: sach(Sheets.Spreadsheets.Values.batchClear({ ranges: inp.ranges || [] }, id)) };
      if (t === 'update_formulas') return { payload: sach(Sheets.Spreadsheets.Values.update({ range: inp.range, majorDimension: 'ROWS', values: inp.formulas || [] }, id, inp.range, { valueInputOption: 'USER_ENTERED' })) };
    } catch (e) { throw loi(e); }
    throw new Error('Không rõ thao tác Google Sheets: ' + t);
  }
  var base = 'https://sheets.googleapis.com/v4/spreadsheets/' + id;
  var goi = function (method, url, body) {
    var o = { method: method, muteHttpExceptions: true, contentType: 'application/json', headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } };
    if (body) o.payload = JSON.stringify(body);
    var r = UrlFetchApp.fetch(url, o), c = r.getResponseCode(), tx = r.getContentText();
    if (c >= 300) throw loi(new Error(c + ' ' + tx.slice(0, 300)));
    return tx ? JSON.parse(tx) : {};
  };
  if (t === 'get_spreadsheet') return { payload: goi('get', base + '?fields=' + encodeURIComponent((inp.fields || []).join(','))) };
  if (t === 'update_spreadsheet') return { payload: goi('post', base + ':batchUpdate', { requests: inp.requests || [] }) };
  if (t === 'batch_clear_values') return { payload: goi('post', base + '/values:batchClear', { ranges: inp.ranges || [] }) };
  if (t === 'update_formulas') return { payload: goi('put', base + '/values/' + encodeURIComponent(inp.range) + '?valueInputOption=USER_ENTERED', { range: inp.range, majorDimension: 'ROWS', values: inp.formulas || [] }) };
  throw new Error('Không rõ thao tác Google Sheets: ' + t);
}
