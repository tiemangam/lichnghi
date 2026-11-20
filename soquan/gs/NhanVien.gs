/**
 * NHÂN VIÊN & MẬT KHẨU – anh Đạt tự quản lý trong app (Điều hành → Nhân viên).
 * Mỗi nhân viên: tên, mã PIN 4 số (không trùng), vai trò, các chức năng được dùng trong Sổ Quán, nghỉ làm.
 * Mật khẩu của anh Đạt: tạo lần đầu bằng mã quản lý (QL_KEY), sau đó đổi trong app.
 */
var SQ_CHUC_NANG = { hd: 'Chụp hóa đơn', dicho: 'Đi chợ', td: 'Hóa đơn quen', tn: 'Thu ngân cuối ca', sapo: 'Gửi số Sapo', huy: 'Hủy hàng', kho: 'Xem kho (cần đặt, tủ bếp, kiểm tủ)', kiemchi: 'Kiểm chi phí ngày (dò hóa đơn, đi chợ, chi tại quầy)', dodung: 'Xuất đồ dùng (cấp / thu áo, tạp dề…)' };
var SQ_VAI_MAC_DINH = {
  'ke toan': ['hd', 'dicho', 'td', 'huy', 'kiemchi'], 'di cho': ['dicho', 'hd', 'td', 'huy'], 'thu ngan': ['tn', 'sapo', 'hd', 'td', 'huy'],
  'quan ly': ['tn', 'sapo', 'kho', 'hd', 'dicho', 'td', 'huy', 'dodung'], 'bep': ['huy', 'hd', 'td', 'dicho'], 'phuc vu': ['huy']
};
function sqTachCN_(s, vai) {
  var ds = String(s || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return SQ_CHUC_NANG[x]; });
  ds = ds.length ? ds : (SQ_VAI_MAC_DINH[khongDau_(vai)] || ['hd', 'dicho', 'td', 'huy']).slice();
  if (khongDau_(vai) === 'thu ngan' && ds.indexOf('tn') < 0) ds.unshift('tn'); // thu ngân luôn có ô Thu ngân cuối ca
  if (/^(thu ngan|quan ly)$/.test(khongDau_(vai)) && ds.indexOf('sapo') < 0) ds.push('sapo'); // thu ngân + quản lý gửi số Sapo mỗi tối
  if (khongDau_(vai) === 'quan ly' && ds.indexOf('dodung') < 0) ds.push('dodung'); // quản lý luôn xuất đồ dùng được
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

/* ================= DỤNG CỤ (chén, ly, đũa, đồng phục…) =================
 * SQ_DungCu   : 1 dòng / món – ID | Nhóm | Tên | Đơn vị | Giá nhập | Tồn đầu | Thứ tự | Ẩn | Sửa lúc
 * SQ_DungCuNK : nhật ký – ID | Ngày | Mã món | Tên món | Loại | SL | Ai | Tên người | Lý do | Ảnh | Người ghi | Tạo lúc | Sheet | Thêm
 *   Loại: nhap (nhập thêm) · cap (cấp cho nhân viên) · thu (thu lại → kho áo cũ, Thêm = mã món nhận) · huy_dung (hủy đồ nhân viên đang dùng) · be (bể / hư trong kho)
 * Tồn = tồn đầu + nhập + thu vào − cấp − bể. Đồ đang dùng của 1 người = cấp − thu − hủy.
 * Mỗi lần ghi, app ghi thêm 1 dòng vào tab "Nhật ký" của file Sheet "Kiểm tra dụng cụ" (công thức trong tab Tổng tự tính). */
var SQ_DC = 'SQ_DungCu', H_SQ_DC = ['ID', 'Nhóm', 'Tên', 'Đơn vị', 'Giá nhập', 'Tồn đầu', 'Thứ tự', 'Ẩn', 'Sửa lúc'];
var SQ_DCNK = 'SQ_DungCuNK', H_SQ_DCNK = ['ID', 'Ngày', 'Mã món', 'Tên món', 'Loại', 'SL', 'Ai', 'Tên người', 'Lý do', 'Ảnh', 'Người ghi', 'Tạo lúc', 'Sheet', 'Thêm'];
var DC_SHEET_ID = '1R2G6qoymY1VpmsD1qINju27qMRPJJdREgKm_lDw31zI', DC_LOAI = { nhap: 1, cap: 1, thu: 1, huy_dung: 1, be: 1 };
function sqDcTab_() { return sqTab_(SQ_DC, H_SQ_DC, function (sh) { sh.getRange('A:D').setNumberFormat('@'); sh.getRange('I:I').setNumberFormat('@'); }); }
function sqDcNKTab_() { return sqTab_(SQ_DCNK, H_SQ_DCNK, function (sh) { sh.getRange('A:E').setNumberFormat('@'); sh.getRange('G:N').setNumberFormat('@'); }); }
function sqDcId_(t) { return khongDau_(t).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 50) || ('mon-' + Date.now().toString(36)); }
function sqDcMon_() {
  var sh = sqDcTab_(), n = sh.getLastRow(); if (n < 2) return [];
  var s = function (v) { return v === '' || v === null ? '' : Number(v); };
  return sh.getRange(2, 1, n - 1, H_SQ_DC.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0] || ''), nhom: String(r[1] || ''), ten: String(r[2] || ''), dv: String(r[3] || ''), gia: s(r[4]), tonDau: s(r[5]) === '' ? 0 : s(r[5]), thuTu: s(r[6]) === '' ? 999 : s(r[6]), an: /x/i.test(String(r[7] || '')) };
  }).filter(function (x) { return x.id && x.ten; });
}
function sqDcNK_() {
  var sh = sqDcNKTab_(), n = sh.getLastRow(); if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_SQ_DCNK.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0] || ''), ngay: chuanNgay_(r[1]), mon: String(r[2] || ''), tenMon: String(r[3] || ''), loai: String(r[4] || ''), sl: so_(r[5]), ai: String(r[6] || ''), nguoi: String(r[7] || ''), ly: String(r[8] || ''),
      anh: String(r[9] || '').split(',').filter(String), ghi: String(r[10] || ''), tao: String(r[11] || ''), sheet: String(r[12] || ''), vao: String(r[13] || '') };
  }).filter(function (x) { return x.id && DC_LOAI[x.loai]; });
}
function sqDcTinh_(mon, nk) {
  var ton = {}, dung = {};
  mon.forEach(function (m) { ton[m.id] = m.tonDau || 0; });
  nk.forEach(function (e) {
    if (!(e.mon in ton)) ton[e.mon] = 0;
    var k = e.nguoi + '|' + e.mon;
    if (e.loai === 'nhap') ton[e.mon] += e.sl;
    else if (e.loai === 'be') ton[e.mon] -= e.sl;
    else if (e.loai === 'cap') { ton[e.mon] -= e.sl; dung[k] = (dung[k] || 0) + e.sl; }
    else if (e.loai === 'thu') { dung[k] = (dung[k] || 0) - e.sl; var v = e.vao || e.mon; ton[v] = (ton[v] || 0) + e.sl; }
    else if (e.loai === 'huy_dung') dung[k] = (dung[k] || 0) - e.sl;
  });
  var nguoi = [];
  Object.keys(dung).forEach(function (k) { if (dung[k] > 0) { var p = k.split('|'); nguoi.push({ ten: p[0], mon: p[1], sl: dung[k] }); } });
  return { ton: ton, dung: nguoi };
}
function sqDcDs_() {
  var mon = sqDcMon_(), nk = sqDcNK_();
  var t = sqDcTinh_(mon, nk);
  return { mon: mon, nk: nk.map(function (e) { var o = {}; for (var k in e) if (k !== 'row') o[k] = e[k]; return o; }), ton: t.ton, dung: t.dung, chuaDongBo: nk.filter(function (e) { return e.sheet !== 'ok'; }).length, sheetLoi: PropertiesService.getScriptProperties().getProperty('DC_SHEET_LOI') || '' };
}
// thêm / sửa món (anh Đạt): ds = [{id?, nhom, ten, dv, gia, tonDau, thuTu, an}]
function sqDcMonLuu_(p) {
  var ds = (p.ds || []).slice(0, 200);
  return khoa_(function () {
    var sh = sqDcTab_(), cu = sqDcMon_(), byId = {}, luc = bayGio_(), moi = [];
    cu.forEach(function (m) { byId[m.id] = m; });
    ds.forEach(function (x) {
      var ten = chu_(x.ten, 80); if (!ten && !x.id) return;
      var id = String(x.id || '') || sqDcId_(ten), m = byId[id] || {};
      var num = function (v, d) { if (v === undefined) return d; v = String(v).replace(/[^\d.-]/g, ''); return v === '' ? '' : Number(v); };
      var row = [id, x.nhom !== undefined ? chu_(x.nhom, 40) : (m.nhom || 'Khác'), ten || m.ten, x.dv !== undefined ? chu_(x.dv, 20) : (m.dv || 'cái'),
        num(x.gia, m.gia === undefined ? '' : m.gia), num(x.tonDau, m.tonDau === undefined ? 0 : m.tonDau), num(x.thuTu, m.thuTu === undefined ? 999 : m.thuTu), x.an !== undefined ? (x.an ? 'x' : '') : (m.an ? 'x' : ''), luc];
      if (m.row) sh.getRange(m.row, 1, 1, H_SQ_DC.length).setValues([row]); else { moi.push(row); byId[id] = { id: id }; }
    });
    if (moi.length) sh.getRange(sh.getLastRow() + 1, 1, moi.length, H_SQ_DC.length).setValues(moi);
    try { sqDcSheetMon_(ds); } catch (e) { }
    return { ok: true };
  });
}
// ghi nhật ký: dong = [{mon, sl}] + loai, ngay, nguoi, ai, ly, anh (b64)
function sqDcGhi_(p, ghiBoi) {
  var loai = String(p.loai || ''); if (!DC_LOAI[loai]) throw new Error('Loại không hợp lệ.');
  var ngay = sqKiemNgay_(p.ngay || homNay_()), nguoi = chu_(p.nguoi, 40), ai = p.ai === 'khach' ? 'Khách' : (loai === 'be' || loai === 'cap' || loai === 'thu' || loai === 'huy_dung' ? 'Nhân viên' : ''), ly = chu_(p.ly, 150);
  var dong = (p.dong || []).slice(0, 30).map(function (d) { return { mon: String(d.mon || ''), sl: Math.round(so_(d.sl) * 100) / 100 }; }).filter(function (d) { return d.mon && d.sl > 0; });
  if (!dong.length) throw new Error('Chưa chọn món / số lượng.');
  if ((loai === 'cap' || loai === 'thu' || loai === 'huy_dung') && !nguoi) throw new Error('Chưa ghi tên nhân viên.');
  if (loai === 'be' && ai === 'Nhân viên' && !nguoi) throw new Error('Chưa ghi tên người làm bể.');
  var canAnh = loai === 'be' || loai === 'huy_dung', anh = (p.anh || []).slice(0, 3);
  if (canAnh && !anh.length && !p.khongAnh) throw new Error('Phải chụp ảnh nhé.');
  return khoa_(function () {
    var mon = sqDcMon_(), byId = {}; mon.forEach(function (m) { byId[m.id] = m; });
    dong.forEach(function (d) { if (!byId[d.mon]) throw new Error('Không có món này trong bảng dụng cụ.'); });
    var id0 = 'DC' + Date.now().toString(36).toUpperCase(), fids = [];
    if (anh.length) { var folder = sqThuMuc_(); anh.forEach(function (a, i) { fids.push(folder.createFile(Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'image/jpeg', id0 + '-' + (i + 1) + '.jpg')).getId()); }); }
    var luc = bayGio_(), rows = [];
    dong.forEach(function (d, i) {
      var vao = '';
      if (loai === 'thu') { // thu lại → kho "áo cũ": tạo món "(cũ)" nếu chưa có
        var g = byId[d.mon], cid = /-cu$/.test(d.mon) ? d.mon : d.mon + '-cu';
        if (!byId[cid]) { sqDcTab_().appendRow([cid, g.nhom, g.ten.replace(/\s*\(mới\)\s*$/i, '') + ' (cũ)', g.dv, g.gia, 0, (g.thuTu || 999) + 0.5, '', luc]); byId[cid] = { id: cid, ten: g.ten + ' (cũ)', nhom: g.nhom }; }
        vao = cid;
      }
      rows.push([id0 + (dong.length > 1 ? '-' + (i + 1) : ''), ngay, d.mon, byId[d.mon].ten, loai, d.sl, ai, nguoi, ly, fids.join(','), chu_(ghiBoi, 40), luc, '', vao]);
    });
    var sh = sqDcNKTab_(); sh.getRange(sh.getLastRow() + 1, 1, rows.length, H_SQ_DCNK.length).setValues(rows);
    var dongBo = ''; try { sqDcDongBo_(); } catch (e) { dongBo = e.message; }
    return { ok: true, id: id0, dongBo: dongBo };
  });
}
function sqDcNKXoa_(p) {
  var id = String(p.id || ''); if (!id) throw new Error('Thiếu mã.');
  return khoa_(function () { var sh = sqDcNKTab_(), r = sqDcNK_().filter(function (e) { return e.id === id; })[0]; if (!r) throw new Error('Không thấy dòng này.'); sh.deleteRow(r.row); return { ok: true }; });
}
// nhập lần đầu từ file Sheet cũ (chỉ khi app chưa có món nào)
function sqDcNhap_(p) {
  return khoa_(function () {
    if (sqDcMon_().length) throw new Error('App đã có dữ liệu dụng cụ rồi.');
    var mon = (p.mon || []).slice(0, 200).map(function (m, i) { return [sqDcId_(m.ten), chu_(m.nhom, 40) || 'Khác', chu_(m.ten, 80), chu_(m.dv, 20) || 'cái', m.gia === '' || m.gia == null ? '' : Number(m.gia), Number(m.tonDau) || 0, i + 1, '', bayGio_()]; });
    if (mon.length) sqDcTab_().getRange(2, 1, mon.length, H_SQ_DC.length).setValues(mon);
    var ids = {}; mon.forEach(function (r) { ids[khongDau_(r[2])] = r[0]; });
    var nk = (p.nk || []).slice(0, 500).map(function (e, i) { var mid = ids[khongDau_(e.tenMon)]; if (!mid) return null;
      return ['DC0-' + (i + 1), sqKiemNgay_(e.ngay), mid, e.tenMon, DC_LOAI[e.loai] ? e.loai : 'cap', Number(e.sl) || 0, e.ai || '', chu_(e.nguoi, 40), chu_(e.ly, 150), '', 'Sheet cũ', bayGio_(), 'ok', '']; }).filter(Boolean);
    if (nk.length) sqDcNKTab_().getRange(2, 1, nk.length, H_SQ_DCNK.length).setValues(nk);
    return { ok: true, mon: mon.length, nk: nk.length };
  });
}
/* ---- ghi sang file Sheet "Kiểm tra dụng cụ" (cần quyền chỉnh sửa cho tài khoản chạy app) ---- */
function sqDcSheet_() { return SpreadsheetApp.openById(DC_SHEET_ID); }
function sqDcTongDong_(ss, ten, dv, tonDau) { // tìm / thêm dòng món trong tab Tổng (giữ công thức)
  var sh = ss.getSheetByName('Tổng'), n = sh.getLastRow(), a = sh.getRange(1, 1, Math.max(n, 1), 1).getValues();
  for (var i = 1; i < a.length; i++) if (khongDau_(a[i][0]) === khongDau_(ten)) return i + 1;
  var r = n + 1, R = 'R[0]C1';
  sh.getRange(r, 1, 1, 3).setValues([[ten, dv || 'cái', tonDau || 0]]);
  sh.getRange(r, 4, 1, 6).setFormulasR1C1([[
    '=IF(RC1="","",SUMIFS(\'Nhật ký\'!C:C,\'Nhật ký\'!B:B,RC1))', '=IF(RC1="","",SUMIFS(\'Nhật ký\'!D:D,\'Nhật ký\'!B:B,RC1))', '=IF(RC1="","",SUMIFS(\'Nhật ký\'!E:E,\'Nhật ký\'!B:B,RC1))',
    '=IF(RC1="","",N(RC3)+RC4-RC5-RC6)', '=IF(RC1="","",SUMIFS(\'Nhật ký\'!E:E,\'Nhật ký\'!B:B,RC1,\'Nhật ký\'!F:F,"Nhân viên"))', '=IF(RC1="","",SUMIFS(\'Nhật ký\'!E:E,\'Nhật ký\'!B:B,RC1,\'Nhật ký\'!F:F,"Khách"))']]);
  return r;
}
function sqDcSheetMon_(ds) { // cập nhật tồn đầu / thêm món mới vào tab Tổng
  if (!ds || !ds.length) return;
  var ss = sqDcSheet_(), sh = ss.getSheetByName('Tổng'), mon = sqDcMon_(), byId = {}; mon.forEach(function (m) { byId[m.id] = m; });
  ds.forEach(function (x) { var m = byId[x.id || sqDcId_(x.ten)]; if (!m) return; var r = sqDcTongDong_(ss, m.ten, m.dv, m.tonDau); if (x.tonDau !== undefined) sh.getRange(r, 3).setValue(m.tonDau); });
}
function sqDcDongBo_() {
  var props = PropertiesService.getScriptProperties();
  var nk = sqDcNK_().filter(function (e) { return e.sheet !== 'ok'; }); if (!nk.length) { props.deleteProperty('DC_SHEET_LOI'); return { n: 0 }; }
  var ss;
  try { ss = sqDcSheet_(); } catch (e) { props.setProperty('DC_SHEET_LOI', 'Chưa có quyền sửa file Sheet "Kiểm tra dụng cụ" – nhờ chủ file chia sẻ quyền chỉnh sửa cho tiemangam2015@gmail.com'); throw new Error(props.getProperty('DC_SHEET_LOI')); }
  var mon = sqDcMon_(), byId = {}; mon.forEach(function (m) { byId[m.id] = m; });
  var sh = ss.getSheetByName('Nhật ký'), out = [];
  nk.forEach(function (e) {
    var m = byId[e.mon] || { ten: e.tenMon, dv: 'cái' }, p = e.ngay.split('-'), d = p[2] + '/' + p[1] + '/' + p[0];
    sqDcTongDong_(ss, m.ten, m.dv, 0);
    if (e.loai === 'nhap') out.push([d, m.ten, e.sl, '', '', '', '', e.ly || 'Nhập thêm']);
    else if (e.loai === 'cap') out.push([d, m.ten, '', e.sl, '', 'Nhân viên', e.nguoi, e.ly || 'Cấp cho ' + e.nguoi]);
    else if (e.loai === 'be') out.push([d, m.ten, '', '', e.sl, e.ai || 'Nhân viên', e.nguoi, e.ly]);
    else if (e.loai === 'thu') { var v = byId[e.vao] || m; sqDcTongDong_(ss, v.ten, v.dv, 0); out.push([d, v.ten, e.sl, '', '', 'Nhân viên', e.nguoi, 'Thu lại của ' + e.nguoi + (e.ly ? ': ' + e.ly : '')]); }
    else if (e.loai === 'huy_dung') out.push([d, m.ten, '', '', '', 'Nhân viên', e.nguoi, 'Hủy đồ đang dùng (' + e.sl + ' ' + (m.dv || 'cái') + ')' + (e.ly ? ': ' + e.ly : '')]);
  });
  if (out.length) {
    var r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, out.length, 8).setValues(out);
    sh.getRange(r0, 9, out.length, 1).setFormulaR1C1('=IF(RC2="","",N(IFERROR(VLOOKUP(RC2,\'Tổng\'!C1:C3,3,FALSE),0))+SUMIFS(R2C3:RC3,R2C2:RC2,RC2)-SUMIFS(R2C4:RC4,R2C2:RC2,RC2)-SUMIFS(R2C5:RC5,R2C2:RC2,RC2))');
  }
  var nks = sqDcNKTab_(); nk.forEach(function (e) { nks.getRange(e.row, 13).setValue('ok'); });
  props.deleteProperty('DC_SHEET_LOI');
  return { n: out.length };
}
// Sổ Quán: danh sách món + tồn + đồ đang dùng (không có giá)
function sqDcNV_(nv) {
  var mon = sqDcMon_().filter(function (m) { return !m.an; }), nk = sqDcNK_(), t = sqDcTinh_(mon, nk);
  return { mon: mon.map(function (m) { return { id: m.id, nhom: m.nhom, ten: m.ten, dv: m.dv, ton: t.ton[m.id] || 0, thuTu: m.thuTu }; }), dung: t.dung,
    nv: sqNVDs_().ds.filter(function (x) { return !x.nghi; }).map(function (x) { return x.ten; }), xuat: (nv.chucNang || []).indexOf('dodung') >= 0 };
}
