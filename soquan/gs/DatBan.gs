/**
 * ĐẶT BÀN + MANG VỀ – link công khai cho khách (lichnghi/datban/), nhân viên xử lý trong Sổ Quán, anh Đạt chỉnh menu trong Điều hành.
 * Tab SQ_Menu: món bán (nhóm, tên, biến thể + giá) · tab SQ_DatBan: các đơn.
 * Cài đặt (lưu ý, mức hỗ trợ ship, hình menu) ở Script Property DB_CAI.
 */
var SQ_MENU = 'SQ_Menu', H_SQ_MENU = ['Nhóm', 'Món', 'Biến thể (JSON)', 'Bán', 'Thứ tự', 'Mô tả'];
var SQ_DB = 'SQ_DatBan', H_SQ_DB = ['Mã', 'Loại', 'Ngày', 'Giờ', 'Số người', 'Tên', 'SĐT', 'Địa chỉ', 'Dịp', 'Món (JSON)', 'Tạm tính', 'Hỗ trợ ship', 'Phí ship', 'Ghi chú', 'Trạng thái', 'Bàn', 'Người xử lý', 'Tạo lúc', 'Cập nhật', 'Giao liền'];
var DB_TT = ['Mới', 'Đã xác nhận', 'Đã đến', 'Đã giao', 'Hủy'];
function dbTabMenu_() { return sqTab_(SQ_MENU, H_SQ_MENU, function (sh) { sh.getRange('A:C').setNumberFormat('@'); sh.getRange('F:F').setNumberFormat('@'); }); }
function dbTab_() { return sqTab_(SQ_DB, H_SQ_DB, function (sh) { sh.getRange('A:D').setNumberFormat('@'); sh.getRange('F:J').setNumberFormat('@'); sh.getRange('N:S').setNumberFormat('@'); }); }
function dbCai_() {
  var c = sqJ_(PropertiesService.getScriptProperties().getProperty('DB_CAI'), {}) || {};
  if (c.luuY == null) c.luuY = 'Quán giữ bàn 15 phút. Tiệc trên 15 người vui lòng gọi 0917 203 778.';
  if (!Array.isArray(c.ship)) c.ship = [[100000, 10000], [200000, 20000]];
  if (!Array.isArray(c.hinh)) c.hinh = [];
  if (c.mo == null) c.mo = true;
  if (!Array.isArray(c.nhom)) c.nhom = [];   // thứ tự nhóm món anh xếp
  if (!c.gioMo) c.gioMo = '10:00'; if (!c.gioDong) c.gioDong = '21:30';
  return c;
}
function dbMenu_(tatCa) {
  var sh = dbTabMenu_(), n = sh.getLastRow(), ds = [];
  if (n > 1) sh.getRange(2, 1, n - 1, H_SQ_MENU.length).getValues().forEach(function (r, i) {
    var ten = String(r[1] || '').trim(); if (!ten) return;
    var bt = sqJ_(String(r[2] || ''), []).filter(function (b) { return b && b.gia >= 0; });
    var ban = String(r[3]) === 'x';
    if (!tatCa && (!ban || !bt.length)) return;
    ds.push({ row: i + 2, nhom: String(r[0] || 'Khác'), ten: ten, bt: bt, ban: ban, tt: so_(r[4]) || 999, mo: String(r[5] || '') });
  });
  var nh = dbCai_().nhom, vt = function (n) { var i = nh.indexOf(n); return i < 0 ? 999 : i; };
  return ds.sort(function (a, b) { return vt(a.nhom) - vt(b.nhom) || a.tt - b.tt; });
}
// mức hỗ trợ ship theo tạm tính
function dbHoTro_(tam, c) { var h = 0; (c.ship || []).forEach(function (m) { if (tam >= so_(m[0]) && so_(m[1]) > h) h = so_(m[1]); }); return h; }

/* ---------- khách (không cần mật khẩu) ---------- */
function khMenu_() {
  var c = dbCai_();
  return { menu: dbMenu_(false).map(function (x) { return { nhom: x.nhom, ten: x.ten, bt: x.bt, mo: x.mo }; }), cai: { luuY: c.luuY, ship: c.ship, hinh: c.hinh, mo: c.mo, gioMo: c.gioMo, gioDong: c.gioDong }, homNay: homNay_() };
}
function khDat_(p) {
  var c = dbCai_(); if (!c.mo) throw new Error('Quán đang tạm ngưng nhận đơn online, vui lòng gọi 0917 203 778.');
  var loai = p.loai === 'mangve' ? 'Mang về' : 'Đặt bàn';
  var ten = chu_(p.ten, 60), sdt = String(p.sdt || '').replace(/[^0-9+]/g, '').slice(0, 15);
  if (!ten) throw new Error('Vui lòng ghi tên.');
  if (!/^(\+?84|0)\d{8,10}$/.test(sdt)) throw new Error('Số điện thoại chưa đúng.');
  // chống spam: 1 số điện thoại tối đa 4 đơn / 30 phút, cả quán 80 đơn / giờ
  var cache = CacheService.getScriptCache(), k1 = 'DBS_' + sdt, k2 = 'DBS_ALL', n1 = +(cache.get(k1) || 0), n2 = +(cache.get(k2) || 0);
  if (n1 >= 4) throw new Error('Bạn vừa gửi nhiều đơn, vui lòng chờ quán gọi lại hoặc gọi 0917 203 778.');
  if (n2 >= 80) throw new Error('Quán đang nhận quá nhiều đơn, vui lòng gọi 0917 203 778.');
  var hn = homNay_(), ngay = /^\d{4}-\d{2}-\d{2}$/.test(String(p.ngay || '')) ? String(p.ngay) : hn;
  if (ngay < hn || ngay > congNgay_(hn, 60)) throw new Error('Ngày chưa đúng.');
  var lien = loai === 'Mang về' && !!p.lien, gio = lien ? '' : (/^\d{1,2}:\d{2}$/.test(String(p.gio || '')) ? String(p.gio) : '');
  if (!lien && !gio) throw new Error('Vui lòng chọn giờ.');
  var soNguoi = loai === 'Đặt bàn' ? Math.max(1, Math.min(200, Math.round(so_(p.soNguoi)))) : '';
  var diaChi = loai === 'Mang về' ? chu_(p.diaChi, 200) : '';
  if (loai === 'Mang về' && diaChi.length < 6) throw new Error('Vui lòng ghi địa chỉ giao.');
  // món: giá lấy theo menu của quán (không tin giá gửi lên)
  var menu = {}; dbMenu_(false).forEach(function (x) { x.bt.forEach(function (b) { menu[x.ten + '|' + (b.t || '')] = b.gia; }); });
  var mon = (Array.isArray(p.mon) ? p.mon : []).slice(0, 60).map(function (m) {
    var k = chu_(m.ten, 80) + '|' + chu_(m.bt, 40), gia = menu[k], sl = Math.max(0, Math.min(99, Math.round(so_(m.sl))));
    return gia == null || !sl ? null : { ten: chu_(m.ten, 80), bt: chu_(m.bt, 40), sl: sl, gia: gia };
  }).filter(Boolean);
  if (loai === 'Mang về' && !mon.length) throw new Error('Vui lòng chọn món.');
  var tam = mon.reduce(function (s, m) { return s + m.sl * m.gia; }, 0), hoTro = loai === 'Mang về' ? dbHoTro_(tam, c) : 0;
  var id = 'DB' + Utilities.formatDate(new Date(), TZ, 'ddMM') + Math.random().toString(36).slice(2, 6).toUpperCase();
  khoa_(function () {
    dbTab_().appendRow([id, loai, ngay, gio, soNguoi, ten, sdt, diaChi, chu_(p.dip, 30), JSON.stringify(mon), tam, hoTro, '', chu_(p.ghiChu, 300), 'Mới', '', '', bayGio_(), '', lien ? 'x' : '']);
  });
  cache.put(k1, String(n1 + 1), 1800); cache.put(k2, String(n2 + 1), 3600);
  try { dbBao_({ id: id, loai: loai, ngay: ngay, gio: gio, lien: lien, soNguoi: soNguoi, ten: ten, dip: chu_(p.dip, 30), soMon: mon.length, tam: tam }); } catch (e) { }
  return { id: id, tam: tam, hoTro: hoTro };
}
// gửi thông báo (Firebase) – có khóa FCM thì gửi, chưa có thì bỏ qua
function dbBao_(o) { if (typeof fcmGuiTatCa_ === 'function') fcmGuiTatCa_(o); }

/* ---------- nhân viên / anh Đạt ---------- */
function dbDong_(r, row) {
  return { row: row, id: String(r[0]), loai: String(r[1]), ngay: chuanNgay_(r[2]), gio: String(r[3] || ''), soNguoi: r[4] === '' ? '' : so_(r[4]), ten: String(r[5] || ''), sdt: String(r[6] || ''),
    diaChi: String(r[7] || ''), dip: String(r[8] || ''), mon: sqJ_(String(r[9] || ''), []), tam: so_(r[10]), hoTro: so_(r[11]), ship: r[12] === '' ? '' : so_(r[12]), ghiChu: String(r[13] || ''),
    tt: String(r[14] || 'Mới'), ban: String(r[15] || ''), nguoi: String(r[16] || ''), tao: String(r[17] || ''), cap: String(r[18] || ''), lien: String(r[19] || '') === 'x' };
}
function dbDs_(p) {
  var hn = homNay_(), tu = /^\d{4}-\d{2}-\d{2}$/.test(String(p.tu || '')) ? p.tu : congNgay_(hn, -1), den = /^\d{4}-\d{2}-\d{2}$/.test(String(p.den || '')) ? p.den : congNgay_(hn, 14);
  var sh = dbTab_(), n = sh.getLastRow(), ds = [];
  if (n > 1) sh.getRange(2, 1, n - 1, H_SQ_DB.length).getValues().forEach(function (r, i) { var x = dbDong_(r, i + 2); if (x.id && x.ngay >= tu && x.ngay <= den) ds.push(x); });
  ds.sort(function (a, b) { return (a.ngay + (a.lien ? '00:00' : a.gio)) < (b.ngay + (b.lien ? '00:00' : b.gio)) ? -1 : 1; });
  return { tu: tu, den: den, homNay: hn, ds: ds, moi: ds.filter(function (x) { return x.tt === 'Mới'; }).length };
}
function dbCapNhat_(nguoi, p) {
  var id = String(p.id || ''); if (!/^DB\w{4,12}$/.test(id)) throw new Error('Mã đơn sai.');
  return khoa_(function () {
    var sh = dbTab_(), n = sh.getLastRow(); if (n < 2) throw new Error('Không thấy đơn.');
    var ids = sh.getRange(2, 1, n - 1, 1).getValues(), row = 0;
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) { row = i + 2; break; }
    if (!row) throw new Error('Không thấy đơn.');
    if (p.tt !== undefined) { if (DB_TT.indexOf(p.tt) < 0) throw new Error('Trạng thái sai.'); sh.getRange(row, 15).setValue(p.tt); }
    if (p.ban !== undefined) sh.getRange(row, 16).setValue(chu_(p.ban, 30));
    if (p.ship !== undefined) sh.getRange(row, 13).setValue(p.ship === '' || p.ship === null ? '' : Math.max(0, Math.round(so_(p.ship))));
    sh.getRange(row, 17).setValue(chu_(nguoi, 40)); sh.getRange(row, 19).setValue(bayGio_());
    return { don: dbDong_(sh.getRange(row, 1, 1, H_SQ_DB.length).getValues()[0], row) };
  });
}
// anh Đạt: nạp món từ Sapo (giữ món đã có: không đổi giá / tick bán anh đã chỉnh), lưu từng món, cài đặt
function dbMenuNap_(p) {
  var ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 400);
  return khoa_(function () {
    var cu = {}; dbMenu_(true).forEach(function (x) { cu[x.ten] = x; });
    var them = [], tt = Object.keys(cu).length;
    ds.forEach(function (x) {
      var ten = chu_(x.ten, 80); if (!ten) return;
      var bt = (Array.isArray(x.bt) ? x.bt : []).slice(0, 12).map(function (b) { return { t: chu_(b.t, 40), gia: Math.max(0, Math.round(so_(b.gia))) }; });
      if (cu[ten]) { // thêm biến thể mới chưa có
        var o = cu[ten], co = {}; o.bt.forEach(function (b) { co[b.t] = 1; }); var moi = bt.filter(function (b) { return !co[b.t]; });
        if (moi.length) dbTabMenu_().getRange(o.row, 3).setValue(JSON.stringify(o.bt.concat(moi)));
        return; }
      cu[ten] = 1; tt++;
      them.push([chu_(x.nhom, 40) || 'Khác', ten, JSON.stringify(bt), x.ban ? 'x' : '', tt, '']);
    });
    if (them.length) { var sh = dbTabMenu_(); sh.getRange(sh.getLastRow() + 1, 1, them.length, H_SQ_MENU.length).setValues(them); }
    return { them: them.length };
  });
}
function dbMenuLuu_(p) {
  var row = Math.round(so_(p.row)); if (row < 2) throw new Error('Dòng sai.');
  var bt = (Array.isArray(p.bt) ? p.bt : []).slice(0, 12).map(function (b) { return { t: chu_(b.t, 40), gia: Math.max(0, Math.round(so_(b.gia))) }; });
  return khoa_(function () {
    dbTabMenu_().getRange(row, 1, 1, H_SQ_MENU.length).setValues([[chu_(p.nhom, 40) || 'Khác', chu_(p.ten, 80), JSON.stringify(bt), p.ban ? 'x' : '', Math.round(so_(p.tt)) || 999, chu_(p.mo, 120)]]);
    return {};
  });
}
function dbMenuBanNhieu_(p) {   // tick / bỏ tick nhiều món 1 lần
  var rows = (Array.isArray(p.rows) ? p.rows : []).slice(0, 400).map(function (r) { return Math.round(so_(r)); }).filter(function (r) { return r >= 2; });
  return khoa_(function () { var sh = dbTabMenu_(); rows.forEach(function (r) { sh.getRange(r, 4).setValue(p.ban ? 'x' : ''); }); return { n: rows.length }; });
}
function dbCaiLuu_(p) {
  var c = dbCai_();
  if (p.luuY !== undefined) c.luuY = chu_(p.luuY, 300);
  if (Array.isArray(p.ship)) c.ship = p.ship.slice(0, 6).map(function (m) { return [Math.max(0, Math.round(so_(m[0]))), Math.max(0, Math.round(so_(m[1])))]; }).filter(function (m) { return m[0] > 0 && m[1] > 0; });
  if (Array.isArray(p.hinh)) c.hinh = p.hinh.slice(0, 6).map(function (h) { return { ten: chu_(h.ten, 40), url: String(h.url || '').slice(0, 300) }; }).filter(function (h) { return /^https:\/\//.test(h.url) || /^[\w\/.-]+$/.test(h.url); });
  if (p.mo !== undefined) c.mo = !!p.mo;
  if (Array.isArray(p.nhom)) c.nhom = p.nhom.slice(0, 40).map(function (n) { return chu_(n, 40); }).filter(String);
  if (/^\d{1,2}:\d{2}$/.test(String(p.gioMo || ''))) c.gioMo = p.gioMo; if (/^\d{1,2}:\d{2}$/.test(String(p.gioDong || ''))) c.gioDong = p.gioDong;
  PropertiesService.getScriptProperties().setProperty('DB_CAI', JSON.stringify(c));
  return { cai: c };
}
