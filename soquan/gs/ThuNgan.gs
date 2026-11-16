/**
 * THU NGÂN CUỐI CA – Sổ Quán (thay app "Sổ Thu Ngân Gặm" cũ).
 * Ô nhập giữ y như app cũ: Giao ca, DTTM, Nội dung chuyển khoản, Các khoản chi, Nhận thêm tiền mặt, Tip, TMTT.
 * TM dự kiến = Giao ca + DTTM − Thẻ/CK − Chi + Nhận thêm;  Chênh lệch = TMTT − (TM dự kiến + Tip).
 * Lưu ở tab SQ_ThuNgan (mỗi ngày + ca 1 dòng). Điều hành đọc ngay. Lệch quá 50.000đ = báo đỏ.
 * Đồng thời gửi tiếp sang file "Sổ thu ngân" cũ để P&L chạy như trước: anh tự thêm 2 Script Properties
 * TN_SYNC_URL (link web app sổ cũ) và TN_SYNC_TOKEN (mã của sổ cũ). Chưa có thì bỏ qua bước gửi.
 */
var SQ_TN = 'SQ_ThuNgan';
var H_SQ_TN = ['Ngày', 'Ca', 'Thu ngân', 'Giao ca', 'DTTM', 'Thẻ/CK', 'Chi', 'Nhận thêm', 'Tip', 'TM dự kiến (gồm tip)', 'TMTT', 'Chênh lệch',
  'Chi tiết CK (JSON)', 'Chi tiết Chi (JSON)', 'Chi tiết Nhận (JSON)', 'Lưu lúc', 'Gửi sổ cũ', 'Đã xử lý',
  'Lý do lệch', 'Đếm tiền (JSON)', 'Chia tiền (JSON)'];
var TN_LECH_DO = 50000;

function sqTabTN_() {
  return sqTab_(SQ_TN, H_SQ_TN, function (sh) { sh.getRange('A:A').setNumberFormat('@'); sh.getRange('P:Q').setNumberFormat('@'); });
}
// tab cũ chưa có 3 cột mới (Lý do lệch, Đếm tiền, Chia tiền) → thêm tiêu đề
function sqTnCotMoi_(sh) {
  if (String(sh.getRange(1, 19).getValue()) === H_SQ_TN[18]) return;
  sh.getRange(1, 19, 1, 3).setValues([H_SQ_TN.slice(18, 21)]).setFontWeight('bold').setBackground('#445803').setFontColor('#F8E8C4');
  sh.getRange('T:U').setNumberFormat('@');
}
var TN_MENH_GIA = [500000, 200000, 100000, 50000, 20000, 10000, 5000, 2000, 1000];
function sqTnDem_(d) { // {to:{500000:n,...}} → {to, tong}
  if (!d || typeof d !== 'object') return null;
  var to = {}, tong = 0, co = false;
  TN_MENH_GIA.forEach(function (m) { var n = Math.max(0, Math.round(so_(d.to && d.to[m]))); if (n) { to[m] = n; tong += n * m; co = true; } });
  return co ? { to: to, tong: tong } : null;
}
function sqTnChia_(d) {
  if (!d || typeof d !== 'object') return null;
  var phan = function (x) { var to = {}; TN_MENH_GIA.forEach(function (m) { var n = Math.max(0, Math.round(so_(x && x[m]))); if (n) to[m] = n; }); return to; };
  return { giaoCa: Math.round(so_(d.giaoCa)), cho: Math.round(so_(d.cho)), son: Math.round(so_(d.son)), muonGiaoCa: Math.round(so_(d.muonGiaoCa)), muonCho: Math.round(so_(d.muonCho)),
    to: { giaoCa: phan(d.to && d.to.giaoCa), cho: phan(d.to && d.to.cho), son: phan(d.to && d.to.son) } };
}
function sqTnQuyen_(nv) { if ((nv.chucNang || []).indexOf('tn') < 0) throw new Error('Em chưa được mở chức năng Thu ngân. Nhờ anh Đạt bật trong Điều hành → Nhân viên.'); }
function sqTnDong_(ds) {
  return (Array.isArray(ds) ? ds : []).slice(0, 60).map(function (r) {
    return { note: chu_(r && r.note, 80), amount: Math.round(so_(r && r.amount)) };
  }).filter(function (r) { return r.note || r.amount; });
}
function sqTnTong_(ds) { return ds.reduce(function (s, r) { return s + r.amount; }, 0); }
function sqTnTinh_(o) {
  o.tr = sqTnTong_(o.ck); o.ex = sqTnTong_(o.chi); o.rec = sqTnTong_(o.nhan);
  o.duKien = o.giaoCa + o.dttm - o.tr - o.ex + o.rec;
  o.duKienTip = o.duKien + o.tip;
  o.lech = o.tmtt - o.duKienTip;
  return o;
}
function sqTnObj_(r, i) {
  return sqTnTinh_({
    row: i + 2, ngay: chuanNgay_(r[0]), ca: so_(r[1]), nguoi: String(r[2] || ''), giaoCa: so_(r[3]), dttm: so_(r[4]), tip: so_(r[8]), tmtt: so_(r[10]),
    ck: sqJ_(r[12], []), chi: sqJ_(r[13], []), nhan: sqJ_(r[14], []), luc: String(r[15] || ''), guiSoCu: String(r[16] || ''), xuLy: String(r[17] || ''),
    lyDo: String(r[18] || ''), dem: sqJ_(r[19], null), chia: sqJ_(r[20], null)
  });
}
function sqTnRows_() {
  var sh = sqTabTN_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_SQ_TN.length).getValues().map(sqTnObj_).filter(function (o) { return o.ngay && (o.ca === 1 || o.ca === 2); });
}
function sqTnGon_(o) {
  return { ngay: o.ngay, ca: o.ca, nguoi: o.nguoi, giaoCa: o.giaoCa, dttm: o.dttm, tip: o.tip, tmtt: o.tmtt, ck: o.ck, chi: o.chi, nhan: o.nhan,
    tr: o.tr, ex: o.ex, rec: o.rec, duKien: o.duKien, duKienTip: o.duKienTip, lech: o.lech, luc: o.luc, guiSoCu: o.guiSoCu, xuLy: o.xuLy,
    lyDo: o.lyDo || '', dem: o.dem || null, chia: o.chia || null };
}

// Nhân viên: số 2 ca của 1 ngày + tóm tắt tháng
function sqTnNgay_(nv, p) {
  sqTnQuyen_(nv);
  var ngay = sqKiemNgay_(p.ngay || homNay_()), ca = [null, null];
  sqTnRows_().forEach(function (o) { if (o.ngay === ngay) ca[o.ca - 1] = sqTnGon_(o); });
  return { ngay: ngay, ca: ca, homNay: homNay_() };
}
function sqTnThang_(nv, p) {
  sqTnQuyen_(nv);
  var thang = /^\d{4}-\d{2}$/.test(String(p.thang || '')) ? p.thang : homNay_().slice(0, 7);
  var ds = sqTnRows_().filter(function (o) { return o.ngay.slice(0, 7) === thang; }).map(sqTnGon_)
    .sort(function (a, b) { return (a.ngay + a.ca) < (b.ngay + b.ca) ? -1 : 1; });
  return { thang: thang, ds: ds };
}

function sqTnLuu_(nv, p, chu) {
  if (!chu) sqTnQuyen_(nv);
  var ngay = sqKiemNgay_(p.ngay), hn = homNay_();
  if (ngay > hn) throw new Error('Không lưu trước cho ngày mai được.');
  if (!chu && ngay < congNgay_(hn, -7)) throw new Error('Chỉ sửa được sổ trong 7 ngày gần nhất. Ngày cũ hơn nhờ anh Đạt sửa.');
  var ca = so_(p.ca);
  if (ca !== 1 && ca !== 2) throw new Error('Chọn Ca 1 hoặc Ca 2.');
  var o = sqTnTinh_({
    ngay: ngay, ca: ca, nguoi: nv.ten, giaoCa: Math.round(so_(p.giaoCa)), dttm: Math.round(so_(p.dttm)), tip: Math.round(so_(p.tip)), tmtt: Math.round(so_(p.tmtt)),
    ck: sqTnDong_(p.ck), chi: sqTnDong_(p.chi), nhan: sqTnDong_(p.nhan), luc: bayGio_(), guiSoCu: '', xuLy: '',
    lyDo: chu_(p.lyDo, 300), dem: sqTnDem_(p.dem), chia: sqTnChia_(p.chia)
  });
  if (o.lech >= 0) o.lyDo = ''; // chỉ giữ lý do khi két thiếu (lệch âm)
  khoa_(function () {
    var sh = sqTabTN_(), rows = sqTnRows_(), row = -1;
    sqTnCotMoi_(sh);
    rows.forEach(function (x) { if (x.ngay === ngay && x.ca === ca) row = x.row; });
    if (row < 0) row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, H_SQ_TN.length).setValues([[ngay, ca, o.nguoi, o.giaoCa, o.dttm, o.tr, o.ex, o.rec, o.tip, o.duKienTip, o.tmtt, o.lech,
      JSON.stringify(o.ck), JSON.stringify(o.chi), JSON.stringify(o.nhan), o.luc, 'Đang gửi', '',
      o.lyDo, o.dem ? JSON.stringify(o.dem) : '', o.chia ? JSON.stringify(o.chia) : '']]);
    o.row = row;
  });
  o.guiSoCu = chu ? 'anh Đạt nhập hộ' : sqTnGuiSoCu_(o); // anh nhập hộ: ca đã có ở sổ cũ, không gửi lại
  try { sqTabTN_().getRange(o.row, 17).setValue(o.guiSoCu); } catch (e) { }
  return { ca: sqTnGon_(o) };
}

// Gửi tiếp 1 ca sang Apps Script "Sổ thu ngân" cũ (P&L đang đọc file đó)
function sqTnTien_(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ' đ'; }
function sqTnGhiChu_(ds) { return ds.filter(function (r) { return r.amount; }).map(function (r) { return (r.note || '(?)') + ': ' + sqTnTien_(r.amount); }).join(' | '); }
function sqTnGuiSoCu_(o) {
  var props = PropertiesService.getScriptProperties();
  var url = props.getProperty('TN_SYNC_URL'), token = props.getProperty('TN_SYNC_TOKEN');
  if (!url || !token) return 'Chưa nối sổ cũ';
  var row = { date: o.ngay, shift: o.ca, handover: o.giaoCa, revenue: o.dttm, tr: o.tr, ex: o.ex, received: o.rec, tips: o.tip,
    expected: o.duKienTip, actual: o.tmtt, diff: o.lech, trNotes: sqTnGhiChu_(o.ck), exNotes: sqTnGhiChu_(o.chi), recNotes: sqTnGhiChu_(o.nhan),
    savedAt: new Date().toISOString() };
  try {
    var res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify({ token: token, row: row }),
      muteHttpExceptions: true, followRedirects: true });
    var code = res.getResponseCode();
    return code < 400 ? 'Đã gửi ' + bayGio_().slice(-5) : 'Lỗi ' + code;
  } catch (e) { return 'Lỗi gửi'; }
}

// Điều hành: các ca trong khoảng ngày + cảnh báo lệch
function sqTnChu_(p) {
  var den = /^\d{4}-\d{2}-\d{2}$/.test(String(p.den || '')) ? p.den : homNay_();
  var tu = /^\d{4}-\d{2}-\d{2}$/.test(String(p.tu || '')) ? p.tu : congNgay_(den, -40);
  var ds = sqTnRows_().filter(function (o) { return o.ngay >= tu && o.ngay <= den; }).map(sqTnGon_)
    .sort(function (a, b) { return (a.ngay + a.ca) < (b.ngay + b.ca) ? 1 : -1; });
  var coNoi = !!PropertiesService.getScriptProperties().getProperty('TN_SYNC_URL');
  // doanh thu theo phương thức thanh toán (file Sapo Sơn gửi cuối ngày) – để anh thấy doanh thu sàn Grab / Shopee / Be
  var san = {};
  try { sqSapoRows_().forEach(function (x) { if (x.loai === 'tt' && x.ngay >= tu && x.ngay <= den) san[x.ngay] = { kenh: x.d.kenh || [], tong: x.d.tong, don: x.d.don, luc: x.luc }; }); } catch (e) { }
  return { tu: tu, den: den, ds: ds, san: san, lechDo: TN_LECH_DO, coNoiSoCu: coNoi };
}
function sqTnXuLy_(p) {
  var ngay = sqKiemNgay_(p.ngay), ca = so_(p.ca), ghi = chu_(p.ghi, 120) || ('Đã xử lý ' + bayGio_());
  return khoa_(function () {
    var x = sqTnRows_().filter(function (o) { return o.ngay === ngay && o.ca === ca; })[0];
    if (!x) throw new Error('Không tìm thấy ca này.');
    sqTabTN_().getRange(x.row, 18).setValue(p.bo ? '' : ghi);
    return { xuLy: p.bo ? '' : ghi };
  });
}

/* =====================================================================
 * 🏦 NGÂN HÀNG – anh Đạt nạp file Excel sao kê (Techcombank, MB…), app đọc trên trình duyệt,
 * gửi từng giao dịch lên lưu tab SQ_NganHang (trùng mã thì bỏ qua). Chỉ anh Đạt xem (mật khẩu chủ).
 * Cài đặt (tên TK, mức khoản nhỏ, giờ khuya, tên TK vợ, % sàn, nội dung quen → mục) ở Script Property NH_CAI.
 * ===================================================================== */
var SQ_NH = 'SQ_NganHang', H_SQ_NH = ['Mã', 'TK', 'Ngày', 'Giờ', 'Nội dung', 'Đối tác', 'Ra (nợ)', 'Vào (có)', 'Số dư', 'Mục', 'Duyệt', 'Nạp lúc'];
function sqTabNH_() { return sqTab_(SQ_NH, H_SQ_NH, function (sh) { sh.getRange('A:F').setNumberFormat('@'); sh.getRange('J:L').setNumberFormat('@'); }); }
function nhCai_() {
  var c = sqJ_(PropertiesService.getScriptProperties().getProperty('NH_CAI'), {}) || {};
  c.tk = c.tk || {}; c.quen = c.quen || {};
  if (!(c.nho > 0)) c.nho = 500000;
  if (c.khuyaTu == null) c.khuyaTu = 23; if (c.khuyaDen == null) c.khuyaDen = 6;
  c.pct = Object.assign({ shopee: 19, grab: 19.5, be: 19.5 }, c.pct || {});
  if (c.vo == null) c.vo = '';
  return c;
}
function nhCaiGhi_(c) { PropertiesService.getScriptProperties().setProperty('NH_CAI', JSON.stringify(c)); }
function nhTien_(v) { var n = Math.round(Math.abs(so_(v))); return n > 0 ? n : 0; }
function nhLuu_(p) {
  var tk = chu_(p.tk, 30); if (!/^[0-9A-Za-z]{4,30}$/.test(tk)) throw new Error('Không đọc được số tài khoản trong file.');
  var ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 5000); if (!ds.length) throw new Error('File không có giao dịch nào.');
  return khoa_(function () {
    var sh = sqTabNH_(), n = sh.getLastRow(), co = {};
    if (n > 1) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (r) { co[String(r[0])] = 1; });
    var luc = bayGio_(), moi = [], trung = 0, tu = '', den = '';
    ds.forEach(function (x) {
      var ngay = chuanNgay_(x.ngay); if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) return;
      var no = nhTien_(x.no), cov = nhTien_(x.co); if (!no && !cov) return;
      var ref = chu_(x.ref, 40), ma = tk + '|' + (ref || (ngay + '|' + no + '|' + cov + '|' + Math.round(so_(x.sd))));
      if (!tu || ngay < tu) tu = ngay; if (!den || ngay > den) den = ngay;
      if (co[ma]) { trung++; return; } co[ma] = 1;
      var gio = /^\d{1,2}:\d{2}/.test(String(x.gio || '')) ? String(x.gio).slice(0, 5) : '';
      moi.push([ma, tk, ngay, gio, chu_(x.nd, 300), chu_(x.dt, 80), no || '', cov || '', Math.round(so_(x.sd)) || '', '', '', luc]);
    });
    if (moi.length) sh.getRange(sh.getLastRow() + 1, 1, moi.length, H_SQ_NH.length).setValues(moi);
    var c = nhCai_(), t = c.tk[tk] || {};
    c.tk[tk] = { ten: chu_(p.ten, 40) || t.ten || ('TK ··' + tk.slice(-4)), nh: chu_(p.nh, 30) || t.nh || '', vai: t.vai || '', tu: !t.tu || (tu && tu < t.tu) ? tu : t.tu, den: !t.den || den > t.den ? den : t.den, luc: luc };
    nhCaiGhi_(c);
    return { moi: moi.length, trung: trung, tu: tu, den: den };
  });
}
function nhXem_(p) {
  var hn = homNay_(), den = /^\d{4}-\d{2}-\d{2}$/.test(String(p.den || '')) ? p.den : hn, tu = /^\d{4}-\d{2}-\d{2}$/.test(String(p.tu || '')) ? p.tu : den;
  if (tu > den) { var t0 = tu; tu = den; den = t0; }
  var tuDoc = congNgay_(tu, -40);   // đọc lùi để biết người nhận / nội dung đã từng gặp
  var sh = sqTabNH_(), n = sh.getLastRow(), ds = [], cu = {};
  if (n > 1) sh.getRange(2, 1, n - 1, H_SQ_NH.length).getValues().forEach(function (r) {
    var ngay = chuanNgay_(r[2]); if (ngay < tuDoc || ngay > den) return;
    var x = [String(r[0]), String(r[1]), ngay, String(r[3] || ''), String(r[4] || ''), String(r[5] || ''), so_(r[6]), so_(r[7]), so_(r[8]), String(r[9] || ''), String(r[10] || '')];
    if (ngay < tu) { if (x[6]) cu[khongDau_(x[5] || x[4]).replace(/[0-9]+/g, '').slice(0, 40)] = 1; return; }
    ds.push(x);
  });
  var c = nhCai_();
  return { tu: tu, den: den, homNay: hn, cai: c, ds: ds, cu: Object.keys(cu) };
}
function nhTimDong_(sh, ma) {
  var n = sh.getLastRow(); if (n < 2) return 0;
  var a = sh.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < a.length; i++) if (String(a[i][0]) === ma) return i + 2;
  return 0;
}
function nhDanh_(p) {
  return khoa_(function () {
    var c = nhCai_(), sh = sqTabNH_(), out = {};
    if (p.ma) {
      var row = nhTimDong_(sh, String(p.ma)); if (!row) throw new Error('Không thấy giao dịch.');
      if (p.muc !== undefined) sh.getRange(row, 10).setValue(chu_(p.muc, 20));
      if (p.duyet !== undefined) sh.getRange(row, 11).setValue(p.duyet ? chu_(p.duyet, 10) + ' · ' + bayGio_() : '');
      out.ma = p.ma;
    }
    if (p.quenKey) { var k = khongDau_(p.quenKey).replace(/[0-9]+/g, '').trim().slice(0, 40); if (k) { if (p.quenMuc === null) delete c.quen[k]; else c.quen[k] = chu_(p.quenMuc || 'ok', 20); nhCaiGhi_(c); } }
    return out;
  });
}
function nhCaiLuu_(p) {
  var c = nhCai_();
  if (p.tk && c.tk[p.tk]) { if (p.ten !== undefined) c.tk[p.tk].ten = chu_(p.ten, 40); if (p.vai !== undefined) c.tk[p.tk].vai = chu_(p.vai, 10); }
  if (p.nho !== undefined && so_(p.nho) > 0) c.nho = Math.round(so_(p.nho));
  if (p.khuyaTu !== undefined) c.khuyaTu = Math.max(0, Math.min(23, Math.round(so_(p.khuyaTu))));
  if (p.khuyaDen !== undefined) c.khuyaDen = Math.max(0, Math.min(23, Math.round(so_(p.khuyaDen))));
  if (p.vo !== undefined) c.vo = chu_(p.vo, 60);
  if (p.pct) ['shopee', 'grab', 'be'].forEach(function (k) { if (p.pct[k] !== undefined && so_(p.pct[k]) >= 0 && so_(p.pct[k]) < 60) c.pct[k] = so_(p.pct[k]); });
  nhCaiGhi_(c);
  return { cai: c };
}

/**
 * THƯỞNG NÓNG – Sổ Quán (thu ngân / quản lý bấm cuối ngày).
 * Chọn mức 1–4 → tick người làm hôm nay (Bếp · Phục vụ · Bảo vệ) → Ca 1 / Ca 2 / Cả ngày → app tự tính tiền.
 * Lưu ở tab SQ_ThuongNong (mỗi người 1 dòng / ngày; lưu lại thì thay cả ngày). Tiền chi ra vẫn ghi ở "Các khoản chi"
 * của sổ thu ngân (dòng "Thưởng nóng …") để két và P&L chạy như cũ – app tự điền dòng đó.
 * Danh sách người = nhân viên đang làm trong SQ_NhanVien (vai Bếp / Đi chợ → Bếp; Phục vụ / Thu ngân → Phục vụ)
 * + tên thu ngân tự thêm (tab SQ_ThuongNguoi). Ai nghỉ làm (anh Đạt đánh dấu nghỉ) hoặc bị bỏ khỏi danh sách thì không hiện.
 */
var SQ_THUONG = 'SQ_ThuongNong', SQ_THUONG_NGUOI = 'SQ_ThuongNguoi';
var H_SQ_THUONG = ['Ngày', 'Mức', 'Tên', 'Nhóm', 'Ca', 'Tiền theo mức', 'Tiền thưởng', 'Lý do chỉnh', 'Ghi bởi', 'Lưu lúc'];
var H_SQ_THUONG_NGUOI = ['Tên', 'Nhóm', 'Bỏ khỏi danh sách (x)', 'Thêm bởi', 'Lúc'];
// tiền theo mức: [cả ngày, 1 ca]
var THUONG_MUC = { 1: [50000, 50000], 2: [100000, 70000], 3: [150000, 100000], 4: [200000, 150000] };
var THUONG_NHOM = { bep: 'Bếp', pv: 'Phục vụ', bv: 'Bảo vệ' };

function sqTabThuong_() { return sqTab_(SQ_THUONG, H_SQ_THUONG, function (sh) { sh.getRange('A:A').setNumberFormat('@'); }); }
function sqTabThuongNguoi_() { return sqTab_(SQ_THUONG_NGUOI, H_SQ_THUONG_NGUOI); }
function sqThuongQuyen_(nv) {
  if ((nv.chucNang || []).indexOf('tn') < 0 && !/^(thu ngan|quan ly)$/.test(khongDau_(nv.vai)))
    throw new Error('Thưởng nóng chỉ thu ngân hoặc quản lý bấm được.');
}
function sqThuongNhomVai_(vai) {
  var v = khongDau_(vai);
  if (/bep|di cho|phu bep/.test(v)) return 'bep';
  if (/phuc vu|thu ngan/.test(v)) return 'pv';
  if (/bao ve/.test(v)) return 'bv';
  return '';
}
function sqThuongRows_() {
  var sh = sqTabThuong_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_SQ_THUONG.length).getValues().map(function (r, i) {
    return { row: i + 2, ngay: chuanNgay_(r[0]), muc: so_(r[1]), ten: String(r[2] || '').trim(), nhom: String(r[3] || ''), ca: String(r[4] || ''),
      tienMuc: so_(r[5]), tien: so_(r[6]), lyDo: String(r[7] || ''), nguoi: String(r[8] || ''), luc: String(r[9] || '') };
  }).filter(function (x) { return x.ngay && x.ten; });
}
function sqThuongNguoiRows_() {
  var sh = sqTabThuongNguoi_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_SQ_THUONG_NGUOI.length).getValues().map(function (r, i) {
    return { row: i + 2, ten: String(r[0] || '').trim(), nhom: String(r[1] || ''), bo: /x/i.test(String(r[2])), boi: String(r[3] || ''), luc: String(r[4] || '') };
  }).filter(function (x) { return x.ten; });
}
// danh sách người hiện trong hộp thưởng nóng
function sqThuongDsNguoi_() {
  var them = {}; sqThuongNguoiRows_().forEach(function (x) { them[khongDau_(x.ten)] = x; });
  var ds = [], co = {};
  sqNVDs_().ds.forEach(function (x) {
    if (x.nghi) return;
    var k = khongDau_(x.ten), t = them[k], nhom = (t && THUONG_NHOM[t.nhom]) ? t.nhom : sqThuongNhomVai_(x.vai);
    if (!nhom || (t && t.bo)) return;
    co[k] = 1; ds.push({ ten: x.ten, nhom: nhom, nguon: 'nv' });
  });
  Object.keys(them).forEach(function (k) {
    var t = them[k]; if (co[k] || t.bo || !THUONG_NHOM[t.nhom]) return;
    co[k] = 1; ds.push({ ten: t.ten, nhom: t.nhom, nguon: 'them' });
  });
  return ds;
}
// p: {ngay}
function sqThuongNgay_(nv, p) {
  sqThuongQuyen_(nv);
  var ngay = sqKiemNgay_(p.ngay || homNay_()), rows = sqThuongRows_();
  var hom = rows.filter(function (x) { return x.ngay === ngay; });
  var caCuoi = {}; // ca lần gần nhất của từng người → mặc định
  rows.filter(function (x) { return x.ngay < ngay; }).sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; })
    .forEach(function (x) { caCuoi[khongDau_(x.ten)] = x.ca; });
  return {
    ngay: ngay, muc: THUONG_MUC, nhom: THUONG_NHOM, nguoi: sqThuongDsNguoi_(), caCuoi: caCuoi,
    daLuu: hom.length ? { muc: hom[0].muc, nguoi: hom[0].nguoi, luc: hom[0].luc,
      ds: hom.map(function (x) { return { ten: x.ten, nhom: x.nhom, ca: x.ca, tienMuc: x.tienMuc, tien: x.tien, lyDo: x.lyDo }; }) } : null
  };
}
// p: {ngay, muc (0 = không thưởng), ds:[{ten, nhom, ca:'1'|'2'|'full', tien, lyDo}]}
function sqThuongLuu_(nv, p) {
  sqThuongQuyen_(nv);
  var ngay = sqKiemNgay_(p.ngay), hn = homNay_();
  if (ngay > hn) throw new Error('Không lưu trước cho ngày mai được.');
  if (ngay < congNgay_(hn, -7)) throw new Error('Chỉ sửa được thưởng nóng trong 7 ngày gần nhất.');
  var muc = so_(p.muc);
  if (muc && !THUONG_MUC[muc]) throw new Error('Mức thưởng không đúng.');
  var luc = bayGio_(), ds = [];
  if (muc) (Array.isArray(p.ds) ? p.ds : []).slice(0, 60).forEach(function (x) {
    var ten = chu_(x && x.ten, 40), nhom = THUONG_NHOM[x && x.nhom] ? x.nhom : '', ca = ['1', '2', 'full'].indexOf(String(x && x.ca)) >= 0 ? String(x.ca) : 'full';
    if (!ten || !nhom) return;
    var tienMuc = THUONG_MUC[muc][ca === 'full' ? 0 : 1], tien = Math.max(0, Math.round(so_(x.tien === '' || x.tien == null ? tienMuc : x.tien)));
    ds.push([ngay, muc, ten, nhom, ca, tienMuc, tien, tien !== tienMuc ? chu_(x.lyDo, 150) : '', nv.ten, luc]);
  });
  if (muc && !ds.length) throw new Error('Chưa tick ai được thưởng.');
  khoa_(function () {
    var sh = sqTabThuong_(), rows = sqThuongRows_(), giu = rows.filter(function (x) { return x.ngay !== ngay; });
    var n = sh.getLastRow();
    if (n > 1) sh.getRange(2, 1, n - 1, H_SQ_THUONG.length).clearContent();
    var all = giu.map(function (x) { return [x.ngay, x.muc, x.ten, x.nhom, x.ca, x.tienMuc, x.tien, x.lyDo, x.nguoi, x.luc]; }).concat(ds)
      .sort(function (a, b) { return String(a[0]) < String(b[0]) ? -1 : String(a[0]) > String(b[0]) ? 1 : 0; });
    if (all.length) sh.getRange(2, 1, all.length, H_SQ_THUONG.length).setValues(all);
  });
  var tong = ds.reduce(function (s, r) { return s + r[6]; }, 0);
  return { ngay: ngay, muc: muc, soNguoi: ds.length, tong: tong };
}
// thêm tên mới / đổi nhóm / bỏ khỏi danh sách. p: {ten, nhom, bo}
function sqThuongNguoiLuu_(nv, p) {
  sqThuongQuyen_(nv);
  var ten = chu_(p.ten, 40), nhom = THUONG_NHOM[p.nhom] ? p.nhom : '';
  if (!ten) throw new Error('Chưa ghi tên.');
  if (!p.bo && !nhom) throw new Error('Chọn nhóm Bếp / Phục vụ / Bảo vệ.');
  khoa_(function () {
    var sh = sqTabThuongNguoi_(), cu = sqThuongNguoiRows_().filter(function (x) { return khongDau_(x.ten) === khongDau_(ten); })[0];
    var vals = [[cu ? cu.ten : ten, nhom || (cu && cu.nhom) || '', p.bo ? 'x' : '', nv.ten, bayGio_()]];
    if (cu) sh.getRange(cu.row, 1, 1, H_SQ_THUONG_NGUOI.length).setValues(vals);
    else sh.getRange(sh.getLastRow() + 1, 1, 1, H_SQ_THUONG_NGUOI.length).setValues(vals);
  });
  return { nguoi: sqThuongDsNguoi_() };
}
// Điều hành: thưởng nóng theo tháng – tổng từng người (để làm bảng lương)
function sqThuongThang_(p) {
  var thang = /^\d{4}-\d{2}$/.test(String(p.thang || '')) ? p.thang : homNay_().slice(0, 7);
  var ng = {}, ngay = {};
  sqThuongRows_().filter(function (x) { return x.ngay.slice(0, 7) === thang; }).forEach(function (x) {
    var k = khongDau_(x.ten), g = ng[k] || (ng[k] = { ten: x.ten, nhom: x.nhom, lan: 0, tong: 0 });
    g.lan++; g.tong += x.tien;
    var d = ngay[x.ngay] || (ngay[x.ngay] = { ngay: x.ngay, muc: x.muc, nguoi: 0, tong: 0, ghiBoi: x.nguoi });
    d.nguoi++; d.tong += x.tien;
  });
  return { thang: thang, nguoi: Object.keys(ng).map(function (k) { return ng[k]; }).sort(function (a, b) { return b.tong - a.tong; }),
    ngay: Object.keys(ngay).sort().map(function (k) { return ngay[k]; }) };
}
