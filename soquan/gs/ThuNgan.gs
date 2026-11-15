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
  'Chi tiết CK (JSON)', 'Chi tiết Chi (JSON)', 'Chi tiết Nhận (JSON)', 'Lưu lúc', 'Gửi sổ cũ', 'Đã xử lý'];
var TN_LECH_DO = 50000;

function sqTabTN_() {
  return sqTab_(SQ_TN, H_SQ_TN, function (sh) { sh.getRange('A:A').setNumberFormat('@'); sh.getRange('P:Q').setNumberFormat('@'); });
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
    ck: sqJ_(r[12], []), chi: sqJ_(r[13], []), nhan: sqJ_(r[14], []), luc: String(r[15] || ''), guiSoCu: String(r[16] || ''), xuLy: String(r[17] || '')
  });
}
function sqTnRows_() {
  var sh = sqTabTN_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_SQ_TN.length).getValues().map(sqTnObj_).filter(function (o) { return o.ngay && (o.ca === 1 || o.ca === 2); });
}
function sqTnGon_(o) {
  return { ngay: o.ngay, ca: o.ca, nguoi: o.nguoi, giaoCa: o.giaoCa, dttm: o.dttm, tip: o.tip, tmtt: o.tmtt, ck: o.ck, chi: o.chi, nhan: o.nhan,
    tr: o.tr, ex: o.ex, rec: o.rec, duKien: o.duKien, duKienTip: o.duKienTip, lech: o.lech, luc: o.luc, guiSoCu: o.guiSoCu, xuLy: o.xuLy };
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
    ck: sqTnDong_(p.ck), chi: sqTnDong_(p.chi), nhan: sqTnDong_(p.nhan), luc: bayGio_(), guiSoCu: '', xuLy: ''
  });
  khoa_(function () {
    var sh = sqTabTN_(), rows = sqTnRows_(), row = -1;
    rows.forEach(function (x) { if (x.ngay === ngay && x.ca === ca) row = x.row; });
    if (row < 0) row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, H_SQ_TN.length).setValues([[ngay, ca, o.nguoi, o.giaoCa, o.dttm, o.tr, o.ex, o.rec, o.tip, o.duKienTip, o.tmtt, o.lech,
      JSON.stringify(o.ck), JSON.stringify(o.chi), JSON.stringify(o.nhan), o.luc, 'Đang gửi', '']]);
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
