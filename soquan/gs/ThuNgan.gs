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

function sqTnLuu_(nv, p) {
  sqTnQuyen_(nv);
  var ngay = sqKiemNgay_(p.ngay), hn = homNay_();
  if (ngay > hn) throw new Error('Không lưu trước cho ngày mai được.');
  if (ngay < congNgay_(hn, -7)) throw new Error('Chỉ sửa được sổ trong 7 ngày gần nhất. Ngày cũ hơn nhờ anh Đạt sửa.');
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
  o.guiSoCu = sqTnGuiSoCu_(o);
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
