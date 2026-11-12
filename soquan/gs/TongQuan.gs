/**
 * TỔNG QUAN – Điều hành (anh Đạt). Gom 1 lần gọi: chi theo ngày (phiếu đi chợ + hóa đơn + hàng tự đặt),
 * sổ thu ngân theo ngày, hủy hàng theo ngày, việc cần xử lý. Doanh thu Sapo: chừa sẵn (dt = null) tới khi có file Excel.
 */
function sqTongNhom_(nhom) {
  nhom = String(nhom || '');
  if (/^HĐ /.test(nhom)) return 'Hóa đơn mối';
  if (nhom === NHOM_TU_DAT) return 'Hóa đơn quen';
  return nhom;
}
function sqTongQuan_(p) {
  var hn = homNay_();
  var thang = /^\d{4}-\d{2}$/.test(String(p.thang || '')) ? p.thang : hn.slice(0, 7);
  var y = +thang.slice(0, 4), m = +thang.slice(5, 7);
  var truoc = (m === 1 ? (y - 1) + '-12' : y + '-' + ('0' + (m - 1)).slice(-2));
  var tu = truoc + '-01', den = thang + '-31';
  var ngay = {};
  var N = function (d) { return ngay[d] || (ngay[d] = { chi: 0, nhom: {}, dttm: null, ck: 0, chiQuay: 0, lech: 0, soCa: 0, lechDo: false, huy: 0, huyChuaGia: 0 }); };

  // chi: tab Phieu (mỗi phiếu 1 dòng: B ngày, C nhóm, F tổng)
  var rows = layFile_().getSheetByName('Phieu').getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i]; if (!String(r[0])) continue;
    var d = chuanNgay_(r[1]); if (d < tu || d > den) continue;
    if (['Chi phí cố định', 'Đầu tư', 'Lương / thưởng'].indexOf(String(r[2])) >= 0) continue;
    var t = so_(r[5]), k = sqTongNhom_(r[2]), o = N(d);
    o.chi += t; o.nhom[k] = (o.nhom[k] || 0) + t;
  }
  // sổ thu ngân
  sqTnRows_().forEach(function (x) {
    if (x.ngay < tu || x.ngay > den) return;
    var o = N(x.ngay);
    o.dttm = (o.dttm || 0) + x.dttm; o.ck += x.tr; o.chiQuay += x.ex; o.lech += x.lech; o.soCa++;
    if (Math.abs(x.lech) > TN_LECH_DO && !x.xuLy) o.lechDo = true;
  });
  // hủy hàng
  sqHuyRows_().forEach(function (x) {
    if (x.ngay < tu || x.ngay > den) return;
    var o = N(x.ngay);
    if (x.tien === '') o.huyChuaGia++; else o.huy += x.tien;
  });

  // số Sapo
  sqSapoRows_().forEach(function (x) {
    if (x.ngay < tu || x.ngay > den) return;
    var o = N(x.ngay);
    if (x.loai === 'tt') { o.dt = x.d.tong; o.don = x.d.don; o.kenh = x.d.kenh; o.sapoLuc = x.luc; }
    else if (x.loai === 'tq') o.tq = x.d;
    else if (x.loai === 'mon' && x.ngay.slice(0, 7) === thang) o.mon = x.d.mon;
  });
  // việc cần xử lý
  var metas = sqMetaRows_(), viec = [];
  var dem = function (f) { return metas.filter(f).length; };
  var choKiem = dem(function (x) { return x.loai === LOAI.HD && TT_CHUA_XAC_NHAN.indexOf(x.tt) >= 0; });
  var choDuyet = dem(function (x) { return x.tt === TT.CHO || x.tt === TT.TRALOI; });
  var hoiLai = dem(function (x) { return x.tt === TT.HOI; });
  sqTnRows_().filter(function (x) { return Math.abs(x.lech) > TN_LECH_DO && !x.xuLy && x.ngay >= congNgay_(hn, -7); })
    .sort(function (a, b) { return a.ngay < b.ngay ? 1 : -1; })
    .forEach(function (x) { viec.push({ muc: 'r', loai: 'Tiền mặt', tab: 'tm', ngay: x.ngay, text: sqDM_(x.ngay) + ' · Ca ' + x.ca + ': két lệch ' + (x.lech < 0 ? '−' : '+') + sqTien_(Math.abs(x.lech)) + ' (' + x.nguoi + ')' }); });
  try {
    sqGiaThayDoi_({ den: hn, tu: congNgay_(hn, -2) }).ds.filter(function (g) { return g.pct >= 0.1; }).slice(0, 4).forEach(function (g) {
      viec.push({ muc: 'r', loai: 'Giá nhập tăng', tab: 'cp', ngay: g.ngay, text: g.ten + ': ' + sqTien_(g.gia) + '/' + (g.dv || 'đv') + ' – tăng ' + Math.round(g.pct * 100) + '% so với lần mua ' + sqDM_(g.truoc.ngay) });
    });
  } catch (e) { }
  // đối chiếu: sổ thu ngân (DTTM) với dòng "Tiền mặt" của Sapo (thu tại quầy)
  Object.keys(ngay).sort().reverse().forEach(function (d) {
    var o = ngay[d]; if (d < congNgay_(hn, -7) || o.dttm == null || !o.kenh) return;
    var tm = o.kenh.filter(function (k) { return khongDau_(k.ten) === 'tien mat'; })[0];
    if (!tm) return;
    var c = o.dttm - tm.tien;
    o.lechSapo = c;
    if (Math.abs(c) > TN_LECH_DO) viec.push({ muc: 'y', loai: 'Đối chiếu Sapo', tab: 'tong', ngay: d, text: sqDM_(d) + ': sổ thu ngân ' + sqTien_(o.dttm) + ' – Sapo tiền mặt ' + sqTien_(tm.tien) + ' (lệch ' + (c < 0 ? '−' : '+') + sqTien_(Math.abs(c)) + ')' });
  });
  // số Sapo: 4 file mỗi tối (thanh toán, mặt hàng, tồn kho, doanh thu tổng quan)
  var co = sqSapoCo_(congNgay_(hn, -1)), thieu = function (d) { var x = co[d] || {}; return SAPO_FILE.filter(function (f) { return !x[f[0]]; }).map(function (f) { return f[1]; }); };
  var gio = +Utilities.formatDate(new Date(), TZ, 'H');
  var tHn = thieu(hn), tHq = thieu(congNgay_(hn, -1));
  if (gio >= 23 && tHn.length) viec.push({ muc: 'r', loai: 'Số Sapo', tab: 'tong', ngay: hn, text: 'Đã quá 23h, chưa có ' + (tHn.length === 4 ? 'số Sapo' : tHn.join(', ')) + ' ngày ' + sqDM_(hn) + ' – nhắc Sơn gửi' });
  if (tHq.length) viec.push({ muc: tHq.length === 4 ? 'y' : 'i', loai: 'Số Sapo', tab: 'tong', ngay: congNgay_(hn, -1), text: 'Ngày ' + sqDM_(congNgay_(hn, -1)) + ' còn thiếu: ' + tHq.join(', ') });
  try { sqKhoViec_().forEach(function (v) { viec.push(v); }); } catch (e) { }
  try { plKCRows_().filter(function (x) { return x.tt === 'sai' && x.ngay >= congNgay_(hn, -7); }).forEach(function (x) {
    viec.push({ muc: 'r', loai: 'Kế toán báo sai', tab: 'cp', ngay: x.ngay, text: sqDM_(x.ngay) + ' · ' + (x.ten || (/^quay\|/.test(x.khoa) ? 'Chi tại quầy ca ' + x.khoa.split('|')[1] : 'hóa đơn')) + ': ' + [x.loaiSai, x.ghi].filter(String).join(' – ') + ' (' + x.nguoi + ')' }); }); } catch (e) { }
  if (hoiLai) viec.push({ muc: 'i', loai: 'Hóa đơn', tab: 'cp', text: hoiLai + ' hóa đơn anh hỏi lại, nhân viên chưa trả lời' });
  if (choKiem) viec.push({ muc: 'y', loai: 'Hóa đơn', tab: 'cp', text: choKiem + ' hóa đơn ảnh nhân viên chưa kiểm tra / xác nhận' });
  var chuaGia = 0; try { chuaGia = sqGiaHuy_().filter(function (x) { return x.gia === ''; }).length; } catch (e) { }
  if (chuaGia) viec.push({ muc: 'i', loai: 'Hủy hàng', tab: 'huy', text: chuaGia + ' món hủy chưa có giá – anh đặt giá để tính tiền hủy' });

  return { thang: thang, truoc: truoc, homNay: hn, ngay: ngay, viec: viec, lechDo: TN_LECH_DO, sapo: null };
}

var SAPO_FILE = [['tt', 'Thanh toán'], ['mon', 'Mặt hàng'], ['kho', 'Tồn kho'], ['tq', 'Doanh thu tổng quan']];
// ngày nào đã có file nào (từ ngày tu): { 'yyyy-mm-dd': { tt: 'lúc · ai', mon:…, kho:…, tq:… } }
function sqSapoCo_(tu) {
  var m = {};
  sqSapoRows_().forEach(function (x) { if (x.ngay >= tu) (m[x.ngay] = m[x.ngay] || {})[x.loai] = x.luc; });
  try { sqKhoMeta_().forEach(function (x) { if (x.loai === 'kho' && x.ngay >= tu) (m[x.ngay] = m[x.ngay] || {}).kho = x.luc; }); } catch (e) { }
  return m;
}
/* ---------- SỐ SAPO (anh tải file Excel xuất từ Sapo trong Điều hành, trình duyệt đọc file rồi gửi số lên) ---------- */
var DH_SAPO = 'DH_Sapo';
var H_DH_SAPO = ['Ngày', 'Loại', 'Số liệu (JSON)', 'Tải lúc'];
function sqTabSapo_() { return sqTab_(DH_SAPO, H_DH_SAPO, function (sh) { sh.getRange('A:B').setNumberFormat('@'); sh.getRange('D:D').setNumberFormat('@'); }); }
function sqSapoRows_() {
  var sh = sqTabSapo_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 4).getValues().map(function (r, i) { return { row: i + 2, ngay: chuanNgay_(r[0]), loai: String(r[1]), d: sqJ_(r[2], null), luc: String(r[3] || '') }; })
    .filter(function (x) { return x.d && /^\d{4}-\d{2}-\d{2}$/.test(x.ngay); });
}
// p.ds = [{ngay, loai:'tt'|'mon'|'tq'|'kho'|'khoky', d:{...}}]
function sqSapoLuu_(p, nguoi) {
  var ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 70);
  if (!ds.length) throw new Error('Chưa có số liệu.');
  return khoa_(function () {
    var sh = sqTabSapo_(), rows = sqSapoRows_(), idx = {}, luc = bayGio_() + (nguoi ? ' · ' + nguoi : ''), n = 0;
    rows.forEach(function (x) { idx[x.ngay + '|' + x.loai] = x.row; });
    var kho = ds.filter(function (x) { return x.loai === 'kho' || x.loai === 'khoky'; });
    n += sqKhoLuu_(kho, luc);
    ds.filter(function (x) { return kho.indexOf(x) < 0; }).forEach(function (x) {
      var ngay = sqKiemNgay_(x.ngay), loai = ['mon', 'monky', 'monkg', 'tq'].indexOf(x.loai) >= 0 ? x.loai : 'tt', q = x.d || {};
      var d = loai === 'tq'
        ? { tienHang: Math.round(so_(q.tienHang)), giam: Math.round(so_(q.giam)), thue: Math.round(so_(q.thue)), phiDV: Math.round(so_(q.phiDV)), phiGiao: Math.round(so_(q.phiGiao)),
            phiDoiTac: Math.round(so_(q.phiDoiTac)), thueSan: Math.round(so_(q.thueSan)), tip: Math.round(so_(q.tip)), dtThuc: Math.round(so_(q.dtThuc)),
            don: Math.round(so_(q.don)), huy: Math.round(so_(q.huy)), tienHuy: Math.round(so_(q.tienHuy)), sl: so_(q.sl) }
        : loai === 'tt'
        ? { tong: Math.round(so_(x.d && x.d.tong)), don: Math.round(so_(x.d && x.d.don)), kenh: (x.d && x.d.kenh || []).slice(0, 20).map(function (k) { return { ten: chu_(k.ten, 40), don: Math.round(so_(k.don)), huy: Math.round(so_(k.huy)), tien: Math.round(so_(k.tien)) }; }) }
        : { tong: Math.round(so_(x.d && x.d.tong)), giam: Math.round(so_(x.d && x.d.giam)), mon: (x.d && x.d.mon || []).slice(0, 300).map(function (m) { return { dm: chu_(m.dm, 30), ten: chu_(m.ten, 80), sl: so_(m.sl), tien: Math.round(so_(m.tien)) }; }) };
      if (loai === 'monky' || loai === 'monkg') d.den = /^\d{4}-\d{2}-\d{2}$/.test(String(q.den || '')) && String(q.den) >= ngay ? String(q.den) : ngay;
      var row = idx[ngay + '|' + loai];
      if (!row) { row = sh.getLastRow() + 1; idx[ngay + '|' + loai] = row; }
      sh.getRange(row, 1, 1, 4).setValues([[ngay, loai, JSON.stringify(d), luc]]);
      n++;
    });
    return { luu: n };
  });
}

// Nhân viên (Sơn / thu ngân) gửi file Sapo từ app Nhà Gặm – chỉ thấy ngày nào đã gửi, không thấy số tiền
function sqSapoQuyen_(nv) { if ((nv.chucNang || []).indexOf('sapo') < 0) throw new Error('Em chưa được mở chức năng Gửi số Sapo. Nhờ anh Đạt bật trong Điều hành → Nhân viên.'); }
function sqSapoNV_(nv) {
  sqSapoQuyen_(nv);
  return sqSapoDs_(9);
}

// danh sách n ngày gần nhất đã có file nào (anh Đạt + nhân viên)
function sqSapoDs_(n) {
  var hn = homNay_(), m = sqSapoCo_(congNgay_(hn, -n)), ds = [];
  for (var i = 0; i < n; i++) { var d = congNgay_(hn, -i), x = m[d] || {}; ds.push({ ngay: d, tt: x.tt || '', mon: x.mon || '', kho: x.kho || '', tq: x.tq || '' }); }
  return { homNay: hn, ds: ds };
}

/* =====================================================================
 * AI NHẬN XÉT SỐ LIỆU (anh Đạt) – Gemini Flash Lite, dùng chung khóa GEMINI_KEY (khóa miễn phí).
 * Nhận xét viết 1 lần/ngày (lần đầu anh mở Tổng quan), lưu lại; ô hỏi tối đa AI_HOI_MAX câu/ngày.
 * KHÔNG gửi định lượng / công thức món, tên nhân viên, PIN.
 * ===================================================================== */
var AI_MODEL = 'gemini-3.5-flash-lite', AI_HOI_MAX = 20, AI_VIET_LAI_MAX = 3;
var AI_THU = ['CN', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7'];
function aiThu_(d) { var p = d.split('-'); return AI_THU[new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay()]; }
function aiDt_(o) { if (!o) return null; if (o.tq && o.tq.dtThuc) return o.tq.dtThuc; return o.dt != null ? o.dt : null; }
// ngày để nhận xét: hôm qua, nếu chưa có số Sapo thì lùi tối đa 3 ngày
function aiNgayNX_(N, hn) {
  for (var i = 1; i <= 4; i++) { var d = congNgay_(hn, -i); if (aiDt_(N[d]) != null) return d; }
  return congNgay_(hn, -1);
}
function aiDuLieu_(ngay, soNgay) {
  var tq = sqTongQuan_({ thang: ngay.slice(0, 7) }), N = tq.ngay, hn = tq.homNay;
  if (!ngay) ngay = aiNgayNX_(N, hn);
  soNgay = soNgay || 14;
  // chi tại quầy thật: bỏ tiền đi chợ / giao ca (luân chuyển), theo cách anh chỉnh trong Chi phí
  var quay = {}, doi = {}; try { doi = sqQuayDoi_(); } catch (e) { }
  sqTnRows_().forEach(function (x) { (x.chi || []).forEach(function (r) { if (!r.amount) return; var auto = !CP_LUAN_CHUYEN.test(khongDau_(r.note)), k = x.ngay + '|' + x.ca + '|' + khongDau_(r.note);
    if (doi[k] === undefined ? auto : doi[k]) quay[x.ngay] = (quay[x.ngay] || 0) + r.amount; }); });
  var ngayDs = [];
  for (var i = soNgay - 1; i >= 0; i--) {
    var d = congNgay_(ngay, -i), o = N[d] || {};
    var kenh = (o.kenh || []).filter(function (k) { return k.tien; }).map(function (k) { return k.ten + ' ' + k.tien + (k.don ? ' (' + k.don + ' đơn)' : ''); });
    ngayDs.push({ ngay: d, thu: aiThu_(d), doanhThu: aiDt_(o), soDon: (o.tq && o.tq.don) || o.don || null, chiHoaDon: o.chi || 0, chiTaiQuay: quay[d] || 0, huyHang: o.huy || 0,
      phiSanGiaoHang: o.tq ? (o.tq.phiDoiTac || 0) + (o.tq.thueSan || 0) : null, giamGia: o.tq ? o.tq.giam : null, kenh: kenh, chiTheoNhom: o.nhom || {} });
  }
  var cungThu = [7, 14, 21, 28].map(function (k) { var d = congNgay_(ngay, -k); return { ngay: d, doanhThu: aiDt_(N[d]) }; });
  // món bán: ngày nhận xét so trung bình 7 ngày trước đó (số phần)
  var monNgay = {}, tu7 = congNgay_(ngay, -7);
  sqSapoRows_().forEach(function (x) { if (x.loai === 'mon' && x.ngay >= tu7 && x.ngay <= ngay) monNgay[x.ngay] = x.d.mon || []; });
  var hom = {}, tb = {}, soNgayMon = 0;
  (monNgay[ngay] || []).forEach(function (m) { hom[m.ten] = (hom[m.ten] || 0) + so_(m.sl); });
  Object.keys(monNgay).forEach(function (d) { if (d === ngay) return; soNgayMon++; monNgay[d].forEach(function (m) { tb[m.ten] = (tb[m.ten] || 0) + so_(m.sl); }); });
  var mon = Object.keys(hom).concat(Object.keys(tb).filter(function (t) { return !(t in hom); })).map(function (t) {
    return { mon: t, hom: hom[t] || 0, tb7: soNgayMon ? Math.round((tb[t] || 0) / soNgayMon * 10) / 10 : null };
  }).filter(function (m) { return m.hom >= 1 || (m.tb7 || 0) >= 2; }).sort(function (a, b) { return b.hom - a.hom; }).slice(0, 40);
  // lũy kế tháng so cùng kỳ tháng trước
  var lk = function (th, denNgay) { var dt = 0, chi = 0; for (var k = 1; k <= denNgay; k++) { var o = N[th + '-' + ('0' + k).slice(-2)]; if (!o) continue; dt += aiDt_(o) || 0; chi += (o.chi || 0) + (quay[th + '-' + ('0' + k).slice(-2)] || 0); } return { doanhThu: dt, chi: chi }; };
  var nd = +ngay.slice(8, 10);
  var viec = (tq.viec || []).slice(0, 14).map(function (v) { return v.loai + ': ' + v.text; });
  return { ngay: ngay, thu: aiThu_(ngay), homNay: hn, cacNgay: ngayDs, cungThuTuanTruoc: cungThu, monBan: mon, soNgayCoSoMon: soNgayMon,
    luyKeThang: { thangNay: lk(ngay.slice(0, 7), nd), cungKyThangTruoc: lk(tq.truoc, nd) }, viecCanXuLy: viec };
}
function aiGoi_(prompt, schema) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) throw new Error('Chưa cài mã Gemini (GEMINI_KEY).');
  var r = sqGoiGemini_(key, AI_MODEL, [{ text: prompt }], schema);
  var j = JSON.parse(r.body), txt = j.candidates[0].content.parts.map(function (x) { return x.text || ''; }).join('');
  return { kq: sqJ_(txt, null), model: r.model };
}
var AI_LUAT = 'Bạn là trợ lý số liệu của anh Đạt, chủ Tiệm Ăn Gặm (quán ăn Việt ở TP.HCM). Xưng "em", gọi "anh". Tiếng Việt, ngắn, rõ, thực tế.\n' +
  'Chỉ dùng số có trong dữ liệu JSON, không bịa; thiếu số thì nói "chưa có số". Tiền là đồng VN: viết kiểu 18,6 tr / 450k. "doanhThu" đã trừ giảm giá; null = chưa có số Sapo. Không nhắc chữ JSON hay tên trường.\n';
function aiNhanXet_(p) {
  var props = PropertiesService.getScriptProperties(), hn = homNay_();
  var luu = props.getProperty('AI_NX_' + hn), cu = luu ? JSON.parse(luu) : null;
  if (cu && cu.v === 2 && !p.lai) return cu;
  if (cu && p.lai && (cu.lanViet || 1) >= AI_VIET_LAI_MAX) throw new Error('Hôm nay đã viết lại ' + AI_VIET_LAI_MAX + ' lần, mai viết tiếp nhé.');
  var cache = CacheService.getScriptCache();
  if (cache.get('AI_DANG_VIET')) throw new Error('AI đang viết nhận xét, đợi chút rồi tải lại nhé.');
  cache.put('AI_DANG_VIET', '1', 120);
  try {
    var dl = aiDuLieu_('', 14);
    var prompt = AI_LUAT + 'Viết NHẬN XÉT ngày ' + dl.thu + ' ' + dl.ngay.slice(8, 10) + '/' + dl.ngay.slice(5, 7) + ' cho anh: 4–6 dòng, mỗi dòng tối đa 30 chữ, bắt đầu bằng 1 emoji.\n' +
      'Thứ tự: (1) doanh thu, số đơn so với cùng thứ 4 tuần trước và trung bình gần đây; (2) món tăng/giảm rõ so với trung bình 7 ngày (chỉ nêu món chênh nhiều); (3) kênh giao hàng / phí sàn nếu đáng chú ý; (4) chi phí: % chi so doanh thu, nhóm chi tăng, giá nhập tăng, hủy hàng; (5) kho / việc cần xử lý quan trọng nếu có; dòng cuối bắt đầu "👉 Gợi ý:" 1–2 việc nên làm hôm nay.\n' +
      'Dữ liệu:\n' + JSON.stringify(dl);
    var r = aiGoi_(prompt, { type: 'OBJECT', properties: { dong: { type: 'ARRAY', items: { type: 'STRING' } } }, required: ['dong'] });
    var dong = ((r.kq && r.kq.dong) || []).map(function (s) { return String(s).trim().slice(0, 300); }).filter(String).slice(0, 7);
    if (!dong.length) throw new Error('AI chưa viết được, thử lại sau nhé.');
    var o = { ngay: dl.ngay, viet: hn, dong: dong, luc: Utilities.formatDate(new Date(), TZ, 'HH:mm'), model: r.model, v: 2, lanViet: (cu ? (cu.lanViet || 1) : 0) + 1 };
    props.setProperty('AI_NX_' + hn, JSON.stringify(o));
    // dọn bản cũ hơn 10 ngày
    Object.keys(props.getProperties()).forEach(function (k) { var m = /^AI_(NX|HOI)_(\d{4}-\d{2}-\d{2})$/.exec(k); if (m && m[2] < congNgay_(hn, -10)) props.deleteProperty(k); });
    return o;
  } finally { cache.remove('AI_DANG_VIET'); }
}
function aiHoi_(p) {
  var cau = chu_(p.cau, 300); if (!cau) throw new Error('Anh gõ câu hỏi nhé.');
  var props = PropertiesService.getScriptProperties(), hn = homNay_(), k = 'AI_HOI_' + hn, n = +(props.getProperty(k) || 0);
  if (n >= AI_HOI_MAX) throw new Error('Hôm nay đã hỏi ' + AI_HOI_MAX + ' câu (giữ khóa miễn phí), mai hỏi tiếp nhé.');
  props.setProperty(k, String(n + 1));
  var cache = CacheService.getScriptCache(), s = cache.get('AI_DL28'), dl = s ? JSON.parse(s) : null;
  if (!dl) { dl = aiDuLieu_('', 28); try { cache.put('AI_DL28', JSON.stringify(dl), 1800); } catch (e) { } }
  var prompt = AI_LUAT + 'Trả lời CÂU HỎI của anh dựa trên dữ liệu 28 ngày dưới đây (món bán chỉ có ngày gần nhất và trung bình 7 ngày trước đó). Tối đa 120 chữ, có số cụ thể; nếu dữ liệu không đủ để trả lời thì nói rõ cần thêm gì.\n' +
    'Câu hỏi: ' + cau + '\nDữ liệu:\n' + JSON.stringify(dl);
  var r = aiGoi_(prompt, { type: 'OBJECT', properties: { traLoi: { type: 'STRING' } }, required: ['traLoi'] });
  var tl = String((r.kq && r.kq.traLoi) || '').trim();
  if (!tl) throw new Error('AI chưa trả lời được, thử lại nhé.');
  return { cau: cau, traLoi: tl.slice(0, 1500), conLai: AI_HOI_MAX - n - 1 };
}
