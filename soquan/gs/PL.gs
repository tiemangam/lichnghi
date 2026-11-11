/**
 * P&L · LÃI LỖ – Điều hành (anh Đạt).
 *  - Chọn 1 ngày → số từ đầu tháng tới ngày đó: tiền vào (tiền mặt, CK, app giao hàng đã trừ phí), tiền ra theo loại, lời/lỗ,
 *    hòa vốn, cần bán thêm, dự kiến cuối tháng, biểu đồ lời/lỗ cộng dồn, bấm từng khoản xem các ngày đã chi.
 *  - Chi phí cố định theo tháng (anh sửa), chia đều theo ngày; có hóa đơn thật thì lấy số thật.
 *  - Cuối tháng: nhập số thật (bảng lương…) → Tổng hợp tháng, % từng khoản so doanh thu, so tháng trước.
 *  - Anh thêm / sửa / xóa hóa đơn bất kỳ, chọn loại chi. Xóa → "Đã xóa" 30 ngày, lấy lại được.
 * Đầu tư để riêng, không trừ lời/lỗ. Hủy hàng chỉ hiện "trong đó" (đã nằm trong tiền nguyên vật liệu).
 */
var PL_THANG = 'PL_Thang', H_PL_THANG = ['Tháng', 'Chi phí cố định (JSON)', 'Số thật cuối tháng (JSON)', 'Sửa lúc'];
var PL_DT = 'PL_DoanhThu', H_PL_DT = ['Ngày', 'Doanh thu thực', 'Tiền mặt', 'Chuyển khoản', 'Nguồn', 'Lúc'];
var PL_XOA = 'PL_DaXoa', H_PL_XOA = ['ID', 'Ngày HĐ', 'Xóa lúc', 'Tổng', 'Tên', 'Dữ liệu (JSON)', 'Ms'];
var PL_DU_TU = '2026-11';      // từ tháng này app ghi đủ hóa đơn → so sánh chi phí với tháng trước mới đúng
var LOAI_CHU = 'Anh nhập';
var PL_LOAI = {
  nvl: { ten: 'Nguyên vật liệu', icon: '🥩', nhom: 'bd' },
  ga: { ten: 'Ga', icon: '🔥', nhom: 'bd', phieu: 'Ga' },
  banhmi: { ten: 'Bánh mì', icon: '🥖', nhom: 'bd', phieu: 'Bánh mì' },
  quay: { ten: 'Chi tại quầy (đá, vặt)', icon: '🧊', nhom: 'bd' },
  cd: { ten: 'Chi phí cố định', icon: '🏠', nhom: 'cd', phieu: 'Chi phí cố định' },
  luongkhac: { ten: 'Thưởng nóng & chi lương khác', icon: '🎁', nhom: 'khac', phieu: 'Lương / thưởng' },
  sua: { ten: 'Sửa chữa', icon: '🔧', nhom: 'khac', phieu: 'Sửa chữa' },
  khac: { ten: 'Chi khác', icon: '📦', nhom: 'khac', phieu: 'Chi khác' },
  dautu: { ten: 'Đầu tư', icon: '💡', nhom: 'dt', phieu: 'Đầu tư' }
};
var PL_CD_MAC_DINH = [
  { id: 'luong', ten: 'Lương nhân viên', icon: '👥', tien: 0 }, { id: 'nha', ten: 'Tiền nhà', icon: '🏠', tien: 0 },
  { id: 'dien', ten: 'Điện', icon: '⚡', tien: 0 }, { id: 'nuoc', ten: 'Nước', icon: '💧', tien: 0 },
  { id: 'net', ten: 'Internet + Sapo', icon: '🌐', tien: 0 }, { id: 'rac', ten: 'Rác + vệ sinh', icon: '♻️', tien: 0 },
  { id: 'qc', ten: 'Quảng cáo cố định', icon: '📣', tien: 0 }
];
// nhóm phiếu không phải hàng mua → không so giá, không tính vào "chi nguyên liệu" ở Tổng quan
var PL_NHOM_KHONG_HANG = ['Chi phí cố định', 'Đầu tư', 'Lương / thưởng', 'Sửa chữa', 'Chi khác'];

function plTabThang_() { return sqTab_(PL_THANG, H_PL_THANG, function (sh) { sh.getRange('A:A').setNumberFormat('@'); }); }
function plTabDT_() { return sqTab_(PL_DT, H_PL_DT, function (sh) { sh.getRange('A:A').setNumberFormat('@'); }); }
function plTabXoa_() { return sqTab_(PL_XOA, H_PL_XOA, function (sh) { sh.getRange('A:C').setNumberFormat('@'); }); }
function plSoNgay_(th) { var p = th.split('-'); return new Date(Date.UTC(+p[0], +p[1], 0)).getUTCDate(); }
function plThangTruoc_(th) { var y = +th.slice(0, 4), m = +th.slice(5, 7); return m === 1 ? (y - 1) + '-12' : y + '-' + ('0' + (m - 1)).slice(-2); }
function plKiemThang_(s) { if (!/^\d{4}-\d{2}$/.test(String(s || ''))) throw new Error('Tháng không hợp lệ.'); return String(s); }

/* ---------- chi phí cố định + số thật theo tháng ---------- */
function plThangRows_() {
  var sh = plTabThang_(), n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (r, i) {
    var th = String(r[0]).trim(); if (r[0] instanceof Date) th = Utilities.formatDate(r[0], TZ, 'yyyy-MM');
    if (/^\d{4}-\d{2}$/.test(th)) m[th] = { cd: sqJ_(r[1], null), that: sqJ_(r[2], {}) || {}, luc: String(r[3] || ''), row: i + 2 };
  });
  return m;
}
// bảng cố định của tháng: tháng đó chưa có thì lấy tháng gần nhất trước đó (chép), chưa có gì thì mặc định
function plCD_(rows, th) {
  var x = rows[th];
  if (x && x.cd) return { ds: x.cd, nguon: 'thang', that: x.that || {}, luc: x.luc };
  var ks = Object.keys(rows).filter(function (k) { return k < th && rows[k].cd; }).sort();
  if (ks.length) return { ds: rows[ks[ks.length - 1]].cd, nguon: 'chep', tu: ks[ks.length - 1], that: (x && x.that) || {} };
  return { ds: PL_CD_MAC_DINH, nguon: 'macdinh', that: (x && x.that) || {} };
}
function plGhiThang_(th, cd, that) {
  return khoa_(function () {
    var sh = plTabThang_(), rows = plThangRows_(), x = rows[th];
    var cu = x || { cd: null, that: {} };
    var row = [th, JSON.stringify(cd === undefined ? cu.cd : cd), JSON.stringify(that === undefined ? (cu.that || {}) : that), bayGio_()];
    if (row[1] === 'null') row[1] = '';
    if (x) sh.getRange(x.row, 1, 1, 4).setValues([row]); else sh.getRange(sh.getLastRow() + 1, 1, 1, 4).setValues([row]);
    return true;
  });
}
function plCDLuu_(p) {
  var th = plKiemThang_(p.thang);
  var dung = {}, ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 30).map(function (x, i) {
    var id = String(x.id || '').replace(/[^a-z0-9_]/gi, '').slice(0, 20) || ('k' + Date.now().toString(36) + i);
    while (dung[id]) id += 'x'; dung[id] = 1;
    return { id: id, ten: chu_(x.ten, 40) || 'Khoản cố định', icon: chu_(x.icon, 4) || '📌', tien: Math.max(0, Math.round(so_(x.tien))) };
  }).filter(function (x) { return x.ten; });
  plGhiThang_(th, ds, undefined);
  return plCDXem_({ thang: th });
}
// số thật cuối tháng: { dt, luong, thuongNong, nvl, ga, banhmi, quay, sua, khac, <id khoản cố định>: số }
function plThatLuu_(p) {
  var th = plKiemThang_(p.thang), v = p.that || {}, o = {};
  Object.keys(v).slice(0, 40).forEach(function (k) {
    if (!/^[a-z0-9_]{1,20}$/i.test(k)) return;
    if (v[k] === '' || v[k] === null || v[k] === undefined) return;
    var n = Math.round(so_(v[k])); if (n >= 0) o[k] = n;
  });
  if (p.ghi) o._ghi = chu_(p.ghi, 200);
  o._luc = bayGio_();
  plGhiThang_(th, undefined, o);
  return plCDXem_({ thang: th });
}
function plCDXem_(p) {
  var th = plKiemThang_(p.thang || homNay_().slice(0, 7)), rows = plThangRows_(), c = plCD_(rows, th);
  // hóa đơn thật đã nhập cho từng khoản trong tháng
  var hd = {};
  plPhieuThang_(th + '-01', th + '-31').forEach(function (x) {
    x.parts.forEach(function (p) { if (p.k !== 'cd') return; var id = p.khoan || ('hd:' + khongDau_(x.ten).slice(0, 18)); hd[id] = (hd[id] || 0) + p.tien; });
  });
  return { thang: th, ds: c.ds, nguon: c.nguon, tu: c.tu || '', that: c.that, hd: hd, truoc: plCD_(rows, plThangTruoc_(th)).ds };
}

/* ---------- doanh thu nhập tay / từ bảng lương (tháng cũ chưa có Sapo) ---------- */
function plDTRows_() {
  var sh = plTabDT_(), n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 6).getValues().forEach(function (r, i) {
    var d = chuanNgay_(r[0]); if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
    m[d] = { dt: soHoacRong_(r[1]), tm: soHoacRong_(r[2]), ck: soHoacRong_(r[3]), nguon: String(r[4] || ''), row: i + 2 };
  });
  return m;
}
function plDTNhap_(p) {
  var ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 400);
  return khoa_(function () {
    var sh = plTabDT_(), m = plDTRows_(), n = 0, luc = bayGio_();
    ds.forEach(function (x) {
      var d = sqKiemNgay_(x.ngay), cu = m[d] || {};
      var v = function (k) { return x[k] === undefined ? (cu[k] === undefined ? '' : cu[k]) : soHoacRong_(x[k]); };
      var row = [d, v('dt'), v('tm'), v('ck'), chu_(x.nguon, 60) || cu.nguon || 'nhập tay', luc];
      if (cu.row) sh.getRange(cu.row, 1, 1, 6).setValues([row]);
      else { var r = sh.getLastRow() + 1; sh.getRange(r, 1, 1, 6).setValues([row]); m[d] = { row: r }; }
      n++;
    });
    return { luu: n };
  });
}

/* ---------- phân loại chi ---------- */
function plTok_(s) { return String(s || '').toLowerCase().normalize('NFC').split(/[^0-9a-zà-ỹđ]+/); }
function plLaGa_(s) { var t = plTok_(s); return t.indexOf('ga') >= 0 || t.indexOf('gas') >= 0; }
function plLaBanhMi_(s) { return /(^|[^a-z])banh mi([^a-z]|$)/.test(khongDau_(s)); }
// tên mối / nhóm → loại chi (khi anh chưa chọn loại)
var PL_TU_KHOA = [   // so trên chữ CÓ DẤU (vd "sữa chua" ≠ "sửa chữa")
  [/ đầu tư /, 'dautu'], [/ cấp nước | nước sinh hoạt | sawaco | tiền nước /, 'cd', 'nuoc'], [/ điện lực | evn | tiền điện /, 'cd', 'dien'],
  [/ tiền nhà | thuê nhà | mặt bằng /, 'cd', 'nha'], [/ internet | fpt | vnpt | cước mạng /, 'cd', 'net'], [/ thu gom rác | tiền rác | vệ sinh môi trường /, 'cd', 'rac'],
  [/ sửa chữa | thợ sửa | sửa máy /, 'sua']
];
function plLoaiTuNhom_(nhom) {
  for (var k in PL_LOAI) if (PL_LOAI[k].phieu && PL_LOAI[k].phieu === nhom) return k;
  return '';
}
// mỗi phiếu → các phần theo loại chi: [{k, khoan, tien, dong:[...]}]
function plTach_(x) {
  var them = (x.meta && x.meta.them) || {};
  var k0 = them.loaiChi || plLoaiTuNhom_(x.nhom);
  if (k0 === 'quen') k0 = 'nvl';
  if (k0) return [{ k: k0, khoan: them.khoan || '', tien: x.tong, dong: x.mon }];
  var dau = (x.moi || '') + ' ' + (x.nhom || '');
  var dk = ' ' + String(dau).toLowerCase().normalize('NFC').replace(/[^0-9a-zà-ỹđ]+/g, ' ').trim() + ' ';
  for (var i = 0; i < PL_TU_KHOA.length; i++) if (PL_TU_KHOA[i][0].test(dk)) return [{ k: PL_TU_KHOA[i][1], khoan: PL_TU_KHOA[i][2] || '', tien: x.tong, dong: x.mon }];
  if (plLaGa_(dau)) return [{ k: 'ga', tien: x.tong, dong: x.mon }];
  if (plLaBanhMi_(dau)) return [{ k: 'banhmi', tien: x.tong, dong: x.mon }];
  var m = {}, out = [];
  (x.mon || []).forEach(function (d) {
    var k = plLaGa_(d.ten) ? 'ga' : plLaBanhMi_(d.ten) ? 'banhmi' : 'nvl';
    if (!m[k]) { m[k] = { k: k, tien: 0, dong: [] }; out.push(m[k]); }
    m[k].tien += so_(d.tien); m[k].dong.push(d);
  });
  if (!out.length) out.push({ k: 'nvl', tien: x.tong, dong: [] });
  return out;
}
// loại chính (để hiện nhãn trong danh sách hóa đơn)
function plLoaiChinh_(x) {
  var ps = plTach_(x).sort(function (a, b) { return b.tien - a.tien; });
  return { k: ps[0].k, khoan: ps[0].khoan || '', ghep: ps.length > 1, cac: ps.map(function (p) { return { k: p.k, tien: p.tien }; }) };
}
function plQuayLoai_(note) {
  if (plLaGa_(note)) return 'ga';
  if (plLaBanhMi_(note)) return 'banhmi';
  if (/thuong nong|thuong nv|tien thuong/.test(khongDau_(note))) return 'luongkhac';
  return 'quay';
}

// phiếu (hóa đơn, đi chợ, anh nhập) trong khoảng ngày kèm loại chi
function plPhieuThang_(tu, den) {
  var metas = {}; sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  return sqPhieuTrong_(tu, den).filter(function (x) { return x.loai !== LOAI.BILL; }).map(function (x) {
    x.meta = metas[x.id] || null;
    var lc = plLoaiChinh_(x);
    var ten = (x.loai === LOAI.HD || x.loai === LOAI.TD) ? (x.moi || x.nhom.replace(/^HĐ /, '')) : (x.loai === LOAI_CHU ? (x.moi || x.nhom) : x.nhom);
    return { id: x.id, ngay: x.ngay, loai: x.loai, nhom: x.nhom, moi: x.moi, nguoi: x.nguoi, tong: x.tong, mon: x.mon, anh: (x.anh || []).length, ten: ten, lc: lc, parts: plTach_(x), meta: x.meta };
  });
}

/* ---------- doanh thu theo ngày ---------- */
var PL_APP = /grab|shopee|(^| )be( |$)|bill be|baemin|gojek|xanh sm/;
function plNapDT_(tu, den) {
  var sapo = {}, tn = {}, nhap = plDTRows_(), out = {};
  sqSapoRows_().forEach(function (x) { if (x.ngay >= tu && x.ngay <= den && (x.loai === 'tq' || x.loai === 'tt')) (sapo[x.ngay] = sapo[x.ngay] || {})[x.loai] = x.d; });
  sqTnRows_().forEach(function (x) { if (x.ngay < tu || x.ngay > den) return; var o = tn[x.ngay] = tn[x.ngay] || { dttm: 0, ck: 0 }; o.dttm += x.dttm; o.ck += x.tr; });
  for (var d = tu; d <= den; d = congNgay_(d, 1)) {
    var s = sapo[d] || {}, t = tn[d], n = nhap[d] || {}, R = null, nguon = '';
    var phi = s.tq ? so_(s.tq.phiDoiTac) + so_(s.tq.thueSan) : 0;
    if (s.tq) { R = so_(s.tq.dtThuc) - phi; nguon = 'sapo'; }
    else if (n.dt !== undefined && n.dt !== '') { R = n.dt; nguon = 'nhap'; }
    else if (s.tt) { R = so_(s.tt.tong); nguon = 'tt'; }
    else if (t) { R = t.dttm; nguon = 'tn'; }
    else if (n.tm !== undefined && n.tm !== '') { R = so_(n.tm) + so_(n.ck); nguon = 'tncu'; }
    if (R === null) { out[d] = null; continue; }
    var ck = t ? t.ck : (n.ck !== '' && n.ck !== undefined ? so_(n.ck) : 0), app = 0, appGop = 0, tm = 0, chua = 0;
    if (s.tt) {
      (s.tt.kenh || []).forEach(function (k) { if (PL_APP.test(khongDau_(k.ten))) appGop += so_(k.tien); });
      app = Math.max(0, appGop - phi);
      tm = R - app - ck;
    } else if (t) tm = t.dttm - ck;
    else if (n.tm !== undefined && n.tm !== '') tm = so_(n.tm);
    chua = R - tm - ck - app;
    if (chua < 0) { tm += chua; chua = 0; }
    if (tm < 0) { ck += tm; tm = 0; if (ck < 0) ck = 0; }
    out[d] = { R: Math.round(R), tm: Math.round(tm), ck: Math.round(ck), app: Math.round(app), appGop: Math.round(appGop), phi: Math.round(phi), chua: Math.round(chua), nguon: nguon, coTT: !!s.tt };
  }
  return out;
}

/* ---------- tính P&L 1 tháng tới ngày n ---------- */
// data nạp sẵn cho cả khoảng (tháng trước + tháng này) để chỉ đọc sheet 1 lần
function plNap_(tu, den) {
  var doi = sqQuayDoi_(), quay = [];
  sqTnRows_().forEach(function (x) {
    if (x.ngay < tu || x.ngay > den) return;
    x.chi.forEach(function (r) {
      if (!r.amount) return;
      var auto = !CP_LUAN_CHUYEN.test(khongDau_(r.note)), k = x.ngay + '|' + x.ca + '|' + khongDau_(r.note);
      var tinh = doi[k] === undefined ? auto : doi[k];
      if (tinh) quay.push({ ngay: x.ngay, ca: x.ca, nguoi: x.nguoi, note: r.note, tien: r.amount, k: plQuayLoai_(r.note) });
    });
  });
  var huy = sqHuyRows_().filter(function (h) { return h.ngay >= tu && h.ngay <= den; });
  return { phieu: plPhieuThang_(tu, den), quay: quay, huy: huy, dt: plNapDT_(tu, den), rows: plThangRows_() };
}
function plTinh_(N, th, n, ngayCuoi) {
  var nd = plSoNgay_(th), tu = th + '-01', den = th + '-' + ('0' + n).slice(-2), full = n >= nd;
  var cd = plCD_(N.rows, th), that = cd.that || {};
  var T = function (k) { return full && that[k] !== undefined ? that[k] : null; };
  // tiền vào
  var vao = { R: 0, tm: 0, ck: 0, app: 0, appGop: 0, phi: 0, chua: 0, soNgay: 0, thieu: [], nguon: {} }, ngay = [];
  for (var i = 1; i <= n; i++) {
    var d = th + '-' + ('0' + i).slice(-2), o = N.dt[d];
    ngay.push(o ? { d: d, R: o.R, tm: o.tm, ck: o.ck, app: o.app, phi: o.phi, chua: o.chua, nguon: o.nguon, bd: 0, khac: 0 } : { d: d, R: null, bd: 0, khac: 0 });
    if (!o) { if (!ngayCuoi || d <= ngayCuoi) vao.thieu.push(d); continue; }
    ['R', 'tm', 'ck', 'app', 'appGop', 'phi', 'chua'].forEach(function (k) { vao[k] += o[k]; });
    vao.soNgay++; vao.nguon[o.nguon] = (vao.nguon[o.nguon] || 0) + 1;
  }
  var dtThat = T('dt');
  if (dtThat !== null) { vao.thatDT = dtThat; vao.Rapp = vao.R; vao.R = dtThat; }
  var nIdx = function (d) { return +d.slice(8) - 1; };
  // tiền ra theo loại
  var L = {}; ['nvl', 'ga', 'banhmi', 'quay', 'luongkhac', 'sua', 'khac', 'dautu'].forEach(function (k) { L[k] = { tien: 0, ct: [] }; });
  var cdHD = {}, cdLe = [];
  N.phieu.forEach(function (x) {
    if (x.ngay < tu || x.ngay > den) return;
    x.parts.forEach(function (p) {
      var ten = x.ten, nd2 = (p.dong || []).slice(0, 5).map(function (dd) { return dd.ten + (dd.sl !== '' && dd.sl != null ? ' ' + String(dd.sl).replace('.', ',') + (dd.dv ? ' ' + dd.dv : '') : ''); }).join(' · ') + ((p.dong || []).length > 5 ? ' …' : '');
      var dong = { ngay: x.ngay, id: x.id, ten: ten, nguoi: x.nguoi, nd: nd2, tien: p.tien, anh: x.anh, loai: x.loai };
      if (p.k === 'cd') {
        var id = p.khoan || ('hd:' + khongDau_(ten).slice(0, 18));
        if (!cdHD[id]) cdHD[id] = { tien: 0, ct: [], ten: ten };
        cdHD[id].tien += p.tien; cdHD[id].ct.push(dong);
        return;
      }
      var k = L[p.k] ? p.k : 'khac';
      L[k].tien += p.tien; L[k].ct.push(dong);
      if (k !== 'dautu') { var g = ngay[nIdx(x.ngay)]; if (g) { if (PL_LOAI[k].nhom === 'bd') g.bd += p.tien; else g.khac += p.tien; } }
    });
  });
  N.quay.forEach(function (q) {
    if (q.ngay < tu || q.ngay > den) return;
    L[q.k].tien += q.tien; L[q.k].ct.push({ ngay: q.ngay, ten: 'Chi tại quầy · Ca ' + q.ca, nguoi: q.nguoi, nd: q.note, tien: q.tien, quay: 1 });
    var g = ngay[nIdx(q.ngay)]; if (g) { if (PL_LOAI[q.k].nhom === 'bd') g.bd += q.tien; else g.khac += q.tien; }
  });
  // hủy hàng (chỉ để biết, đã nằm trong NVL)
  var huy = { tien: 0, chuaGia: 0, ct: [] };
  N.huy.forEach(function (h) {
    if (h.ngay < tu || h.ngay > den) return;
    if (h.tien === '') huy.chuaGia++; else huy.tien += h.tien;
    huy.ct.push({ ngay: h.ngay, ten: h.ten, nd: String(h.sl).replace('.', ',') + ' ' + h.dv + (h.lyDo ? ' · ' + h.lyDo : ''), nguoi: h.nguoi, tien: h.tien === '' ? null : h.tien, anh: (h.anh || []).length, ca: h.ca });
  });
  // số thật cuối tháng thay số app (khi xem cả tháng)
  ['nvl', 'ga', 'banhmi', 'quay', 'sua', 'khac'].forEach(function (k) { var t = T(k); if (t !== null) { L[k].app = L[k].tien; L[k].tien = t; L[k].that = true; } });
  var tN = T('thuongNong'); if (tN !== null) { L.luongkhac.app = L.luongkhac.tien; L.luongkhac.tien = tN; L.luongkhac.that = true; }
  // chi phí cố định: số thật > hóa đơn thật > ước tính; chia theo ngày
  var heSo = n / nd, cdRows = [], F = 0;
  cd.ds.forEach(function (c) {
    var hd = cdHD[c.id], th2 = that[c.id], thang, ghi, nguon;
    if (th2 !== undefined) { thang = th2; nguon = 'that'; }
    else if (hd && hd.tien) { thang = hd.tien; nguon = 'hd'; }
    else { thang = c.tien; nguon = 'uoc'; }
    F += thang;
    var tien = Math.round(thang * heSo);
    ghi = nguon === 'that' ? 'số thật cuối tháng' : nguon === 'hd' ? 'hóa đơn thật ' + sqTr_(hd.tien) : (thang ? sqTr_(thang) + '/tháng' : 'chưa nhập số');
    if (!full) ghi += ' × ' + n + '/' + nd;
    var ct = [{ ngay: '', ten: nguon === 'uoc' ? 'Ước tính mỗi tháng' : nguon === 'that' ? 'Số thật cuối tháng' : 'Hóa đơn thật trong tháng', nd: full ? '' : 'chia đều ' + n + '/' + nd + ' ngày', tien: thang, uoc: 1 }].concat(hd ? hd.ct : []);
    cdRows.push({ k: 'cd:' + c.id, id: c.id, ten: c.ten, icon: c.icon, nhom: 'cd', tien: tien, thang: thang, nguon: nguon, ghi: ghi, ct: ct, chuaNhap: !thang });
    if (hd) delete cdHD[c.id];
  });
  Object.keys(cdHD).forEach(function (id) {   // hóa đơn cố định không thuộc khoản nào trong bảng
    var h = cdHD[id]; F += h.tien;
    cdRows.push({ k: 'cd:' + id, id: id, ten: h.ten, icon: '📌', nhom: 'cd', tien: Math.round(h.tien * heSo), thang: h.tien, nguon: 'hd', ghi: 'hóa đơn ' + sqTr_(h.tien) + (full ? '' : ' × ' + n + '/' + nd), ct: h.ct });
  });
  // dòng P&L
  var R = vao.R, pc = function (v) { return R > 0 ? v / R : null; };
  var mk = function (k) { var x = L[k], m = PL_LOAI[k]; return { k: k, ten: m.ten, icon: m.icon, nhom: m.nhom, tien: Math.round(x.tien), ct: x.ct.sort(function (a, b) { return a.ngay < b.ngay ? 1 : a.ngay > b.ngay ? -1 : b.tien - a.tien; }), that: !!x.that, app: x.app }; };
  var ra = [mk('nvl'), mk('ga'), mk('banhmi'), mk('quay')];
  var luong = cdRows.filter(function (c) { return c.id === 'luong'; });
  ra = ra.concat(luong).concat(cdRows.filter(function (c) { return c.id !== 'luong'; })).concat([mk('luongkhac'), mk('sua'), mk('khac')]);
  ra.forEach(function (r) { r.pct = pc(r.tien); });
  var tong = ra.reduce(function (s, r) { return s + r.tien; }, 0);
  var bien = ra.filter(function (r) { return r.nhom === 'bd'; }).reduce(function (s, r) { return s + r.tien; }, 0);
  var khac = ra.filter(function (r) { return r.nhom === 'khac'; }).reduce(function (s, r) { return s + r.tien; }, 0);
  var cdT = tong - bien - khac;
  var loi = R - tong;
  // hòa vốn
  var v = R > 0 ? bien / R : 0, hv = null;
  if (R > 0 && v < 1) {
    var BE = F / (1 - v), dn = vao.soNgay || n, tb = R / Math.max(1, dn), con = nd - n, can = Math.max(0, BE - R);
    var duKien = full ? R : R + tb * con;
    var khacThang = full ? khac : khac / Math.max(1, n) * nd;
    hv = { F: Math.round(F), v: v, BE: Math.round(BE), BEngay: Math.round(BE / nd), dat: R / BE, can: Math.round(can), conNgay: con, canNgay: con > 0 ? Math.round(can / con) : null,
      tb: Math.round(tb), duKien: Math.round(duKien), loiDuKien: Math.round(duKien * (1 - v) - F - khacThang), ngayHV: tb > 0 ? Math.ceil(BE / tb) : null };
  }
  // lời/lỗ cộng dồn từng ngày (cố định chia đều mỗi ngày)
  var luy = 0, gop = 0, fNgay = F / nd, ngayHVthat = null;
  ngay.forEach(function (g) { var l = (g.R || 0) - g.bd - g.khac - fNgay; g.loi = Math.round(l); luy += l; g.luy = Math.round(luy); gop += (g.R || 0) - g.bd - g.khac; if (ngayHVthat === null && F > 0 && gop >= F) ngayHVthat = g.d; });
  var dauTu = mk('dautu');
  return { thang: th, n: n, nd: nd, den: den, full: full, vao: vao, ra: ra, tong: Math.round(tong), bien: Math.round(bien), cd: Math.round(cdT), khac: Math.round(khac), loi: Math.round(loi), bienLoi: pc(loi),
    hv: hv, ngay: ngay, ngayHV: ngayHVthat, huy: { tien: Math.round(huy.tien), chuaGia: huy.chuaGia, ct: huy.ct.sort(function (a, b) { return a.ngay < b.ngay ? 1 : -1; }), pctNVL: L.nvl.tien ? huy.tien / L.nvl.tien : null, pct: pc(huy.tien) },
    dauTu: { tien: dauTu.tien, ct: dauTu.ct }, cdNguon: cd.nguon, cdTu: cd.tu || '', that: that, du: th >= PL_DU_TU };
}
function sqTr_(n) { var a = Math.abs(n); return a >= 1e6 ? (Math.round(n / 1e5) / 10).toString().replace('.', ',') + ' tr' : a >= 1e3 ? Math.round(n / 1e3) + 'k' : Math.round(n) + 'đ'; }
function plGon_(r) { // bỏ chi tiết để so sánh tháng trước nhẹ hơn
  return { thang: r.thang, n: r.n, nd: r.nd, R: r.vao.R, tong: r.tong, loi: r.loi, bien: r.bien, cd: r.cd, khac: r.khac, du: r.du, huy: r.huy.tien,
    ra: r.ra.map(function (x) { return { k: x.k, ten: x.ten, tien: x.tien, pct: x.pct }; }), vao: { tm: r.vao.tm, ck: r.vao.ck, app: r.vao.app, chua: r.vao.chua } };
}

// P&L tháng của ngày đã chọn, từ đầu tháng tới ngày đó + so cùng kỳ tháng trước
function plXem_(p) {
  plApTenLoai_();
  var hn = homNay_(), ngay = /^\d{4}-\d{2}-\d{2}$/.test(String(p.ngay || '')) ? p.ngay : hn;
  if (ngay > hn) ngay = hn;
  var th = ngay.slice(0, 7), n = +ngay.slice(8), tr = plThangTruoc_(th), ndTr = plSoNgay_(tr);
  var N = plNap_(tr + '-01', ngay);
  var a = plTinh_(N, th, n, hn), b = plTinh_(N, tr, Math.min(n, ndTr), hn);
  a.truoc = plGon_(b);
  a.ngayChon = ngay; a.homNay = hn; a.loaiTen = plTenLoai_();
  return a;
}
// Tổng hợp tháng (cả tháng, số thật nếu có) + tháng trước
function plTongHop_(p) {
  plApTenLoai_();
  var hn = homNay_(), th = plKiemThang_(p.thang || hn.slice(0, 7)), tr = plThangTruoc_(th), nd = plSoNgay_(th);
  var cuoi = th === hn.slice(0, 7) ? hn : th + '-' + nd;
  if (th > hn.slice(0, 7)) throw new Error('Tháng chưa tới.');
  var N = plNap_(tr + '-01', cuoi);
  var a = plTinh_(N, th, +cuoi.slice(8), hn), b = plTinh_(N, tr, plSoNgay_(tr), hn);
  a.homNay = hn; a.dangChay = cuoi < th + '-' + nd;
  a.truoc = plGon_(b); a.truoc.huyPctNVL = b.huy.pctNVL; a.truoc.that = b.that;
  return a;
}

/* ---------- anh Đạt thêm / sửa / xóa hóa đơn ---------- */
function plAnhLuu_(ds, id) {
  var folder = null, stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
  return (Array.isArray(ds) ? ds : []).slice(0, 4).filter(function (a) { return a && a.b64; }).map(function (a, j) {
    folder = folder || sqThuMuc_();
    return folder.createFile(Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'image/jpeg', 'CHU-' + id + '-' + stamp + (j ? '-' + (j + 1) : '') + '.jpg')).getId();
  });
}
function plKiemLoai_(k) { if (k !== 'quen' && !PL_LOAI[k]) throw new Error('Chưa chọn loại chi.'); if (k === 'quay') throw new Error('Chi tại quầy lấy từ sổ thu ngân.'); return k; }
function plThem_(p) {
  var id = p.id ? sqKiemId_(p.id) : 'A' + Date.now().toString(36).toUpperCase() + Math.floor(Math.random() * 1296).toString(36).toUpperCase();
  var ngay = sqKiemNgay_(p.ngay), k = plKiemLoai_(String(p.loaiChi || '')), moi = chu_(p.moi, 60), ghi = chu_(p.ghiChu, 200);
  var mon = sqMon_(p.mon);
  if (!mon.length) {
    var tien = Math.round(so_(p.tien));
    if (!(tien > 0)) throw new Error('Chưa nhập số tiền.');
    mon = [{ ten: chu_(p.ten, 80) || (PL_LOAI[k] ? PL_LOAI[k].ten : 'Chi'), sl: '', dv: '', gia: '', tien: tien }];
  }
  var tong = mon.reduce(function (s, d) { return s + d.tien; }, 0);
  var loai = k === 'nvl' ? LOAI.HD : k === 'quen' ? LOAI.TD : LOAI_CHU;
  if ((k === 'nvl' || k === 'quen') && !moi) moi = 'Anh Đạt mua';
  var nhom = loai === LOAI.HD ? 'HĐ ' + moi : loai === LOAI.TD ? NHOM_TU_DAT : PL_LOAI[k].phieu;
  return khoa_(function () {
    if (sqTimMeta_(id) || docPhieu_(function (x) { return x.id === id; }).length) throw new Error('Mã hóa đơn bị trùng, bấm Lưu lại.');
    var anh = plAnhLuu_(p.anh, id);
    sqApiLuu_({ id: id, ngay: ngay, nhom: nhom, nguoiMua: 'anh Đạt', ghiChu: [loai === LOAI_CHU && moi ? moi : '', ghi].filter(String).join(' · '),
      mon: mon.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) });
    var thumb = /^[A-Za-z0-9+/=]+$/.test(p.thumb || '') && p.thumb.length < 16000 ? p.thumb : '';
    var quenId = '';
    if (p.luuQuen && chu_(p.luuQuen.ten, 40)) {
      try { quenId = sqQuenLuu_({ ten: 'anh Đạt' }, { ten: p.luuQuen.ten, icon: p.luuQuen.icon, moi: moi || (PL_LOAI[k] ? PL_LOAI[k].ten : ''), mon: mon.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia !== '' ? d.gia : d.tien }; }), traLien: false }).quen.id; } catch (e) { }
    }
    sqGhiMeta_({ id: id, loai: loai, ngay: ngay, nguoi: 'anh Đạt', moi: moi, anh: anh, dong: mon, cong: tong, ghiChu: ghi, tt: TT.OK, taoLuc: bayGio_(),
      them: { loaiChi: k === 'quen' ? '' : k, khoan: k === 'cd' ? chu_(p.khoan, 20) : '', chuTao: bayGio_(), thumb: thumb, quenId: k === 'quen' ? chu_(p.quenId, 30) : quenId }, tong: tong });
    return { id: id, tong: tong, quenId: quenId };
  });
}
// sửa: ngày, loại chi, mối, ghi chú, các dòng, thêm/bỏ ảnh
function plSua_(p) {
  var id = sqKiemId_(p.id), mon = sqMon_(p.mon);
  if (!mon.length) throw new Error('Hóa đơn phải còn ít nhất 1 dòng.');
  return khoa_(function () {
    var ph = docPhieu_(function (x) { return x.id === id; })[0];
    if (!ph) throw new Error('Không tìm thấy hóa đơn (hóa đơn nhân viên chưa xác nhận thì chưa sửa ở đây được).');
    var m = sqTimMeta_(id) || { id: id, loai: LOAI.DC, ngay: ph.ngay, nguoi: ph.nguoiMua, moi: '', anh: [], dong: [], tt: TT.OK, taoLuc: bayGio_(), them: {} };
    var ngay = p.ngay ? sqKiemNgay_(p.ngay) : ph.ngay;
    var them = Object.assign({}, m.them || {});
    if (p.loaiChi !== undefined) {
      var k = String(p.loaiChi || '');
      if (k) { plKiemLoai_(k); them.loaiChi = k === 'quen' ? '' : k; them.khoan = k === 'cd' ? chu_(p.khoan, 20) : ''; }
      else { delete them.loaiChi; delete them.khoan; }
    }
    var moi = p.moi !== undefined ? chu_(p.moi, 60) : m.moi;
    var nhom = ph.nhom;
    if (m.loai === LOAI.HD) nhom = 'HĐ ' + (moi || 'Chưa rõ mối');
    else if (m.loai === LOAI_CHU && them.loaiChi && PL_LOAI[them.loaiChi] && PL_LOAI[them.loaiChi].phieu) nhom = PL_LOAI[them.loaiChi].phieu;
    var ghi = p.ghiChu !== undefined ? chu_(p.ghiChu, 200) : ph.ghiChu;
    sqApiLuu_({ id: id, ngay: ngay, nhom: nhom, nguoiMua: ph.nguoiMua, tienNhan: ph.tienNhan, ghiChu: ghi,
      mon: mon.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) });
    var tong = mon.reduce(function (s, d) { return s + d.tien; }, 0), cu = m.dong || [];
    m.dong = mon.map(function (d, i) { var c = cu[i] || {}; return Object.assign({}, c, { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }); });
    var bo = Array.isArray(p.anhBo) ? p.anhBo.map(String) : [];
    m.anh = (m.anh || []).filter(function (f) { return bo.indexOf(f) < 0; }).concat(plAnhLuu_(p.anhThem, id));
    m.ngay = ngay; m.moi = moi; m.cong = tong; m.tong = tong;
    if (m.loai === LOAI.DC) m.moi = moi || '';
    if (p.ghiChu !== undefined && m.loai !== LOAI.DC) m.ghiChu = ghi;
    them.chuSua = bayGio_();
    m.them = them;
    sqGhiMeta_(m);
    return { id: id, tong: tong, ngay: ngay };
  });
}
function plXoa_(p) {
  var id = sqKiemId_(p.id);
  return khoa_(function () {
    var ss = layFile_(), shP = ss.getSheetByName('Phieu'), shC = ss.getSheetByName('ChiTiet'), shM = ss.getSheetByName(SQ_META);
    var rP = timDong_(shP, id), phieu = rP > 0 ? shP.getRange(rP, 1, 1, H_PHIEU.length).getValues()[0] : null;
    var ct = [], nC = shC.getLastRow();
    if (nC >= 2) shC.getRange(2, 1, nC - 1, H_CT.length).getValues().forEach(function (r) { if (String(r[0]) === id) ct.push(r); });
    var m = sqTimMeta_(id), meta = m ? shM.getRange(m._row, 1, 1, H_SQ_META.length).getValues()[0] : null;
    if (!phieu && !meta) throw new Error('Không tìm thấy hóa đơn.');
    var nen = function (v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : v; };
    var data = { phieu: phieu ? phieu.map(nen) : null, ct: ct.map(function (r) { return r.map(nen); }), meta: meta ? meta.map(nen) : null };
    var ngay = phieu ? chuanNgay_(phieu[1]) : m.ngay, tong = phieu ? so_(phieu[5]) : (m.tong || m.cong);
    var ten = m ? (m.moi || m.loai) : String(phieu[2]);
    plTabXoa_().appendRow([id, ngay, bayGio_(), tong, ten, JSON.stringify(data), Date.now()]);
    if (rP > 0) shP.deleteRow(rP);
    xoaChiTiet_(ss, id);
    if (m) { shM.deleteRow(m._row); sqMetaMoi_(); }
    PropertiesService.getScriptProperties().setProperty('SQ_TONGHOP_CU', '1');
    return { id: id };
  });
}
function plDaXoa_() {
  var sh = plTabXoa_(), n = sh.getLastRow(), han = Date.now() - 30 * 864e5, ds = [], cu = [];
  if (n < 2) return { ds: [] };
  sh.getRange(2, 1, n - 1, H_PL_XOA.length).getValues().forEach(function (r, i) {
    if (so_(r[6]) && so_(r[6]) < han) { cu.push(i + 2); return; }
    var d = sqJ_(r[5], {}), mon = (d.ct || []).map(function (c) { return String(c[4]); });
    ds.push({ id: String(r[0]), ngay: chuanNgay_(r[1]), xoaLuc: String(r[2]), tong: so_(r[3]), ten: String(r[4]), mon: mon.slice(0, 6).join(', ') + (mon.length > 6 ? '…' : '') });
  });
  if (cu.length) khoa_(function () { cu.reverse().forEach(function (r) { sh.deleteRow(r); }); });
  return { ds: ds.reverse() };
}
function plKhoiPhuc_(p) {
  var id = sqKiemId_(p.id);
  return khoa_(function () {
    var sh = plTabXoa_(), n = sh.getLastRow(), row = -1, data = null;
    if (n >= 2) sh.getRange(2, 1, n - 1, 6).getValues().forEach(function (r, i) { if (String(r[0]) === id) { row = i + 2; data = sqJ_(r[5], null); } });
    if (row < 0 || !data) throw new Error('Không tìm thấy trong Đã xóa.');
    var ss = layFile_();
    if (data.phieu) { var shP = ss.getSheetByName('Phieu'); if (timDong_(shP, id) < 0) shP.getRange(shP.getLastRow() + 1, 1, 1, data.phieu.length).setValues([data.phieu]); }
    if (data.ct && data.ct.length) { var shC = ss.getSheetByName('ChiTiet'); shC.getRange(shC.getLastRow() + 1, 1, data.ct.length, data.ct[0].length).setValues(data.ct); }
    if (data.meta) { var shM = ss.getSheetByName(SQ_META); if (timDong_(shM, id) < 0) shM.getRange(shM.getLastRow() + 1, 1, 1, data.meta.length).setValues([data.meta]); sqMetaMoi_(); }
    sh.deleteRow(row);
    PropertiesService.getScriptProperties().setProperty('SQ_TONGHOP_CU', '1');
    return { id: id };
  });
}

/* =====================================================================
 * ẢNH HÓA ĐƠN ĐI CHỢ – app đi chợ (Index.html) gọi qua google.script.run
 * Ảnh lưu Drive, ghi vào dòng meta (SQ_Phieu) của phiếu: anh[] + them.thumbs[]. Điều hành tự hiện ảnh.
 * ===================================================================== */
function apiAnhChoDs(id) {
  id = sqKiemId_(id);
  var m = sqTimMeta_(id), t = (m && m.them) || {};
  return { anh: m ? (m.anh || []) : [], thumbs: t.thumbs || [] };
}
function apiAnhChoLuu(p) {
  p = p || {};
  var id = sqKiemId_(p.id), ngay = /^\d{4}-\d{2}-\d{2}$/.test(String(p.ngay || '')) ? p.ngay : homNay_();
  var them = (Array.isArray(p.them) ? p.them : []).slice(0, 8), bo = Array.isArray(p.bo) ? p.bo.map(String) : [];
  return khoa_(function () {
    var m = sqTimMeta_(id) || { id: id, loai: LOAI.DC, ngay: ngay, nguoi: chu_(p.nguoi, 40), moi: '', anh: [], dong: [], cong: 0, tt: TT.CHO, taoLuc: bayGio_(), them: {}, tong: 0 };
    var t = m.them = m.them || {}, anh = m.anh || [], thumbs = t.thumbs || [];
    while (thumbs.length < anh.length) thumbs.push('');
    var giu = []; anh.forEach(function (f, i) { if (bo.indexOf(f) < 0) giu.push([f, thumbs[i]]); });
    var moiFid = plAnhLuu_(them, id);
    moiFid.forEach(function (f, i) { var th = String(them[i] && them[i].thumb || ''); giu.push([f, /^[A-Za-z0-9+/=]+$/.test(th) && th.length < 12000 ? th : '']); });
    giu = giu.slice(0, 12);
    var quy = 24000; giu.forEach(function (x) { if (x[1] && x[1].length <= quy) quy -= x[1].length; else x[1] = ''; }); // ô Sheets tối đa 50k ký tự
    m.anh = giu.map(function (x) { return x[0]; }); t.thumbs = giu.map(function (x) { return x[1]; });
    t.thumb = t.thumbs[0] || ''; t.anhLuc = bayGio_();
    if (!m.ngay) m.ngay = ngay;
    sqGhiMeta_(m);
    return { anh: m.anh, thumbs: t.thumbs };
  });
}

/* =====================================================================
 * ĐẶT LẠI TÊN: mối / nhà cung cấp, nhóm đi chợ, loại chi P&L (hóa đơn quen dùng sqQuenLuu_/sqQuenXoa_)
 * ===================================================================== */
function plMoiDs_() {
  var m = {};
  sqMetaRows_().forEach(function (x) {
    if (!x.moi || x.loai === LOAI.BILL || x.loai === LOAI.DC) return;
    var o = m[x.moi] || (m[x.moi] = { ten: x.moi, so: 0, gan: '' }); o.so++; if (x.ngay > o.gan) o.gan = x.ngay;
  });
  sqQuenRows_().forEach(function (q) { if (q.moi && !m[q.moi]) m[q.moi] = { ten: q.moi, so: 0, gan: '' }; });
  return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return b.so - a.so; });
}
function plTenDs_() {
  var cfg = docCaiDat_(), lt = plTenLoai_();
  var nhomSo = {}; docPhieu_(function (x) { return cfg.nhom.indexOf(x.nhom) >= 0; }).forEach(function (x) { nhomSo[x.nhom] = (nhomSo[x.nhom] || 0) + 1; });
  return { moi: plMoiDs_(), quen: sqQuenRows_().map(function (q) { delete q._row; return q; }), nhom: cfg.nhom.map(function (n) { return { ten: n, so: nhomSo[n] || 0, giaoCa: cfg.giaoCa.indexOf(n) >= 0 }; }),
    loai: Object.keys(PL_LOAI).map(function (k) { return { k: k, ten: lt[k] ? lt[k].ten : PL_LOAI[k].ten, icon: lt[k] ? lt[k].icon : PL_LOAI[k].icon, goc: PL_LOAI[k].ten }; }) };
}
// đổi cột của 1 tab theo bảng {cũ → mới}; cot = số cột (1 = A); tien = hàm biến giá trị cũ thành mới hoặc null
function plDoiCot_(ten, cot, doi) {
  var sh = layFile_().getSheetByName(ten); if (!sh) return 0;
  var n = sh.getLastRow(); if (n < 2) return 0;
  var r = sh.getRange(2, cot, n - 1, 1), v = r.getValues(), dem = 0;
  v.forEach(function (x) { var moi = doi(String(x[0])); if (moi !== null && moi !== String(x[0])) { x[0] = moi; dem++; } });
  if (dem) r.setValues(v);
  return dem;
}
// p.ds = [{cu, moi}]
function plDoiTenMoi_(p) {
  var bang = {};
  (Array.isArray(p.ds) ? p.ds : []).slice(0, 60).forEach(function (x) { var cu = String(x.cu || '').trim(), moi = chu_(x.moi, 60); if (cu && moi && cu !== moi) bang[cu] = moi; });
  if (!Object.keys(bang).length) throw new Error('Chưa có tên nào đổi.');
  return khoa_(function () {
    var dem = 0;
    dem += plDoiCot_(SQ_META, M.MOI + 1, function (s) { return bang[s] || null; });
    dem += plDoiCot_('Phieu', 3, function (s) { var m = /^HĐ (.*)$/.exec(s); return m && bang[m[1]] ? 'HĐ ' + bang[m[1]] : null; });
    plDoiCot_('ChiTiet', 3, function (s) { var m = /^HĐ (.*)$/.exec(s); return m && bang[m[1]] ? 'HĐ ' + bang[m[1]] : null; });
    plDoiCot_('DanhMuc', 1, function (s) { var m = /^HĐ (.*)$/.exec(s); return m && bang[m[1]] ? 'HĐ ' + bang[m[1]] : null; });
    plDoiCot_(SQ_ALIAS, 1, function (s) { return bang[s] || null; });
    plDoiCot_(SQ_HOC, 1, function (s) { return bang[s] || null; });
    plDoiCot_(SQ_QUEN, 4, function (s) { return bang[s] || null; });
    sqMetaMoi_(); sqCacheXoa_('sq_dm');
    PropertiesService.getScriptProperties().setProperty('SQ_TONGHOP_CU', '1');
    return { doi: dem, ten: plTenDs_() };
  });
}
// đổi tên 1 nhóm đi chợ: CaiDat (Nhóm, Nhóm có tiền giao ca), Phieu, ChiTiet, DanhMuc, HayMua
function plDoiTenNhom_(p) {
  var cu = String(p.cu || '').trim(), moi = chu_(p.moi, 40).replace(/[,|<>"]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!cu || !moi || cu === moi) throw new Error('Chưa nhập tên mới.');
  if (/^HĐ /.test(moi) || moi === NHOM_TU_DAT || PL_NHOM_KHONG_HANG.indexOf(moi) >= 0) throw new Error('Tên này app đang dùng riêng, chọn tên khác.');
  return khoa_(function () {
    var sh = layFile_().getSheetByName('CaiDat'), rows = sh.getDataRange().getValues(), cfg = docCaiDat_();
    if (cfg.nhom.indexOf(cu) < 0) throw new Error('Không thấy nhóm ' + cu + '.');
    if (cfg.nhom.some(function (n) { return n !== cu && khongDau_(n) === khongDau_(moi); })) throw new Error('Đã có nhóm tên ' + moi + '.');
    for (var i = 1; i < rows.length; i++) {
      var k = String(rows[i][0]).trim();
      if (k === 'Nhóm' || k === 'Nhóm có tiền giao ca') sh.getRange(i + 1, 2).setValue(tachDs_(rows[i][1]).map(function (n) { return n === cu ? moi : n; }).join(', '));
    }
    _CFG = null;
    var d = function (s) { return s === cu ? moi : null; };
    var dem = plDoiCot_('Phieu', 3, d); plDoiCot_('ChiTiet', 3, d); plDoiCot_('DanhMuc', 1, d); plDoiCot_('HayMua', 1, d);
    sqCacheXoa_('sq_dm');
    PropertiesService.getScriptProperties().setProperty('SQ_TONGHOP_CU', '1');
    return { doi: dem, ten: plTenDs_() };
  });
}
// tên hiển thị của loại chi trong P&L (lưu Script Property PL_TEN_LOAI)
function plTenLoai_() { return sqJ_(PropertiesService.getScriptProperties().getProperty('PL_TEN_LOAI'), {}) || {}; }
function plDoiTenLoai_(p) {
  var o = {};
  (Array.isArray(p.ds) ? p.ds : []).forEach(function (x) { if (!PL_LOAI[x.k]) return; var ten = chu_(x.ten, 40), icon = chu_(x.icon, 4);
    if ((ten && ten !== PL_LOAI[x.k].ten) || (icon && icon !== PL_LOAI[x.k].icon)) o[x.k] = { ten: ten || PL_LOAI[x.k].ten, icon: icon || PL_LOAI[x.k].icon }; });
  PropertiesService.getScriptProperties().setProperty('PL_TEN_LOAI', JSON.stringify(o));
  return { ten: plTenDs_() };
}
function plApTenLoai_() { var lt = plTenLoai_(); Object.keys(lt).forEach(function (k) { if (PL_LOAI[k]) { PL_LOAI[k].ten = lt[k].ten; PL_LOAI[k].icon = lt[k].icon; } }); return lt; }

/* =====================================================================
 * KIỂM CHI PHÍ NGÀY – kế toán dò từng khoản chi trong ngày (Sổ Quán), tick ✓ đúng / ! sai + ghi chú
 * ===================================================================== */
var PL_KC = 'SQ_KiemChi', H_PL_KC = ['Ngày', 'Khoản', 'Trạng thái', 'Loại sai', 'Ghi chú', 'Người kiểm', 'Lúc', 'Tên khoản'];
var KC_LOAI_SAI = ['Sai số tiền', 'Sai số lượng', 'Thiếu hóa đơn / ảnh', 'Trùng', 'Khác'];
function plTabKC_() { return sqTab_(PL_KC, H_PL_KC, function (sh) { sh.getRange('A:B').setNumberFormat('@'); }); }
function plKCRows_() {
  var sh = plTabKC_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_PL_KC.length).getValues().map(function (r, i) {
    return { ngay: chuanNgay_(r[0]), khoa: String(r[1]), tt: String(r[2]), loaiSai: String(r[3]), ghi: String(r[4]), nguoi: String(r[5]), luc: String(r[6]), ten: String(r[7] || ''), row: i + 2 };
  }).filter(function (x) { return x.khoa; });
}
function plKCQuyen_(nv) { if ((nv.chucNang || []).indexOf('kiemchi') < 0) throw new Error('Em chưa được mở chức năng Kiểm chi phí ngày. Nhờ anh Đạt bật trong Điều hành → Nhân viên.'); }
var KC_DUA_CHO = /(^| )(tien cho|di cho|dua cho|cho doan|doan)( |$)/, KC_NHAN_CHO = /(du cho|tien cho|cho tra|tra lai|thoi lai|doan tra|du di cho)/;
function plKCNgay_(nv, p) {
  plKCQuyen_(nv);
  var hn = homNay_(), ngay = /^\d{4}-\d{2}-\d{2}$/.test(String(p.ngay || '')) ? p.ngay : hn;
  if (ngay > hn) ngay = hn;
  if (ngay < congNgay_(hn, -7)) throw new Error('Chỉ xem được 7 ngày gần nhất.');
  var danh = {}; plKCRows_().forEach(function (x) { if (x.ngay === ngay) danh[x.khoa] = { tt: x.tt, loaiSai: x.loaiSai, ghi: x.ghi, nguoi: x.nguoi, luc: x.luc }; });
  var metas = {}; sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  var hd = sqPhieuTrong_(ngay, ngay).filter(function (x) { return x.loai !== LOAI.BILL; }).map(function (x) {
    var m = metas[x.id] || {}, t = m.them || {}, lc = null; try { lc = plLoaiChinh_(Object.assign({}, x, { meta: metas[x.id] || null })); } catch (e) { }
    var ten = (x.loai === LOAI.HD || x.loai === LOAI.TD) ? (x.moi || x.nhom.replace(/^HĐ /, '')) : x.loai === LOAI_CHU ? (x.moi || x.nhom) : ((x.nguoi ? x.nguoi + ' – ' : '') + x.nhom);
    return { khoa: x.id, id: x.id, ten: ten, loai: x.loai, nguoi: x.nguoi, tong: x.tong, soDong: x.mon.length, mon: x.mon.slice(0, 40), anh: x.anh || [], thumb: t.thumb || '',
      tienNhan: x.tienNhan, nhom: x.nhom, lc: lc ? lc.k : '', kc: danh[x.id] || null };
  }).sort(function (a, b) { return b.tong - a.tong; });
  var doi = sqQuayDoi_(), quay = [], duaCho = 0, nhanLai = 0, coTN = false;
  sqTnRows_().filter(function (x) { return x.ngay === ngay; }).sort(function (a, b) { return a.ca - b.ca; }).forEach(function (x) {
    coTN = true;
    var dong = x.chi.filter(function (r) { return r.amount; }).map(function (r) {
      var auto = !CP_LUAN_CHUYEN.test(khongDau_(r.note)), k = x.ngay + '|' + x.ca + '|' + khongDau_(r.note), tinh = doi[k] === undefined ? auto : doi[k];
      if (!tinh && KC_DUA_CHO.test(khongDau_(r.note))) duaCho += r.amount;
      return { note: r.note, tien: r.amount, tinh: tinh };
    });
    (x.nhan || []).forEach(function (r) { if (r.amount && KC_NHAN_CHO.test(khongDau_(r.note))) nhanLai += r.amount; });
    if (!dong.length) return;
    var khoa = 'quay|' + x.ca;
    quay.push({ khoa: khoa, ca: x.ca, nguoi: x.nguoi, dong: dong, tong: dong.reduce(function (s, d) { return s + (d.tinh ? d.tien : 0); }, 0), kc: danh[khoa] || null });
  });
  // đối chiếu tiền đi chợ: thu ngân đưa ↔ phiếu đi chợ (có tiền giao ca) ↔ dư trả lại
  var gc = docCaiDat_().giaoCa, pc = hd.filter(function (x) { return x.loai === LOAI.DC && gc.indexOf(x.nhom) >= 0; });
  var phNhan = pc.reduce(function (s, x) { return s + (x.tienNhan === '' || x.tienNhan == null ? 0 : so_(x.tienNhan)); }, 0), phChi = pc.reduce(function (s, x) { return s + x.tong; }, 0);
  var dc = null;
  if (pc.length || duaCho) {
    var duPhai = phNhan - phChi;
    dc = { duaCho: duaCho, phNhan: phNhan, phChi: phChi, duPhai: duPhai, nhanLai: nhanLai, coTN: coTN, nguoi: pc.map(function (x) { return x.nguoi; }).filter(String).join(', '),
      lechNhan: coTN && pc.length ? phNhan - duaCho : null, lechDu: coTN && pc.length ? nhanLai - duPhai : null };
  }
  var tongChi = hd.reduce(function (s, x) { return s + x.tong; }, 0) + quay.reduce(function (s, q) { return s + q.tong; }, 0);
  var all = hd.length + quay.length, ok = 0, sai = 0;
  hd.concat(quay).forEach(function (x) { if (x.kc && x.kc.tt === 'ok') ok++; else if (x.kc && x.kc.tt === 'sai') sai++; });
  return { ngay: ngay, homNay: hn, hd: hd, quay: quay, dc: dc, tongChi: tongChi, soKhoan: all, daKiem: ok + sai, sai: sai, loaiSai: KC_LOAI_SAI };
}
// p: {ngay, khoa, ten, tt:'ok'|'sai'|'', loaiSai, ghi}
function plKCDanh_(nv, p) {
  plKCQuyen_(nv);
  var ngay = sqKiemNgay_(p.ngay), khoa = String(p.khoa || '').slice(0, 60), tt = ['ok', 'sai', ''].indexOf(p.tt) >= 0 ? p.tt : '';
  if (!khoa) throw new Error('Thiếu khoản.');
  if (tt === 'sai' && !chu_(p.ghi, 300) && !p.loaiSai) throw new Error('Ghi rõ sai ở đâu nhé.');
  return khoa_(function () {
    var sh = plTabKC_(), cu = plKCRows_().filter(function (x) { return x.ngay === ngay && x.khoa === khoa; })[0];
    var row = [ngay, khoa, tt, tt === 'sai' ? chu_(p.loaiSai, 40) : '', tt === 'sai' ? chu_(p.ghi, 300) : '', nv.ten, bayGio_(), chu_(p.ten, 80)];
    if (cu) sh.getRange(cu.row, 1, 1, H_PL_KC.length).setValues([row]); else sh.getRange(sh.getLastRow() + 1, 1, 1, H_PL_KC.length).setValues([row]);
    // sai trên hóa đơn ảnh / hóa đơn quen → gửi câu hỏi cho người nhập (hiện trên app của họ)
    if (tt === 'sai' && !/^quay\|/.test(khoa)) {
      var m = sqTimMeta_(khoa);
      if (m && (m.loai === LOAI.HD || m.loai === LOAI.TD) && m.nguoi && m.nguoi !== nv.ten && m.nguoi !== 'anh Đạt' && TT_CHUA_XAC_NHAN.indexOf(m.tt) < 0) {
        m.hoi = 'Kế toán ' + nv.ten + ': ' + [chu_(p.loaiSai, 40), chu_(p.ghi, 250)].filter(String).join(' – '); m.traLoi = ''; m.tt = TT.HOI; sqGhiMeta_(m);
      }
    }
    return { khoa: khoa, tt: tt };
  });
}
// cho Điều hành: dấu kế toán theo khoản trong khoảng ngày
function plKCMap_(tu, den) {
  var m = {}; plKCRows_().forEach(function (x) { if (x.ngay >= tu && x.ngay <= den && x.tt) m[x.ngay + '|' + x.khoa] = { tt: x.tt, loaiSai: x.loaiSai, ghi: x.ghi, nguoi: x.nguoi, luc: x.luc }; });
  return m;
}
