/**
 * CHI PHÍ & HÓA ĐƠN – Điều hành (anh Đạt). Thay tab Duyệt phiếu + Giá nhập:
 *  - hóa đơn / phiếu đi chợ / hóa đơn quen theo ngày (hình nhỏ, ảnh thật, các dòng, giá so với lần trước)
 *  - chi tại quầy trong sổ thu ngân (đá, bánh mì…) – tiền đưa đi chợ / giao ca tự để riêng, anh bấm đổi được
 *  - cảnh báo: giá nhập (vàng <10%, đỏ 10–20%, đỏ đậm >20%, xanh giảm ≥5%), hóa đơn cần xem, chi tại quầy tăng bất thường
 *  - anh sửa hóa đơn (ghi lại Phieu/ChiTiet → P&L), hỏi nhân viên (chu_duyet tt 'Hỏi lại'), xem các lần mua của 1 mối
 */
var DH_QUAY = 'DH_ChiQuay';
var H_DH_QUAY = ['Ngày', 'Ca', 'Ghi chú', 'Tính chi phí (có/không)', 'Lúc'];
var CP_LUAN_CHUYEN = /(^| )(tien cho|di cho|dua cho|giao ca|doan|cho doan|ung)( |$)|^cho( |$)/;

function sqCpMuc_(pct) { return pct < 0 ? (pct <= -0.05 ? 'g' : '') : pct > 0.2 ? 'rr' : pct >= 0.1 ? 'r' : pct >= 0.005 ? 'y' : ''; }
function sqCpKhoaMoi_(x) {
  if (x.loai === LOAI.HD || x.loai === LOAI.TD) { var m = x.moi || x.nhom.replace(/^HĐ /, ''); return { khoa: 'm:' + khongDau_(m), ten: m }; }
  if (x.loai === LOAI_CHU) return { khoa: x.moi ? 'm:' + khongDau_(x.moi) : 'c:' + khongDau_(x.nhom), ten: (x.moi || x.nhom) + (x.mon && x.mon[0] ? ' · ' + x.mon[0].ten : '') };
  var t = (x.nguoi ? x.nguoi + ' – ' : '') + x.nhom;
  return { khoa: 'n:' + khongDau_(x.nhom) + '|' + khongDau_(x.nguoi), ten: t };
}
function sqTabQuay_() { return sqTab_(DH_QUAY, H_DH_QUAY, function (sh) { sh.getRange('A:A').setNumberFormat('@'); }); }
function sqQuayDoi_() {
  var sh = sqTabQuay_(), n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (r) { m[chuanNgay_(r[0]) + '|' + so_(r[1]) + '|' + khongDau_(r[2])] = String(r[3]) === 'có'; });
  return m;
}
function sqQuayKhoa_(note) { return khongDau_(note).replace(/[0-9.,]+/g, ' ').replace(/\s+/g, ' ').trim().split(' ').slice(0, 2).join(' '); }

function sqCpNgay_(p) {
  var hn = homNay_();
  var den = /^\d{4}-\d{2}-\d{2}$/.test(String(p.den || '')) ? p.den : hn;
  var tu = /^\d{4}-\d{2}-\d{2}$/.test(String(p.tu || '')) ? p.tu : den;
  if (tu > den) { var t0 = tu; tu = den; den = t0; }
  if (tu < congNgay_(den, -31)) tu = congNgay_(den, -31);
  var coThumb = tu >= congNgay_(den, -8);
  var metas = {};
  sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  var gia = sqTinhGia_(sqDongMua_());
  var kcMap = {}; try { kcMap = plKCMap_(tu, den); } catch (e) { }
  var ds = sqPhieuTrong_(tu, den).filter(function (x) { return x.loai !== LOAI.BILL; }).map(function (x) {
    var m = metas[x.id] || {}, k = sqCpKhoaMoi_(x), cb = [];
    var gp = (gia.theoPhieu[x.id] || []).map(function (g) { return { stt: g.stt, ten: g.ten, khoa: g.khoa, dv: g.dv, gia: g.gia, sl: g.sl, pct: g.pct, muc: sqCpMuc_(g.pct), chenh: g.chenh, traThem: g.traThem, truoc: g.truoc }; });
    if (x.loai === LOAI.HD && x.tongAnh !== '' && x.tongAnh > 0 && Math.abs(x.tongAnh - x.tong) > 1000)
      cb.push({ loai: 'tổng lệch', text: 'Cộng các dòng ' + sqTien_(x.tong) + ' ≠ tổng trên ảnh ' + sqTien_(x.tongAnh) });
    if (x.soSua) cb.push({ loai: 'sửa tay ' + x.soSua + ' dòng', text: 'Nhân viên sửa khác chỗ app đọc' });
    if (x.loai === LOAI.HD && (!x.anh || !x.anh.length) && x.nguoi !== 'anh Đạt') cb.push({ loai: 'không ảnh', text: 'Hóa đơn không có ảnh' });
    if (x.tong && x.tong % 100000 === 0 && x.mon.length > 1) cb.push({ loai: 'tổng tròn', text: 'Tổng phiếu tròn ' + sqTien_(x.tong) });
    var dongAi = (m.dong || []).map(function (d) { return { ten: d.ten, chu: d.chu, anh: d.anh, vung: d.vung, sua: d.sua, aiTen: d.aiTen, aiTien: d.aiTien }; });
    return { id: x.id, ngay: x.ngay, loai: x.loai, nhom: x.nhom, moi: x.moi, nguoi: m.nguoi || x.nguoi, tong: x.tong, tongAnh: x.tongAnh, mon: x.mon, ghiChu: x.ghiChu,
      tt: x.tt, hoi: x.hoi, traLoi: x.traLoi, anh: x.anh || [], khoa: k.khoa, tenMoi: k.ten, gia: gp, cb: cb, dongAi: dongAi,
      kc: kcMap[x.ngay + '|' + x.id] || null,
      lc: (function () { try { return plLoaiChinh_(Object.assign({}, x, { meta: metas[x.id] || null })); } catch (e) { return null; } })(), lcTay: !!(m.them && m.them.loaiChi),
      thumb: coThumb && m.them ? (m.them.thumb || '') : '', kiemBoi: m.them ? (m.them.kiemBoi || '') : '', chuSua: m.them ? (m.them.chuSua || '') : '', suaLuc: x.suaLuc };
  }).sort(function (a, b) { return a.ngay < b.ngay ? 1 : a.ngay > b.ngay ? -1 : b.tong - a.tong; });

  // chi tại quầy (sổ thu ngân)
  var doi = sqQuayDoi_(), tn = sqTnRows_(), quay = [];
  var tinhDong = function (x, r) {
    var auto = !CP_LUAN_CHUYEN.test(khongDau_(r.note)), k = x.ngay + '|' + x.ca + '|' + khongDau_(r.note);
    return { note: r.note, amount: r.amount, auto: auto, tinh: doi[k] === undefined ? auto : doi[k] };
  };
  tn.filter(function (x) { return x.ngay >= tu && x.ngay <= den; }).forEach(function (x) {
    var dong = x.chi.filter(function (r) { return r.amount; }).map(function (r) { return tinhDong(x, r); });
    if (!dong.length) return;
    quay.push({ ngay: x.ngay, ca: x.ca, nguoi: x.nguoi, luc: x.luc, dong: dong, kc: kcMap[x.ngay + '|quay|' + x.ca] || null, tong: dong.reduce(function (s, d) { return s + (d.tinh ? d.amount : 0); }, 0) });
  });
  quay.sort(function (a, b) { return a.ngay < b.ngay ? 1 : a.ngay > b.ngay ? -1 : a.ca - b.ca; });

  // cảnh báo giá nhập trong kỳ
  var cbGia = [];
  ds.forEach(function (x) { x.gia.forEach(function (g) { if (g.muc) cbGia.push({ id: x.id, ngay: x.ngay, tenMoi: x.tenMoi, nguoi: x.nguoi, stt: g.stt, ten: g.ten, khoa: g.khoa, dv: g.dv, gia: g.gia, sl: g.sl, pct: g.pct, muc: g.muc, chenh: g.chenh, traThem: g.traThem, truoc: g.truoc }); }); });
  var ten = {}; docPhieu_(function (x) { return x.ngay >= congNgay_(tu, -120) && x.ngay <= den; }).forEach(function (x) { ten[x.id] = sqCpKhoaMoi_({ loai: (metas[x.id] || {}).loai || LOAI.DC, moi: (metas[x.id] || {}).moi || '', nhom: x.nhom, nguoi: x.nguoiMua }).ten; });
  cbGia.forEach(function (g) { g.truoc.tenMoi = ten[g.truoc.id] || ''; });
  var hang = { rr: 0, r: 1, y: 2, g: 3 };
  cbGia.sort(function (a, b) { return hang[a.muc] - hang[b.muc] || Math.abs(b.pct) - Math.abs(a.pct); });

  // chi tại quầy tăng bất thường: so tổng theo khoản trong ngày với trung bình 7 ngày trước
  var theoNgay = {};
  tn.filter(function (x) { return x.ngay >= congNgay_(tu, -7) && x.ngay <= den; }).forEach(function (x) {
    x.chi.forEach(function (r) {
      if (!r.amount) return; var d = tinhDong(x, r); if (!d.tinh) return;
      var k = sqQuayKhoa_(r.note); if (!k) return;
      var o = (theoNgay[x.ngay] = theoNgay[x.ngay] || {}); o[k] = o[k] || { ten: r.note, tien: 0 }; o[k].tien += r.amount;
    });
  });
  var cbQuay = [];
  Object.keys(theoNgay).filter(function (d) { return d >= tu && d <= den; }).forEach(function (d) {
    Object.keys(theoNgay[d]).forEach(function (k) {
      var cur = theoNgay[d][k], tr = [];
      for (var i = 1; i <= 7; i++) { var o = theoNgay[congNgay_(d, -i)]; if (o && o[k]) tr.push(o[k].tien); }
      if (tr.length < 2) return;
      var tb = tr.reduce(function (s, v) { return s + v; }, 0) / tr.length, pct = (cur.tien - tb) / tb;
      if (cur.tien < 30000 || pct < 0.1) return;
      cbQuay.push({ ngay: d, ten: cur.ten, tien: cur.tien, tb: Math.round(tb), soNgay: tr.length, pct: pct, muc: sqCpMuc_(pct) });
    });
  });
  cbQuay.sort(function (a, b) { return b.pct - a.pct; });

  return { tu: tu, den: den, homNay: hn, ds: ds, quay: quay, loaiTen: (function () { try { return plTenLoai_(); } catch (e) { return {}; } })(),
    canhBao: { gia: cbGia, hd: ds.filter(function (x) { return x.cb.length; }).map(function (x) { return { id: x.id, ngay: x.ngay, tenMoi: x.tenMoi, tong: x.tong, cb: x.cb }; }), quay: cbQuay } };
}

// Các lần mua của 1 mối / người mua (120 ngày)
function sqCpMoi_(p) {
  var khoa = String(p.khoa || ''), hn = homNay_(), metas = {};
  sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  var ds = sqPhieuTrong_(congNgay_(hn, -120), hn).filter(function (x) { return x.loai !== LOAI.BILL && sqCpKhoaMoi_(x).khoa === khoa; })
    .map(function (x) { return { id: x.id, ngay: x.ngay, tong: x.tong, nguoi: (metas[x.id] || {}).nguoi || x.nguoi, soDong: x.mon.length, ten: sqCpKhoaMoi_(x).ten, mon: x.mon.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) }; })
    .sort(function (a, b) { return a.ngay < b.ngay ? 1 : -1; }).slice(0, 30);
  return { khoa: khoa, ten: ds.length ? ds[0].ten : '', ds: ds };
}

// Anh Đạt sửa hóa đơn: ghi lại Phieu/ChiTiet (P&L đọc) + meta
function sqCpSua_(p) {
  var id = sqKiemId_(p.id);
  var mon = sqMon_(p.mon);
  if (!mon.length) throw new Error('Hóa đơn phải còn ít nhất 1 dòng.');
  return khoa_(function () {
    var ph = docPhieu_(function (x) { return x.id === id; })[0];
    if (!ph) throw new Error('Không tìm thấy hóa đơn.');
    sqApiLuu_({ id: id, ngay: ph.ngay, nhom: ph.nhom, nguoiMua: ph.nguoiMua, tienNhan: ph.tienNhan, ghiChu: ph.ghiChu,
      mon: mon.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) });
    var m = sqTimMeta_(id), tong = mon.reduce(function (s, d) { return s + d.tien; }, 0);
    if (m) {
      var cu = m.dong || [];
      m.dong = mon.map(function (d, i) { var c = cu[i] || {}; return Object.assign({}, c, { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }); });
      m.cong = tong; m.tong = tong;
      m.them = Object.assign({}, m.them || {}, { chuSua: bayGio_() });
      sqGhiMeta_(m);
    }
    return { id: id, tong: tong };
  });
}

// Đổi 1 khoản chi tại quầy: tính / không tính chi phí
function sqCpQuay_(p) {
  var ngay = sqKiemNgay_(p.ngay), ca = so_(p.ca), note = chu_(p.note, 80), tinh = p.tinh === true;
  return khoa_(function () {
    var sh = sqTabQuay_(), n = sh.getLastRow(), row = -1;
    if (n >= 2) sh.getRange(2, 1, n - 1, 3).getValues().forEach(function (r, i) { if (chuanNgay_(r[0]) === ngay && so_(r[1]) === ca && khongDau_(r[2]) === khongDau_(note)) row = i + 2; });
    if (row < 0) row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 5).setValues([[ngay, ca, note, tinh ? 'có' : 'không', bayGio_()]]);
    return { tinh: tinh };
  });
}

// Tổng theo mối: mọi hóa đơn (ảnh, quen, anh nhập) + phiếu đi chợ trong khoảng ngày, kèm cùng kỳ tháng trước.
// Trả dòng thô để Điều hành tự cộng theo mối / theo món / tìm món.
function sqThangLui_(d) { var p = d.split('-'), y = +p[0], m = +p[1] - 1; if (!m) { m = 12; y--; } var cuoi = new Date(Date.UTC(y, m, 0)).getUTCDate(); return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + Math.min(+p[2], cuoi)).slice(-2); }
function sqCpTong_(p) {
  var hn = homNay_(), N = /^\d{4}-\d{2}-\d{2}$/;
  var den = N.test(String(p.den || '')) ? p.den : hn, tu = N.test(String(p.tu || '')) ? p.tu : den.slice(0, 8) + '01';
  if (tu > den) { var t0 = tu; tu = den; den = t0; }
  if (tu < congNgay_(den, -95)) tu = congNgay_(den, -95);
  var tuT = sqThangLui_(tu), denT = sqThangLui_(den);
  var moiDs = [], moiIdx = {}, ph = [], dong = [], truoc = { moi: {}, mon: {} };
  var khoaMoi = function (x) {
    if (x.loai === LOAI.DC) return { k: 'd:' + khongDau_(x.nhom), ten: x.nhom, dc: 1 };
    if (x.loai === LOAI.HD || x.loai === LOAI.TD) { var m = x.moi || x.nhom.replace(/^HĐ /, ''); return { k: 'm:' + khongDau_(m), ten: m }; }
    return x.moi ? { k: 'm:' + khongDau_(x.moi), ten: x.moi } : { k: 'c:' + khongDau_(x.nhom), ten: x.nhom };
  };
  var metas = {}; sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  sqPhieuTrong_(tuT < tu ? tuT : tu, den).forEach(function (x) {
    if (x.loai === LOAI.BILL || PL_NHOM_KHONG_HANG.indexOf(x.nhom) >= 0 || !x.tong) return;
    // chỉ hàng mua (nguyên liệu, ga, bánh mì) – bỏ cố định (điện nước), đầu tư, sửa chữa, lương, chi khác
    try { var lc = plLoaiChinh_(Object.assign({}, x, { meta: metas[x.id] || null })); if (['nvl', 'ga', 'banhmi'].indexOf(lc.k) < 0) return; } catch (e) { }
    var km = khoaMoi(x);
    if (x.ngay >= tuT && x.ngay <= denT && x.ngay < tu) { // cùng kỳ tháng trước
      truoc.moi[km.k] = (truoc.moi[km.k] || 0) + x.tong;
      x.mon.forEach(function (d) { var k = km.k + '|' + khongDau_(d.ten); truoc.mon[k] = (truoc.mon[k] || 0) + so_(d.tien); });
      return;
    }
    if (x.ngay < tu) return;
    var mi = moiIdx[km.k];
    if (mi === undefined) { mi = moiIdx[km.k] = moiDs.length; moiDs.push({ k: km.k, ten: km.ten, dc: km.dc || 0, nguoi: [] }); }
    if (x.nguoi && moiDs[mi].nguoi.indexOf(x.nguoi) < 0 && moiDs[mi].nguoi.length < 6) moiDs[mi].nguoi.push(x.nguoi);
    var pi = ph.length; ph.push([x.id, x.ngay, mi, x.tong]);
    x.mon.forEach(function (d) { if (d.ten) dong.push([pi, d.ten, d.sl === '' ? '' : so_(d.sl), d.dv || '', d.gia === '' ? '' : so_(d.gia), so_(d.tien)]); });
  });
  return { tu: tu, den: den, tuT: tuT, denT: denT, homNay: hn, moi: moiDs, p: ph, d: dong, truoc: truoc };
}
