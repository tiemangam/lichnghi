/**
 * CHI PHÍ & HÓA ĐƠN – Điều hành (anh Đạt). Thay tab Duyệt phiếu + Giá nhập:
 *  - hóa đơn / phiếu đi chợ / hóa đơn quen theo ngày (hình nhỏ, ảnh thật, các dòng, giá so với lần trước)
 *  - chi tại quầy trong sổ thu ngân (đá, bánh mì…) – tiền đưa đi chợ / giao ca tự để riêng, anh bấm đổi được
 *  - cảnh báo: giá nhập (vàng <10%, đỏ 10–20%, đỏ đậm >20%, xanh giảm ≥5%), hóa đơn cần xem, chi tại quầy tăng bất thường
 *  - anh sửa hóa đơn (ghi lại Phieu/ChiTiet → P&L), hỏi nhân viên (chu_duyet tt 'Hỏi lại'), xem các lần mua của 1 mối
 */
var DH_QUAY = 'DH_ChiQuay';
var H_DH_QUAY = ['Ngày', 'Ca', 'Ghi chú', 'Tính chi phí (có/không)', 'Lúc'];
var CP_LUAN_CHUYEN = /(^| )(tien cho|di cho|dua cho|giao ca|doan|cho doan|ung|chi ngoc|c ngoc)( |$)|^cho( |$)/;

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

// Sổ thu ngân trong khoảng ngày, để tìm theo từ khóa (chi, nhận thêm như bán ve chai / dầu cặn, chuyển khoản, lý do két thiếu)
// [ngày, ca, loại (chi|nhan|ck|lydo), nội dung, tiền, người, tính chi phí (chỉ với chi: 1 có, 0 không)]
function sqCpQuayDs_(tu, den) {
  var doi = {}; try { doi = sqQuayDoi_(); } catch (e) { }
  var out = [];
  sqTnRows_().forEach(function (x) {
    if (x.ngay < tu || x.ngay > den) return;
    (x.chi || []).forEach(function (r) {
      if (!r.note && !r.amount) return;
      var auto = !CP_LUAN_CHUYEN.test(khongDau_(r.note)), k = x.ngay + '|' + x.ca + '|' + khongDau_(r.note);
      out.push([x.ngay, x.ca, 'chi', r.note, r.amount, x.nguoi, (doi[k] === undefined ? auto : doi[k]) ? 1 : 0]);
    });
    (x.nhan || []).forEach(function (r) { if (r.note || r.amount) out.push([x.ngay, x.ca, 'nhan', r.note, r.amount, x.nguoi, 1]); });
    (x.ck || []).forEach(function (r) { if (r.note || r.amount) out.push([x.ngay, x.ca, 'ck', r.note, r.amount, x.nguoi, 1]); });
    if (x.lyDo) out.push([x.ngay, x.ca, 'lydo', x.lyDo, x.lech, x.nguoi, 1]);
  });
  out.sort(function (a, b) { return a[0] < b[0] ? 1 : a[0] > b[0] ? -1 : a[1] - b[1]; });
  return out;
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
  var nhomDC = {}; try { (docCaiDat_().nhom || []).forEach(function (n) { nhomDC[khongDau_(n)] = 1; }); } catch (e) { }
  var khoaMoi = function (x) {
    if (x.loai === LOAI.DC) return { k: 'd:' + khongDau_(x.nhom), ten: x.nhom, dc: 1 };
    var m = (x.loai === LOAI.HD || x.loai === LOAI.TD) ? (x.moi || x.nhom.replace(/^HĐ /, '')) : (x.moi || x.nhom), km = khongDau_(m);
    // mối trùng tên nhóm đi chợ (vd hóa đơn ảnh ghi mối "Đi chợ hằng ngày") hoặc không có mối → gộp vào nhóm đó
    return (nhomDC[km] || (!x.moi && x.loai === LOAI_CHU)) ? { k: 'd:' + km, ten: m } : { k: 'm:' + km, ten: m };
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
    if (km.dc) moiDs[mi].dc = 1;
    if (x.nguoi && moiDs[mi].nguoi.indexOf(x.nguoi) < 0 && moiDs[mi].nguoi.length < 6) moiDs[mi].nguoi.push(x.nguoi);
    var pi = ph.length; ph.push([x.id, x.ngay, mi, x.tong]);
    x.mon.forEach(function (d) { if (d.ten) dong.push([pi, d.ten, d.sl === '' ? '' : so_(d.sl), d.dv || '', d.gia === '' ? '' : so_(d.gia), so_(d.tien)]); });
  });
  return { tu: tu, den: den, tuT: tuT, denT: denT, homNay: hn, moi: moiDs, p: ph, d: dong, truoc: truoc, quay: sqCpQuayDs_(tu, den) };
}

/* =====================================================================
 * TỔNG QUAN CHI PHÍ THEO NHÓM (Điều hành → Chi phí & Hóa đơn → 📊 Tổng quan)
 *  - mọi khoản chi trong khoảng ngày (hóa đơn ảnh, phiếu đi chợ, hóa đơn quen, chi tại quầy, cúng kiến, sửa quán, ship, lương, điện nước…)
 *    gom theo NHÓM → MÓN → NGÀY MUA; kèm % doanh thu (Sapo) và % tổng chi, so cùng kỳ tháng trước
 *  - nhóm của từng món: anh đã xếp (nhớ) → theo loại chi anh chọn → app đoán theo tên → "Chưa xếp nhóm"
 *  - đổi tên / gộp tên món chỉ đổi trong màn tổng quan, KHÔNG sửa hóa đơn gốc (P&L không đổi)
 * ===================================================================== */
var CPN_NHOM = 'DH_NhomChi', H_CPN_NHOM = ['Mã', 'Tên', 'Icon', 'Thứ tự', 'Gộp vào', 'Sửa lúc'];
var CPN_MON = 'DH_NhomMon', H_CPN_MON = ['Kiểu', 'Khóa', 'Tên chuẩn', 'Nhóm', 'Sửa lúc'];   // Kiểu: mon (đúng tên món) | chua (tên món có chứa chữ này)
var CPN_MAC_DINH = [
  ['changa', 'Chân gà', '🍗'], ['bo', 'Bò', '🥩'], ['heo', 'Heo & thịt khác', '🐷'], ['ca', 'Cá & hải sản', '🐟'],
  ['giavi', 'Gia vị · hành tỏi ớt · xay', '🌶'], ['rau', 'Rau củ · trái cây', '🥬'], ['gao', 'Gạo · bún · bánh tráng · trứng', '🍚'],
  ['nuoc', 'Nước uống · sữa chua', '🥤'], ['baobi', 'Bao bì · giấy · vật tư', '📦'], ['ship', 'Tiền ship', '🛵'],
  ['cung', 'Cúng kiến', '🙏'], ['sua', 'Sửa quán · vật dụng', '🔧'], ['ga', 'Ga · gas', '🔥'], ['luong', 'Lương · thưởng', '👥'],
  ['codinh', 'Tiền nhà · điện nước · cố định', '🏠'], ['khac', 'Chi khác', '📌']
];
// loại chi anh đã chọn ở hóa đơn → nhóm (nvl / khác để app đoán theo tên món)
var CPN_THEO_LOAI = { cd: 'codinh', luongkhac: 'luong', sua: 'sua', ga: 'ga', banhmi: 'gao' };
// app đoán theo tên món (so trên chữ có dấu, chữ thường). Thứ tự = ưu tiên.
var CPN_DOAN = [
  ['cung', /cúng|nhang|hương|vàng mã|tiền vàng|(^| )nến( |$)|(^| )hoa( |$)/],
  ['ship', /ship|vận chuyển|phí giao|giao hàng|grab|ahamove|(^| )bế( |$)|xe ôm/],
  ['ga', /(^| )(ga|gas)( |$)|bình ga|bình gas/],
  ['sua', /sửa|(^| )sơn( |$)|(^| )thợ( |$)|bóng đèn|ổ cắm|dây điện|ống nước|xi măng|đinh|ốc vít|bản lề|ổ khóa|công thợ|vòi nước|quạt|máy lạnh|ghế|(^| )bàn( |$)|kệ |nồi|chảo|(^| )dao( |$)|thớt|chén|dĩa|đĩa|(^| )tô( |$)|(^| )khay( |$)|rổ|(^| )thau( |$)|xô|vá|keo dán/],
  ['gao', /trứng|bánh mì|bột mì/],
  ['rau', /húng|rau răm|tía tô/],
  ['changa', /chân gà|cánh gà|đùi gà|sụn gà|(^| )gà( |$)|gà ta|gà cắt|rút xương/],
  ['bo', /(^| )bò( |$)|gân bò|bắp bò|(^| )nạm( |$)|thăn bò|(^| )gầu( |$)|đuôi bò|xương bò|lá lách|bắp hoa/],
  ['ca', /(^| )cá( |$)|(^| )tôm( |$)|(^| )mực( |$)|(^| )cua( |$)|(^| )ghẹ( |$)|(^| )ốc( |$)|nghêu|(^| )sò( |$)|hải sản|lươn|(^| )ếch( |$)|chả cá|cá viên/],
  ['giavi', /gia vị|(^| )hành|(^| )tỏi|(^| )ớt|(^| )xay( |$)|(^| )sả( |$)|gừng|tiêu|(^| )muối|(^| )đường|nước mắm|nước tương|dầu ăn|(^| )bột( |$)|hạt nêm|bột ngọt|mì chính|sa tế|nghệ|riềng|(^| )quế( |$)|(^| )hồi( |$)|tương ớt|tương cà|giấm|dầu hào|mắm|(^| )me( |$)|ớt bột|sa tê|tương/],
  ['nuoc', /(^| )nước( |$)|sữa|coca|pepsi|(^| )bia( |$)|(^| )trà( |$)|cà phê|(^| )đá( |$)|sting|7up|number one|(^| )c2( |$)|yaourt|soda|nước mía|nước suối|lavie|aquafina|rượu|lipton|tắc|(^| )lon( |$)|sương sáo/],
  ['baobi', /(^| )(bao|túi|hộp|ly|giấy|màng|ống hút|khăn|nilon|ni lông|găng|bọc|đũa|muỗng|thìa|nĩa|tăm|bột giặt|chổi|cước|dây thun|băng keo|decal|cuộn|khẩu trang|pin)( |$)|nước rửa|nước lau|xà phòng|giấy in|mực in|bao tay|bao rác|bao bì/],
  ['gao', /(^| )gạo( |$)|(^| )bún( |$)|(^| )phở( |$)|bánh tráng|bánh phở|(^| )miến( |$)|(^| )mì( |$)|(^| )nui( |$)|trứng|bánh mì|bột mì|(^| )nếp( |$)|hủ tiếu|bánh canh|(^| )xôi( |$)|bánh/],
  ['rau', /(^| )rau|(^| )củ( |$)|(^| )cải|dưa leo|dưa chuột|cà chua|cà rốt|(^| )cà( |$)|(^| )húng|(^| )ngò|thì là|(^| )hẹ( |$)|(^| )giá( |$)|(^| )đậu|(^| )nấm|khoai|(^| )bí( |$)|chanh|(^| )cam( |$)|dưa hấu|trái cây|chuối|xoài|(^| )thơm( |$)|(^| )dứa( |$)|(^| )bơ( |$)|(^| )dâu( |$)|(^| )nho( |$)|(^| )táo( |$)|(^| )ổi( |$)|mướp|(^| )măng|su hào|su su|xà lách|(^| )lá( |$)|(^| )tía tô|diếp cá|kinh giới|răm|(^| )dưa( |$)|bắp cải|mồng tơi|(^| )mơ( |$)/],
  ['heo', /heo|sườn|ba chỉ|ba rọi|3 rọi|(^| )thịt|(^| )giò( |$)|(^| )chả|(^| )lòng( |$)|(^| )tim( |$)|(^| )cật( |$)|móng giò|(^| )nạc|(^| )dồi|tai heo|(^| )xương|(^| )mỡ( |$)|(^| )nem|(^| )vịt|(^| )ngan|(^| )ngỗng|(^| )dê( |$)|(^| )cừu|lạp xưởng|xúc xích/]
];
function cpnKhoa_(s) { return khongDau_(String(s || '')).toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim(); }
function cpnTab_(ten, head) { return sqTab_(ten, head, function (sh) { sh.getRange('A:C').setNumberFormat('@'); }); }
function cpnNhomDoc_() {
  var sh = cpnTab_(CPN_NHOM, H_CPN_NHOM), n = sh.getLastRow();
  if (n < 2) {
    var luc = bayGio_();
    sh.getRange(2, 1, CPN_MAC_DINH.length, 6).setValues(CPN_MAC_DINH.map(function (g, i) { return [g[0], g[1], g[2], i + 1, '', luc]; }));
    n = sh.getLastRow();
  }
  return sh.getRange(2, 1, n - 1, 6).getValues().map(function (r, i) { return { id: String(r[0]), ten: String(r[1]), icon: String(r[2] || '📌'), tt: so_(r[3]) || (i + 1), gop: String(r[4] || ''), row: i + 2 }; });
}
function cpnLuat_() {
  var sh = cpnTab_(CPN_MON, H_CPN_MON), n = sh.getLastRow(), ex = {}, chua = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, 5).getValues().forEach(function (r, i) {
    var k = String(r[1] || ''); if (!k) return;
    var o = { kieu: String(r[0]), khoa: k, ten: String(r[2] || ''), nhom: String(r[3] || ''), row: i + 2 };
    if (o.kieu === 'chua') chua.push(o); else ex[k] = o;
  });
  chua.sort(function (a, b) { return b.khoa.length - a.khoa.length; });
  return { ex: ex, chua: chua, sh: sh };
}
function cpnDoan_(ten) {
  var t = ' ' + String(ten || '').toLowerCase().normalize('NFC') + ' ';
  for (var i = 0; i < CPN_DOAN.length; i++) if (CPN_DOAN[i][1].test(t)) return CPN_DOAN[i][0];
  return '';
}
// → { g: mã nhóm, src: 't' (anh đã xếp) | 'k' (theo loại chi) | 'd' (app đoán) | 'c' (chưa xếp), ten: tên hiển thị }
function cpnXep_(ten, k, L, gopMap) {
  var kh = cpnKhoa_(ten), r = L.ex[kh], disp = ten, g = '', src = 'c';
  if (r && r.ten) { disp = r.ten; }
  var dk = r && r.ten ? cpnKhoa_(r.ten) : kh;
  var r2 = r || L.ex[dk];
  if (r2 && r2.nhom) { g = r2.nhom; src = 't'; }
  if (!g) for (var i = 0; i < L.chua.length; i++) if (L.chua[i].nhom && (' ' + dk + ' ').indexOf(' ' + L.chua[i].khoa + ' ') >= 0) { g = L.chua[i].nhom; src = 't'; break; }
  if (!g && CPN_THEO_LOAI[k]) { g = CPN_THEO_LOAI[k]; src = 'k'; }
  if (!g) { var d = cpnDoan_(disp) || cpnDoan_(ten); if (d) { g = d; src = 'd'; } }
  if (!g && k === 'khac') { g = 'khac'; src = 'k'; }
  if (g) { for (var s = 0; s < 6 && gopMap[g] && gopMap[g].gop; s++) g = gopMap[g].gop; if (!gopMap[g]) { g = ''; src = 'c'; } }
  return { g: g || 'chua', src: g ? src : 'c', ten: disp };
}
function cpnGopMap_(ds) { var m = {}; ds.forEach(function (g) { m[g.id] = g; }); return m; }

// các dòng chi trong khoảng ngày: [{pi, ngay, ten, sl, dv, tien, k}] + danh sách phiếu
function cpnDong_(tu, den, metas) {
  var ph = [], dong = [], boDauTu = 0;
  sqPhieuTrong_(tu, den).forEach(function (x) {
    if (x.loai === LOAI.BILL || !x.tong) return;
    var parts = []; try { parts = plTach_(Object.assign({}, x, { meta: metas[x.id] || null })); } catch (e) { parts = [{ k: 'nvl', tien: x.tong, dong: x.mon }]; }
    var kp = function (p) { return p.k === 'cd' && p.khoan === 'luong' ? 'luongkhac' : p.k; };
    var kOf = []; parts.forEach(function (p) { (p.dong || []).forEach(function (d) { var i = x.mon.indexOf(d); if (i >= 0) kOf[i] = kp(p); }); });
    var k0 = parts[0] ? kp(parts[0]) : 'nvl', moi = sqCpKhoaMoi_(x).ten;
    var pi = ph.length, tong = 0, co = 0, bo = 0;
    var them = function (ten, sl, dv, tien, k) { if (k === 'dautu') { bo += tien; return; } dong.push([pi, ten, sl, dv, tien, k]); tong += tien; co++; };
    x.mon.forEach(function (d, i) { var t = so_(d.tien); if (!t || !d.ten) return; them(d.ten, d.sl === '' || d.sl == null ? '' : so_(d.sl), d.dv || '', t, kOf[i] || k0); });
    if (!co && !bo) them(moi || x.nhom, '', '', so_(x.tong), k0);
    boDauTu += bo;
    if (co) ph.push([x.id, x.ngay, moi, tong, (x.anh || []).slice(0, 3), x.loai === LOAI.HD ? 'hd' : x.loai === LOAI.TD ? 'td' : x.loai === LOAI_CHU ? 'chu' : 'dc', x.nguoi || '']);
  });
  var doi = sqQuayDoi_();
  sqTnRows_().forEach(function (x) {
    if (x.ngay < tu || x.ngay > den) return;
    var pi = -1, tong = 0;
    (x.chi || []).forEach(function (r) {
      if (!r.amount) return;
      var auto = !CP_LUAN_CHUYEN.test(khongDau_(r.note)), kk = x.ngay + '|' + x.ca + '|' + khongDau_(r.note);
      if (!(doi[kk] === undefined ? auto : doi[kk])) return;
      if (pi < 0) { pi = ph.length; ph.push(['q:' + x.ngay + '|' + x.ca, x.ngay, 'Chi tại quầy · ca ' + x.ca, 0, [], 'quay', x.nguoi || '']); }
      var k = plQuayLoai_(r.note); dong.push([pi, r.note, '', '', r.amount, k]); tong += r.amount; ph[pi][3] = tong;
    });
  });
  return { ph: ph, dong: dong, boDauTu: boDauTu };
}

function sqCpNhom_(p) {
  var hn = homNay_(), N = /^\d{4}-\d{2}-\d{2}$/;
  var den = N.test(String(p.den || '')) ? p.den : hn, tu = N.test(String(p.tu || '')) ? p.tu : den.slice(0, 8) + '01';
  if (tu > den) { var t0 = tu; tu = den; den = t0; }
  if (tu < congNgay_(den, -95)) tu = congNgay_(den, -95);
  var tuT = sqThangLui_(tu), denT = sqThangLui_(den);
  var metas = {}; sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  var ds = cpnNhomDoc_(), gop = cpnGopMap_(ds), L = cpnLuat_();
  var A = cpnDong_(tu, den, metas), B = cpnDong_(tuT, denT, metas);
  var phTruoc = {}; // cùng kỳ tháng trước: chỉ cần tổng theo nhóm
  var truoc = {}, truocTong = 0;
  B.dong.forEach(function (r) { var x = cpnXep_(r[1], r[5], L, gop); truoc[x.g] = (truoc[x.g] || 0) + r[4]; truocTong += r[4]; });
  var dong = A.dong.map(function (r) { var x = cpnXep_(r[1], r[5], L, gop); return [r[0], r[1], x.g, r[2], r[3], r[4], x.src, x.ten === r[1] ? '' : x.ten]; });
  var dt = 0, dtNgay = 0, thieu = 0, dtT = 0;
  try { var m = plNapDT_(tu, den); Object.keys(m).forEach(function (d) { if (m[d]) { dt += m[d].R; dtNgay++; } else thieu++; }); } catch (e) { }
  try { var mt = plNapDT_(tuT, denT); Object.keys(mt).forEach(function (d) { if (mt[d]) dtT += mt[d].R; }); } catch (e) { }
  return { tu: tu, den: den, tuT: tuT, denT: denT, homNay: hn, dt: Math.round(dt), dtNgay: dtNgay, dtThieu: thieu, dtTruoc: Math.round(dtT),
    nhom: ds.filter(function (g) { return !g.gop; }).sort(function (a, b) { return a.tt - b.tt; }).map(function (g) { return { id: g.id, ten: g.ten, icon: g.icon }; }),
    ph: A.ph, d: dong, truoc: truoc, truocTong: truocTong, boDauTu: A.boDauTu, quay: sqCpQuayDs_(tu, den) };
}

// Gán món vào nhóm / đổi tên món (anh Đạt). mon: danh sách tên gốc trên hóa đơn; ten: tên chuẩn mới (bỏ trống = giữ); nho: nhớ cả chữ này
function sqCpNhomGan_(p) {
  var mon = (Array.isArray(p.mon) ? p.mon : []).map(function (s) { return chu_(s, 120); }).filter(function (s) { return cpnKhoa_(s); }).slice(0, 60);
  var nhom = chu_(p.nhom, 20), ten = chu_(p.ten, 120), nho = p.nho === true;
  if (!mon.length) throw new Error('Chưa chọn món.');
  if (nhom && nhom !== 'chua' && !cpnNhomDoc_().some(function (g) { return g.id === nhom; })) throw new Error('Không có nhóm này.');
  if (nhom === 'chua') nhom = '';
  return khoa_(function () {
    var L = cpnLuat_(), sh = L.sh, luc = bayGio_(), n = 0;
    var ghi = function (kieu, khoa, tenChuan, nh) {
      var cu = kieu === 'chua' ? L.chua.filter(function (o) { return o.khoa === khoa; })[0] : L.ex[khoa], row = [kieu, khoa, tenChuan, nh, luc];
      if (cu) sh.getRange(cu.row, 1, 1, 5).setValues([row]);
      else { var r = sh.getLastRow() + 1; sh.getRange(r, 1, 1, 5).setValues([row]); var o = { kieu: kieu, khoa: khoa, ten: tenChuan, nhom: nh, row: r }; if (kieu === 'chua') L.chua.push(o); else L.ex[khoa] = o; }
      n++;
    };
    var tenKhoa = ten ? cpnKhoa_(ten) : '';
    mon.forEach(function (t) {
      var kh = cpnKhoa_(t), cu = L.ex[kh] || {};
      ghi('mon', kh, ten || cu.ten || '', nhom || cu.nhom || '');
    });
    if (ten && tenKhoa && !mon.some(function (t) { return cpnKhoa_(t) === tenKhoa; })) { var cu2 = L.ex[tenKhoa] || {}; ghi('mon', tenKhoa, '', nhom || cu2.nhom || ''); }
    if (nho && nhom) { var kk = cpnKhoa_(ten || mon[0]); if (kk.length >= 3) ghi('chua', kk, '', nhom); }
    return { luu: n };
  });
}

// Danh sách nhóm: đổi tên / icon / thứ tự / thêm nhóm mới / gộp. ds: [{id, ten, icon}] theo thứ tự mới; gop: {tu, vao}
function sqCpNhomDs_(p) {
  var ds = Array.isArray(p.ds) ? p.ds.slice(0, 40) : null, gop = p.gop && p.gop.tu ? { tu: chu_(p.gop.tu, 20), vao: chu_(p.gop.vao, 20) } : null;
  return khoa_(function () {
    var cu = cpnNhomDoc_(), sh = cpnTab_(CPN_NHOM, H_CPN_NHOM), luc = bayGio_(), byId = cpnGopMap_(cu);
    if (gop) {
      if (!byId[gop.tu] || !byId[gop.vao] || gop.tu === gop.vao) throw new Error('Không gộp được 2 nhóm này.');
      var r = byId[gop.tu]; sh.getRange(r.row, 5, 1, 2).setValues([[gop.vao, luc]]);
      var L = cpnLuat_();
      var n = L.sh.getLastRow(); if (n >= 2) { var v = L.sh.getRange(2, 4, n - 1, 1).getValues(), ch = false; v.forEach(function (x) { if (String(x[0]) === gop.tu) { x[0] = gop.vao; ch = true; } }); if (ch) L.sh.getRange(2, 4, n - 1, 1).setValues(v); }
      return { gop: true };
    }
    if (!ds) throw new Error('Thiếu danh sách nhóm.');
    var tt = 0, them = [];
    ds.forEach(function (g) {
      var ten = chu_(g.ten, 40), icon = chu_(g.icon, 4) || '📌'; if (!ten) return; tt++;
      var id = chu_(g.id, 20), r = byId[id];
      if (r) sh.getRange(r.row, 2, 1, 5).setValues([[ten, icon, tt, r.gop, luc]]);
      else them.push(['n' + Date.now().toString(36) + tt, ten, icon, tt, '', luc]);
    });
    if (them.length) sh.getRange(sh.getLastRow() + 1, 1, them.length, 6).setValues(them);
    return { luu: ds.length, moi: them.length };
  });
}
