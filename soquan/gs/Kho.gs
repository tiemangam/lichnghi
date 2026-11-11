/**
 * KHO – số tồn kho Sapo (file "Báo cáo kho hàng") Sơn gửi mỗi tối + file nhiều ngày (tháng) để tính trung bình.
 * Điều hành → tab Kho: cần đặt (đã hết / sắp hết theo mức tối thiểu), xuất dùng so với định mức (báo dùng đột biến),
 * nhập – xuất trong ngày, đối chiếu "Thêm" với hóa đơn mua và Hủy hàng với "Giảm".
 * Anh chỉnh từng món ở tab SQ_KhoCaiDat (qua nút ✎): loại, mức tối thiểu, định mức dùng (vd 1 can trong 4–5 ngày).
 */
var DH_KHO = 'DH_Kho';
var H_DH_KHO = ['Ngày', 'Loại', 'Dữ liệu tới ngày', 'Số ngày', 'Số liệu (JSON)', 'Tải lúc'];
var SQ_KHO_CD = 'SQ_KhoCaiDat';
var H_SQ_KHO_CD = ['Món', 'Loại', 'Mức tối thiểu', 'Định mức SL', 'Trong (ngày) từ', 'đến', 'Đã đặt ngày', 'Sửa lúc', 'Nhóm', 'Quy đổi (1 đv kho = ?)', 'Đv định lượng', 'Hao hụt', 'Bếp đặt lại ngày', 'Bếp đặt lại số'];
var KHO_LOAI = { tuoi: 'Đồ tươi', kho: 'Gia vị – xuất dùng', bb: 'Bao bì – vật dụng', ban: 'Bán sẵn', tp: 'Thành phẩm bếp' };
var KHO_THEO_DOI = ['tuoi', 'kho', 'bb'];   // loại được theo dõi tốc độ dùng
var KHO_CUA_SO = 3;                          // xét 3 ngày gần nhất
var KHO_SO_NGAY = 14;

function sqTabKho_() { return sqTab_(DH_KHO, H_DH_KHO, function (sh) { sh.getRange('A:D').setNumberFormat('@'); sh.getRange('F:F').setNumberFormat('@'); }); }
function sqTabKhoCD_() {
  var sh = sqTab_(SQ_KHO_CD, H_SQ_KHO_CD, function (sh) { sh.getRange('A:B').setNumberFormat('@'); sh.getRange('G:I').setNumberFormat('@'); });
  var n = H_SQ_KHO_CD.length, h = sh.getRange(1, 1, 1, n).getValues()[0];
  if (!h[n - 1]) { sh.getRange(1, 1, 1, n).setValues([H_SQ_KHO_CD]).setFontWeight('bold').setBackground('#445803').setFontColor('#F8E8C4'); sh.getRange(2, 11, 500, 1).setNumberFormat('@'); sh.getRange(2, 13, 500, 1).setNumberFormat('@'); }
  return sh;
}

// [tên, đv, đầu, bán, thêm, giảm, tối thiểu Sapo, cuối]
function sqKhoMon_(ds) {
  return (Array.isArray(ds) ? ds : []).slice(0, 500).map(function (m) {
    m = Array.isArray(m) ? m : [];
    return [chu_(m[0], 70), chu_(m[1], 15), so_(m[2]), so_(m[3]), so_(m[4]), so_(m[5]), so_(m[6]), so_(m[7])];
  }).filter(function (m) { return m[0]; });
}
// lưu (gọi trong khoa_ của sqSapoLuu_)
function sqKhoLuu_(list, luc) {
  if (!list.length) return 0;
  var sh = sqTabKho_(), idx = {};
  sqKhoMeta_().forEach(function (x) { idx[x.ngay + '|' + x.loai] = x.row; });
  list.forEach(function (x) {
    var ngay = sqKiemNgay_(x.ngay), loai = x.loai === 'khoky' ? 'khoky' : 'kho', d = x.d || {};
    var mon = sqKhoMon_(d.mon);
    if (!mon.length) throw new Error('File tồn kho không có món nào.');
    var den = loai === 'kho' ? ngay : (/^\d{4}-\d{2}-\d{2}$/.test(String(d.den || '')) ? String(d.den) : ngay);
    var n = loai === 'kho' ? 1 : Math.max(1, Math.round(so_(d.n) * 100) / 100);
    var row = idx[ngay + '|' + loai];
    if (!row) { row = sh.getLastRow() + 1; idx[ngay + '|' + loai] = row; }
    sh.getRange(row, 1, 1, 6).setValues([[ngay, loai, den, String(n), JSON.stringify(mon), luc]]);
  });
  return list.length;
}
function sqKhoMeta_() {
  var sh = sqTabKho_(), n = sh.getLastRow();
  if (n < 2) return [];
  var a = sh.getRange(2, 1, n - 1, 4).getValues(), f = sh.getRange(2, 6, n - 1, 1).getValues();
  return a.map(function (r, i) { return { row: i + 2, ngay: chuanNgay_(r[0]), loai: String(r[1]), den: chuanNgay_(r[2]), n: so_(r[3]), luc: String(f[i][0] || '') }; })
    .filter(function (x) { return /^\d{4}-\d{2}-\d{2}$/.test(x.ngay) && (x.loai === 'kho' || x.loai === 'khoky'); });
}
function sqKhoDoc_(rows) {
  var sh = sqTabKho_(), out = {};
  rows.forEach(function (x) { out[x.row] = sqKhoMon_(sqJ_(String(sh.getRange(x.row, 5).getValue() || ''), [])); });
  return out;
}

/* ---------- cài đặt từng món ---------- */
function sqKhoK_(t) { return khongDau_(t); }
function sqKhoCD_() {
  var sh = sqTabKhoCD_(), n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, H_SQ_KHO_CD.length).getValues().forEach(function (r, i) {
    var t = String(r[0] || ''); if (!t) return;
    var sr = function (v) { return v === '' || v === null ? '' : so_(v); };
    m[sqKhoK_(t)] = { row: i + 2, ten: t, loai: KHO_LOAI[String(r[1])] ? String(r[1]) : '', min: sr(r[2]), sl: sr(r[3]), n1: sr(r[4]), n2: sr(r[5]), dat: chuanNgay_(r[6]) || '', nhom: String(r[8] || ''), quy: sr(r[9]), dvg: String(r[10] || ''), hao: sr(r[11]), bepNgay: chuanNgay_(r[12]) || '', bepSl: r[13] === '' || r[13] === null ? '' : Number(r[13]) };
  });
  return m;
}
function sqKhoCDMoi_(ten) { return { ten: ten, loai: '', min: '', sl: '', n1: '', n2: '', dat: '', nhom: '', quy: '', dvg: '', hao: '', bepNgay: '', bepSl: '' }; }
function sqKhoCDDong_(ten, o) { var v = function (x) { return x === undefined || x === null ? '' : x; }; return [ten, o.loai, o.min, o.sl, o.n1, o.n2, o.dat, bayGio_(), v(o.nhom), v(o.quy), v(o.dvg), v(o.hao), v(o.bepNgay), v(o.bepSl)]; }
function sqKhoCDLuu_(p) {
  var ten = chu_(p.ten, 70); if (!ten) throw new Error('Thiếu tên món.');
  var so0 = function (v, am) { if (v === '' || v === null || v === undefined) return ''; var x = Number(v); if (!isFinite(x) || x < 0 || (!am && x === 0)) return ''; return Math.round(x * 1000) / 1000; };
  return khoa_(function () {
    var sh = sqTabKhoCD_(), cd = sqKhoCD_(), x = cd[sqKhoK_(ten)];
    var o = x || sqKhoCDMoi_(ten);
    if ('loai' in p) o.loai = KHO_LOAI[String(p.loai)] ? String(p.loai) : '';
    if ('min' in p) o.min = so0(p.min, true);
    if ('sl' in p) o.sl = so0(p.sl);
    if ('n1' in p) o.n1 = so0(p.n1);
    if ('n2' in p) o.n2 = so0(p.n2);
    if ('nhom' in p) o.nhom = /^[a-z0-9_]{1,20}$/.test(String(p.nhom || '')) ? String(p.nhom) : '';
    if ('quy' in p) o.quy = so0(p.quy);
    if ('dvg' in p) o.dvg = chu_(p.dvg, 12);
    if ('hao' in p) { var hh = Number(p.hao); o.hao = isFinite(hh) && hh > 0 && hh < 0.95 ? Math.round(hh * 1000) / 1000 : ''; }
    if ('bepSl' in p) { var bs = p.bepSl === '' || p.bepSl === null || p.bepSl === undefined ? NaN : Number(String(p.bepSl).replace(',', '.')); if (isFinite(bs)) { o.bepSl = Math.round(bs * 1000) / 1000; o.bepNgay = sqKiemNgay_(p.bepNgay); } else { o.bepSl = ''; o.bepNgay = ''; } }
    if ('dat' in p) o.dat = /^\d{4}-\d{2}-\d{2}$/.test(String(p.dat || '')) ? String(p.dat) : '';
    if (o.sl === '' || o.n1 === '') { o.sl = ''; o.n1 = ''; o.n2 = ''; }
    if (o.n2 !== '' && o.n2 < o.n1) { var t = o.n1; o.n1 = o.n2; o.n2 = t; }
    var row = x ? x.row : sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, H_SQ_KHO_CD.length).setValues([sqKhoCDDong_(ten, o)]);
    return { ok: true };
  });
}

/* ---------- tính ---------- */
var KHO_BB = ['boc', 'hop', 'ly ', 'nap', 'ong hut', 'muong', 'dua', 'giay', 'tem', 'ghim', 'khan', 'tam', 'bao tay', 'gas', 'nhang', 'nuoc rua', 'hu sot', 'vim'];
var KHO_TUOI = ['ga nguyen con', 'sun', 'thit heo', 'trung cut', 'trung ga', 'trung muoi', 'trung muc', 'hen', 'nem', 'cha', 'bo', 'sua hop'];
function sqKhoLoaiDoan_(t, coBan) {
  var k = sqKhoK_(t);
  if (/\(gia thuong\)/.test(k)) return 'ban';
  if (KHO_BB.some(function (b) { return k.indexOf(b) === 0; })) return 'bb';
  if (KHO_TUOI.indexOf(k) >= 0) return 'tuoi';
  if (coBan) return 'tp';
  return 'kho';
}
/* ---------- NHÓM HIỂN THỊ (thứ tự từ trên xuống; anh sắp xếp / đổi tên / chuyển món bằng nút ✎ Sắp xếp nhóm) ---------- */
// kw: khớp đầu tên (theo cả chữ); '=x' = đúng nguyên tên
var KHO_NHOM_MD = [
  { id: 'thit', ten: 'Thịt heo – bò', mo: true, kw: ['thit', 'bo goi', 'bo met', 'bo nhung', 'bo vien', 'bo tai', 'bo bit tet'] },
  { id: 'ga', ten: 'Gà – vịt', mo: true, kw: ['ga nguyen con', 'ga dot', 'ga lau', 'ga', 'dui ga', 'dui vit', 'vit', 'cari ga'] },
  { id: 'chan', ten: 'Chân gà – sụn', mo: true, kw: ['chan sa', 'chan thai', 'chan ga', 'rut sa', 'rut thai', 'sun'] },
  { id: 'hs', ten: 'Cá – mực – hến', mo: true, kw: ['ca', 'rang muc', 'trung muc', 'muc', 'hen', 'tom'] },
  { id: 'cha', ten: 'Chả – nem', mo: true, kw: ['cha', 'nem', 'moc'] },
  { id: 'trung', ten: 'Trứng', mo: true, kw: ['trung'] },
  { id: 'lau', ten: 'Nước lẩu', mo: true, kw: ['nuoc lau'] },
  { id: 'hanh', ten: 'Hành phi – đậu phộng', mo: true, kw: ['hanh phi', 'dau phong'] },
  { id: 'chinh', ten: 'Món chính khác', mo: true, kw: ['xoi chien'] },
  { id: 'trang', ten: 'Tráng miệng', mo: false, kw: ['kem chuoi', 'kem', 'banh flan', 'sua chua'] },
  { id: 'nuoc', ten: 'Nước uống', mo: false, kw: ['pepsi', '7 up', '7up', 'mirinda', 'sting', 'nuoc suoi', 'tiger', 'strongbow', 'coca', 'sprite', 'bia', 'heineken', 'saigon', '333', 'tra', 'nuoc ngot', 'redbull', 'red bull'] },
  { id: 'kho', ten: 'Bánh – bột – đồ khô', mo: false, kw: ['banh trang', 'banh da', 'banh', 'mi ki', 'mi', 'bun', 'bot mi', 'bot nang', 'bot nep', 'bot chien', 'sua dac', 'sua hop', 'sua', '=bo', 'keo'] },
  { id: 'gv', ten: 'Gia vị', mo: false, kw: ['giam', 'hat nem', 'nuoc mam', 'mam', 'bot ngot', 'duong', 'muoi', 'dau an', 'dau hao', 'nuoc tuong', 'tuong ot', 'tuong', 'sate', 'chao', 'cot chanh', 'bot ot', 'bot cari', 'bot', 'ot mam gung', 'ot', 'tieu', 'toi', 'hanh', 'sa'] },
  { id: 'bb', ten: 'Bao bì – vật dụng', mo: false, kw: ['boc', 'hop', 'ly', 'nap', 'ong hut', 'muong', 'dua', 'giay', 'tem', 'ghim', 'khan', 'tam', 'bao tay', 'gas', 'nhang', 'nuoc rua', 'hu sot', 'vim', 'tui', 'chen', 'dia', 'to'] },
  { id: 'khac', ten: 'Khác', mo: false, kw: [] }
];
function sqKhoKw_(k, w) { if (w.charAt(0) === '=') return k === w.slice(1); return k === w || k.indexOf(w + ' ') === 0; }
function sqKhoNhomDoan_(ten) {
  var k = sqKhoK_(ten).replace(/\(gia thuong\)/g, '').replace(/\s+/g, ' ').trim();
  for (var i = 0; i < KHO_NHOM_MD.length; i++) if (KHO_NHOM_MD[i].kw.some(function (w) { return sqKhoKw_(k, w); })) return KHO_NHOM_MD[i].id;
  return 'khac';
}
function sqKhoNhom_() {
  var v = PropertiesService.getScriptProperties().getProperty('KHO_NHOM'), luu = v ? sqJ_(v, null) : null, out = [], co = {};
  (Array.isArray(luu) ? luu : []).forEach(function (g) { if (g && g.id && !co[g.id]) { co[g.id] = 1; out.push({ id: String(g.id), ten: String(g.ten || g.id), mo: !!g.mo }); } });
  KHO_NHOM_MD.forEach(function (g) { if (!co[g.id]) { co[g.id] = 1; out.push({ id: g.id, ten: g.ten, mo: g.mo }); } });
  return out;
}
function sqKhoNhomLuu_(p) {
  var ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 40).map(function (g) {
    return { id: /^[a-z0-9_]{1,20}$/.test(String(g && g.id || '')) ? String(g.id) : '', ten: chu_(g && g.ten, 40), mo: !!(g && g.mo) };
  }).filter(function (g) { return g.id && g.ten; });
  if (!ds.length) PropertiesService.getScriptProperties().deleteProperty('KHO_NHOM');
  else PropertiesService.getScriptProperties().setProperty('KHO_NHOM', JSON.stringify(ds));
  return { ok: true, nhomDs: sqKhoNhom_() };
}
// nhóm của 1 món: anh đã chuyển (cột Nhóm) → theo đó, không thì app đoán theo tên
function sqKhoNhomCua_(ten, s, ids) { return s && s.nhom && ids[s.nhom] ? s.nhom : sqKhoNhomDoan_(ten); }
function sqKhoIds_(ds) { var m = {}; ds.forEach(function (g) { m[g.id] = 1; }); return m; }
function sqKhoTok_(s) { return khongDau_(String(s || '').replace(/\(.*?\)/g, ' ')).replace(/[^a-z0-9 ]/g, ' ').split(' ').filter(function (w) { return w.length >= 2; }); }
// điểm khớp tên: 2 = tên kho nằm trọn trong tên dòng, 1 = tên dòng nằm trọn trong tên kho (dòng có ≥2 chữ hoặc kho 1 chữ)
function sqKhoDiem_(a, b) {
  if (!a.length || !b.length) return 0;
  if (a.every(function (w) { return b.indexOf(w) >= 0; })) return 2 + a.length / 10;
  if (b.every(function (w) { return a.indexOf(w) >= 0; }) && (b.length >= 2 || a.length === 1)) return 1 + b.length / 10;
  return 0;
}
function sqKhoNgayDs_(den, n) { var out = []; for (var i = n - 1; i >= 0; i--) out.push(congNgay_(den, -i)); return out; }
function sqKhoSoNgay_(a, b) { var p = a.split('-'), q = b.split('-'); return Math.round((Date.UTC(+q[0], +q[1] - 1, +q[2]) - Date.UTC(+p[0], +p[1] - 1, +p[2])) / 864e5); }

function sqKhoTinh_(p) {
  p = p || {};
  var meta = sqKhoMeta_(), ngays = meta.filter(function (x) { return x.loai === 'kho'; }).map(function (x) { return x.ngay; }).sort();
  var kys = meta.filter(function (x) { return x.loai === 'khoky'; }).map(function (x) { return { tu: x.ngay, den: x.den, n: x.n, row: x.row }; });
  // chưa có file 1 ngày nào mới hơn file nhiều ngày → lấy tồn cuối kỳ của file nhiều ngày mới nhất làm số tồn
  var kyMoi = kys.slice().sort(function (a, b) { return a.den < b.den ? 1 : a.den > b.den ? -1 : b.n - a.n; })[0], gia = null;
  if (kyMoi && (!ngays.length || kyMoi.den > ngays[ngays.length - 1])) {
    gia = { loai: 'gia', ngay: kyMoi.den, row: kyMoi.row, tu: kyMoi.tu, luc: (meta.filter(function (x) { return x.row === kyMoi.row; })[0] || {}).luc || '' };
    ngays.push(kyMoi.den);
  }
  if (!ngays.length) return { trong: true, ngays: [], ky: kys.map(function (k) { return { tu: k.tu, den: k.den, n: k.n }; }), loaiTen: KHO_LOAI, nhomDs: sqKhoNhom_() };
  var ngay = /^\d{4}-\d{2}-\d{2}$/.test(String(p.ngay || '')) ? String(p.ngay) : ngays[ngays.length - 1];
  var truoc = ngays.filter(function (d) { return d <= ngay; });
  ngay = truoc.length ? truoc[truoc.length - 1] : ngays[0];
  var days = sqKhoNgayDs_(ngay, KHO_SO_NGAY);
  var can = meta.filter(function (x) { return x.loai === 'kho' && x.ngay >= days[0] && x.ngay <= ngay; });
  if (gia && gia.ngay >= days[0] && gia.ngay <= ngay) can.push(gia);
  // các file nhiều ngày trong 75 ngày, không chồng nhau (ưu tiên file mới)
  kys = kys.filter(function (k) { return k.den >= congNgay_(ngay, -75) && k.tu <= ngay; }).sort(function (a, b) { return a.tu < b.tu ? 1 : -1; });
  var dung = [], moc = '9999-12-31';
  kys.forEach(function (k) { if (k.den < moc) { dung.push(k); moc = k.tu; } });
  var data = sqKhoDoc_(can.concat(dung));
  var snap = {}, snapGia = {}, lucNgay = '', curGia = null;
  can.forEach(function (x) {
    var rows = data[x.row] || [];
    if (x.loai === 'gia') { rows = rows.map(function (a) { return [a[0], a[1], a[7], 0, 0, 0, a[6], a[7]]; }); snapGia[x.ngay] = true; if (x.ngay === ngay) curGia = rows; }
    var m = {}; rows.forEach(function (a) { m[sqKhoK_(a[0])] = a; }); snap[x.ngay] = m; if (x.ngay === ngay) lucNgay = x.luc; });
  var kyTong = {}, kyN = 0, kyTu = '', kyDen = '';
  dung.forEach(function (k) {
    kyN += k.n; if (!kyTu || k.tu < kyTu) kyTu = k.tu; if (!kyDen || k.den > kyDen) kyDen = k.den;
    (data[k.row] || []).forEach(function (a) { var key = sqKhoK_(a[0]), o = kyTong[key] || (kyTong[key] = { dung: 0, ban: 0 }); o.dung += -(a[3] + a[5]); o.ban += -a[3]; });
  });
  var nhomDs = sqKhoNhom_(), nhomIds = sqKhoIds_(nhomDs);
  var cd = sqKhoCD_(), cur = curGia || data[can.filter(function (x) { return x.ngay === ngay; })[0].row] || [];
  var cuaSo = days.slice(-KHO_CUA_SO), nen = days.slice(0, KHO_SO_NGAY - KHO_CUA_SO);
  var r1 = function (v) { return Math.round(v * 100) / 100; };

  var items = cur.map(function (a) {
    var key = sqKhoK_(a[0]), s = cd[key] || {};
    var d14 = days.map(function (d) { var x = !snapGia[d] && snap[d] && snap[d][key]; return x ? r1(-(x[3] + x[5])) : null; });
    var n14 = days.map(function (d) { var x = !snapGia[d] && snap[d] && snap[d][key]; return x ? [x[2], x[4], r1(-(x[3] + x[5])), x[7]] : null; });
    var coBan = (kyTong[key] && kyTong[key].ban > 0) || days.some(function (d) { var x = snap[d] && snap[d][key]; return x && x[3] < 0; });
    var loaiDoan = sqKhoLoaiDoan_(a[0], coBan), loai = s.loai || loaiDoan;
    // tốc độ dùng tự tính
    var nenCo = nen.map(function (d, i) { return d14[i]; }).filter(function (v) { return v !== null; });
    var rateTu = null, nguon = '';
    if (nenCo.length >= 7) { rateTu = nenCo.reduce(function (x, y) { return x + y; }, 0) / nenCo.length; nguon = 'TB ' + nenCo.length + ' ngày gần nhất'; }
    else if (kyN && kyTong[key]) { rateTu = kyTong[key].dung / kyN; nguon = 'TB ' + sqDM_(kyTu) + '–' + sqDM_(kyDen) + ' (' + Math.round(kyN) + ' ngày)'; }
    var dm = s.sl !== '' && s.sl > 0 && s.n1 > 0 ? { sl: s.sl, n1: s.n1, n2: s.n2 === '' ? s.n1 : s.n2 } : null;
    var rate = dm ? dm.sl / ((dm.n1 + dm.n2) / 2) : rateTu;
    // 3 ngày gần nhất
    var w = cuaSo.map(function (d, i) { return d14[KHO_SO_NGAY - KHO_CUA_SO + i]; }).filter(function (v) { return v !== null; });
    var dung3 = w.length ? r1(w.reduce(function (x, y) { return x + y; }, 0)) : null, exp3 = rate != null ? r1(rate * w.length) : null, muc = '', tile = null;
    if (KHO_THEO_DOI.indexOf(loai) >= 0 && w.length >= 2 && rate > 0 && dung3 != null) {
      tile = Math.round(dung3 / exp3 * 100) / 100;
      if (dung3 - exp3 >= 1) muc = tile >= 2 ? 'r' : tile >= 1.5 ? 'y' : '';
    }
    var min = s.min !== '' && s.min !== undefined ? s.min : a[6], cuoi = a[7];
    return { t: a[0], dv: a[1], dau: a[2], ban: a[3], them: a[4], giam: a[5], minSapo: a[6], min: min, minAnh: s.min !== '' && s.min !== undefined, cuoi: cuoi,
      tt: cuoi <= 0 ? 'het' : cuoi <= min ? 'sap' : '', loai: loai, loaiDoan: loaiDoan, dm: dm, rate: rate == null ? null : Math.round(rate * 1000) / 1000, rateTu: rateTu == null ? null : Math.round(rateTu * 1000) / 1000, nguon: nguon,
      nhom: sqKhoNhomCua_(a[0], s, nhomIds), nhomDoan: sqKhoNhomDoan_(a[0]), quy: s.quy || '', dvg: s.dvg || '', hao: s.hao || '', d14: d14, n14: n14, soNgay3: w.length, dung3: dung3, exp3: exp3, tile: tile, muc: muc, duNgay: rate > 0 ? Math.round(cuoi / rate * 10) / 10 : null, dat: !!(s.dat && s.dat >= ngay) };
  });

  // nguyên liệu đã có món liên quan → theo dõi bằng "Xuất dùng so với món bán" (tồn bếp), không dùng định mức theo ngày
  // nguyên liệu còn lại (gia vị, bao bì…) → so lượng dùng với doanh thu thay vì theo ngày
  try {
    var dlm = sqDL_(), dtN = {};
    sqSapoRows_().forEach(function (x) { if (x.ngay < days[0] || x.ngay > ngay) return;
      if (x.loai === 'tt' && x.d.tong > 0) dtN[x.ngay] = x.d.tong; else if (x.loai === 'tq' && x.d.dtThuc > 0 && dtN[x.ngay] == null) dtN[x.ngay] = x.d.dtThuc; });
    items.forEach(function (it, ii) {
      var key = sqKhoK_(it.t), lk = (dlm[key] || []).some(function (l) { return l.bat && l.sl > 0; });
      if (lk) { it.lk = true; it.muc = ''; it.tile = null; return; }
      if (KHO_THEO_DOI.indexOf(it.loai) < 0) return;
      var u3 = 0, d3 = 0, n3 = 0, ub = 0, db = 0;
      days.forEach(function (d, i) { var v = it.d14[i]; if (v == null || !(dtN[d] > 0)) return;
        if (i >= KHO_SO_NGAY - KHO_CUA_SO) { u3 += v; d3 += dtN[d]; n3++; } else { ub += v; db += dtN[d]; } });
      if (n3 < 2 || !(d3 > 0) || !(db > 0) || !(ub > 0)) return;
      var p3 = u3 / (d3 / 1e7), pb = ub / (db / 1e7), ex = pb * d3 / 1e7, tl = p3 / pb;
      it.p10 = r1(p3); it.p10b = r1(pb); it.dt3 = Math.round(d3); it.exp3 = r1(ex); it.tile = Math.round(tl * 100) / 100; it.theoDt = true;
      it.muc = u3 - ex >= 1 ? (tl >= 2 ? 'r' : tl >= 1.5 ? 'y' : '') : '';
    });
  } catch (e) { }

  // đối chiếu "Thêm" ↔ hóa đơn mua (ngày trước + ngày đó)
  var them = [], huy = [];
  var dongMua = [];
  try {
    sqPhieuTrong_(congNgay_(ngay, -1), ngay).forEach(function (x) {
      if (x.loai === LOAI.BILL) return;
      var ten = ''; try { ten = sqCpKhoaMoi_(x).ten; } catch (e) { ten = x.moi || x.nhom || x.nguoi || ''; }
      (x.mon || []).forEach(function (m) { dongMua.push({ tok: sqKhoTok_(m.ten), ten: m.ten, sl: m.sl, dv: m.dv, moi: ten, ngay: x.ngay, id: x.id }); });
    });
  } catch (e) { }
  items.filter(function (it) { return it.them > 0; }).forEach(function (it) {
    if (it.loai === 'tp') { them.push({ t: it.t, sl: it.them, dv: it.dv, s: 'b', txt: 'Bếp nấu / sơ chế, không cần hóa đơn' }); return; }
    var tk = sqKhoTok_(it.t), best = [], diem = 0;
    dongMua.forEach(function (d) { var sc = sqKhoDiem_(tk, d.tok); if (sc > diem) { diem = sc; best = [d]; } else if (sc && sc === diem) best.push(d); });
    if (best.length) them.push({ t: it.t, sl: it.them, dv: it.dv, s: 'ok', id: best[0].id, txt: best.slice(0, 2).map(function (d) { return d.moi + ' · ' + d.ten + (d.sl !== '' ? ' ' + String(d.sl).replace('.', ',') + ' ' + d.dv : '') + (d.ngay !== ngay ? ' (' + sqDM_(d.ngay) + ')' : ''); }).join(' + ') });
    else them.push({ t: it.t, sl: it.them, dv: it.dv, s: 'no', txt: 'Không thấy hóa đơn mua ' + sqDM_(congNgay_(ngay, -1)) + '–' + sqDM_(ngay) });
  });
  // đối chiếu Hủy hàng ↔ "Giảm"
  try {
    var itTok = items.map(function (it) { return sqKhoTok_(it.t); });
    sqHuyRows_().filter(function (h) { return h.ngay === ngay; }).forEach(function (h) {
      var tk = sqKhoTok_(h.ten), bi = -1, diem = 0;
      itTok.forEach(function (a, i) { var sc = sqKhoDiem_(a, tk); if (sc > diem) { diem = sc; bi = i; } });
      var it = bi >= 0 ? items[bi] : null, mo = h.ten + ' ' + String(h.sl).replace('.', ',') + ' ' + h.dv + (h.lyDo ? ' · ' + h.lyDo : '') + (h.nguoi ? ' · ' + h.nguoi : '');
      if (!it) huy.push({ t: h.ten, s: 'i', txt: 'Hủy ' + mo + ' – không thấy món này trên kho Sapo' });
      else if (it.giam < 0) huy.push({ t: it.t, sl: it.giam, dv: it.dv, s: 'ok', txt: 'Hủy hàng: ' + mo });
      else huy.push({ t: it.t, sl: 0, dv: it.dv, s: 'no', txt: 'Có phiếu hủy (' + mo + ') nhưng Sapo chưa trừ – nhắc Sơn bấm Giảm' });
    });
  } catch (e) { }

  return { tuKy: curGia ? { tu: gia.tu, den: gia.ngay } : null, ngay: ngay, ngays: ngays.slice(-60), luc: lucNgay, homNay: homNay_(), days: days, items: items, doiChieu: { them: them, huy: huy },
    ky: kyN ? { tu: kyTu, den: kyDen, n: kyN } : null, loaiTen: KHO_LOAI, cuaSo: KHO_CUA_SO, nhomDs: nhomDs };
}

// Điều hành
function sqKhoChu_(p) { return sqKhoTinh_(p); }

// việc cần xử lý cho Tổng quan
function sqKhoViec_() {
  var j = sqKhoTinh_({}), out = [];
  if (j.trong || j.ngay < congNgay_(homNay_(), -2)) return out;
  var het = j.items.filter(function (x) { return x.tt === 'het' && !x.dat; }), sap = j.items.filter(function (x) { return x.tt === 'sap' && !x.dat; });
  j.items.filter(function (x) { return x.muc === 'r'; }).sort(function (a, b) { return b.tile - a.tile; }).slice(0, 4).forEach(function (x) {
    out.push({ muc: 'r', loai: 'Kho – dùng nhiều', tab: 'kho', ngay: j.ngay, text: x.t + ': dùng ' + String(x.dung3).replace('.', ',') + ' ' + String(x.dv).toLowerCase() + ' trong ' + x.soNgay3 + ' ngày – gấp ' + String(Math.round(x.tile * 10) / 10).replace('.', ',') + ' lần ' + (x.theoDt ? 'so với doanh thu' : 'định mức') });
  });
  try {
    var t = sqKhoTuan_({});
    t.items.filter(function (x) { return x.muc === 'r'; }).slice(0, 4).forEach(function (x) {
      var T = t.tuanInfo[t.tuanInfo.length - 1];
      out.push({ muc: 'r', loai: 'Kho – tuần ' + sqDM_(T.tu) + '–' + sqDM_(T.den), tab: 'kho', ngay: T.den, text: x.t + ': dùng ' + String(x.dung).replace('.', ',') + ' ' + String(x.dv).toLowerCase() + ' – gấp ' + String(Math.round(x.tile * 10) / 10).replace('.', ',') + ' lần so với ' + (x.dau && x.ppcBase ? 'số món chiên bán' : 'doanh thu') + ' (4 tuần trước)' });
    });
  } catch (e) { }
  try {
    var bp = sqKhoBep_({});
    var so = function (v) { return String(v).replace('.', ','); }, hn2 = homNay_();
    if (!bp.trong) bp.items.forEach(function (x) {
      var dv = String(x.dv).toLowerCase();
      (x.kiemDs || []).filter(function (e) { return e.ngay >= congNgay_(hn2, -2) && e.lech !== '' && e.lech <= -Math.max(0.25, 0.1 * (x.avg || 0)); }).slice(0, 1).forEach(function (e) {
        out.push({ muc: e.ly && e.ly !== 'khong' ? 'i' : 'r', loai: 'Kho – hụt tủ bếp', tab: 'kho', ngay: e.ngay, text: x.t + ': ' + (e.nguoi || '') + ' kiểm tủ ' + (e.luc === 'toi' ? 'tối ' + sqDM_(e.ngay) : 'sáng ' + sqDM_(congNgay_(e.ngay, 1))) + ' còn ' + so(e.dem) + ' ' + dv + ', app tính ' + so(e.du) + ' → hụt ' + so(-e.lech) + ' ' + dv + ' (' + (KT_LYDO[e.ly] || 'chưa ghi lý do').toLowerCase() + ')' });
      });
    });
    if (!bp.trong && bp.ngay >= congNgay_(hn2, -2)) bp.items.filter(function (x) { return !x.tuTru && x.bep != null && x.muc === 'r' && /^âm/.test(x.nhan); }).slice(0, 3).forEach(function (x) {
      var dv = String(x.dv).toLowerCase();
      out.push({ muc: 'r', loai: 'Kho – bếp', tab: 'kho', ngay: bp.ngay, text: x.t + ': bán nhiều hơn số đã xuất kho ' + so(-x.bep) + ' ' + dv + ' (quên bấm xuất, hoặc định lượng chưa đúng)' });
    });
  } catch (e) { }
  var vang = j.items.filter(function (x) { return x.muc === 'y'; }).length;
  if (het.length || sap.length || vang) out.push({ muc: het.length ? 'y' : 'i', loai: 'Kho', tab: 'kho', ngay: j.ngay,
    text: sqDM_(j.ngay) + ': ' + [het.length ? het.length + ' món đã hết' : '', sap.length ? sap.length + ' món sắp hết' : '', vang ? vang + ' món dùng nhiều hơn thường lệ' : ''].filter(String).join(' · ') });
  return out;
}

/* ---------- THEO TUẦN (Thứ 2 → Chủ nhật): lượng dùng so với doanh thu, dầu ăn so với món chiên ---------- */
var KHO_CHIEN_MAC_DINH = ['sun ga', 'chan ga chien', 'trung muc chien', 'cha cuon ca trich', 'ga dot', 'com chien', 'lau vit nau chao', 'goi chay', 'ca loc'];
var KHO_DAU = ['dau an'];
function sqThu2_(d) { var p = d.split('-'), w = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay(); return congNgay_(d, -((w + 6) % 7)); }
function sqKhoChien_() { var v = PropertiesService.getScriptProperties().getProperty('KHO_MON_CHIEN'); var ds = v ? sqJ_(v, null) : null; return Array.isArray(ds) && ds.length ? ds : KHO_CHIEN_MAC_DINH; }
function sqKhoChienLuu_(p) {
  var ds = (Array.isArray(p.ds) ? p.ds : []).map(function (x) { return khongDau_(chu_(x, 80)); }).filter(String).slice(0, 80);
  PropertiesService.getScriptProperties().setProperty('KHO_MON_CHIEN', JSON.stringify(ds));
  return { ok: true, ds: ds };
}
function sqKhoLaChien_(ten, ds) { var k = khongDau_(ten); return ds.some(function (p) { return k === p || k.indexOf(p) >= 0; }); }

function sqKhoTuan_(p) {
  p = p || {};
  var hn = homNay_(), tuNay = sqThu2_(hn);
  var tuan = /^\d{4}-\d{2}-\d{2}$/.test(String(p.tuan || '')) ? sqThu2_(String(p.tuan)) : '';
  if (!tuan) { // mặc định: tuần mới nhất có số kho
    var moi = ''; sqKhoMeta_().forEach(function (x) { var d = x.loai === 'kho' ? x.ngay : (sqThu2_(x.ngay) === x.ngay && x.den === congNgay_(x.ngay, 6) ? x.den : ''); if (d && d > moi) moi = d; });
    tuan = moi ? sqThu2_(moi > hn ? hn : moi) : tuNay;
  }
  var SO = 12, mons = []; for (var i = SO - 1; i >= 0; i--) mons.push(congNgay_(tuan, -7 * i));
  var start = mons[0], end = congNgay_(tuan, 6);
  // kho: đọc 1 lần
  var sh = sqTabKho_(), n = sh.getLastRow(), daily = {}, weekly = {}, ten = {};
  if (n >= 2) {
    var a = sh.getRange(2, 1, n - 1, 5).getValues();
    a.forEach(function (r) {
      var ngay = chuanNgay_(r[0]), loai = String(r[1]), den = chuanNgay_(r[2]);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(ngay)) return;
      var hop = loai === 'kho' ? (ngay >= start && ngay <= end) : (loai === 'khoky' && ngay >= start && ngay <= tuan && sqThu2_(ngay) === ngay && den === congNgay_(ngay, 6));
      if (!hop) return;
      var m = {}; sqKhoMon_(sqJ_(String(r[4] || ''), [])).forEach(function (x) { var k = sqKhoK_(x[0]); m[k] = x; ten[k] = x; });
      if (loai === 'kho') daily[ngay] = m; else weekly[ngay] = m;
    });
  }
  // doanh thu + món bán
  var dt = {}, don = {}, monNgay = {}, monTuan = {}, tenMon = {}, monGop = [];
  sqSapoRows_().forEach(function (x) {
    if (x.loai === 'tt' && x.ngay >= start && x.ngay <= end) { dt[x.ngay] = x.d.tong; don[x.ngay] = x.d.don; }
    else if (x.loai === 'tq' && x.ngay >= start && x.ngay <= end) { if (dt[x.ngay] == null) dt[x.ngay] = x.d.dtThuc; if (don[x.ngay] == null) don[x.ngay] = x.d.don; }
    else if (x.loai === 'mon' && x.ngay >= start && x.ngay <= end) monNgay[x.ngay] = x.d.mon || [];
    else if (x.loai === 'monkg' && x.d.den && sqThu2_(x.ngay) === x.ngay && sqThu2_(x.d.den) === congNgay_(x.d.den, -6) && x.ngay >= start && x.d.den < tuan) monGop.push({ tu: x.ngay, den: x.d.den, mon: x.d.mon || [] });
    else if (x.loai === 'monky' && x.ngay >= start && x.ngay <= tuan && sqThu2_(x.ngay) === x.ngay && x.d.den === congNgay_(x.ngay, 6)) monTuan[x.ngay] = x.d;
  });
  var chienDs = sqKhoChien_();
  var tuanInfo = mons.map(function (m) {
    var days = [0, 1, 2, 3, 4, 5, 6].map(function (k) { return congNgay_(m, k); });
    var dKho = days.filter(function (d) { return daily[d]; }), dDt = days.filter(function (d) { return dt[d] != null; }), dMon = days.filter(function (d) { return monNgay[d]; });
    var mt = monTuan[m];
    var doanhThu = dDt.length ? dDt.reduce(function (s, d) { return s + dt[d]; }, 0) : null, soDon = dDt.length ? dDt.reduce(function (s, d) { return s + (don[d] || 0); }, 0) : null;
    if (mt && dDt.length < 7) { doanhThu = mt.tong; soDon = null; }
    var chien = {}, coMon = false;
    var cong = function (ds) { (ds || []).forEach(function (x) { if (sqKhoLaChien_(x.ten, chienDs)) { chien[x.ten] = (chien[x.ten] || 0) + so_(x.sl); } tenMon[x.ten] = 1; }); };
    if (mt && dMon.length < 7) { cong(mt.mon); coMon = true; } else if (dMon.length) { dMon.forEach(function (d) { cong(monNgay[d]); }); coMon = true; }
    var tongChien = 0; for (var c in chien) tongChien += chien[c];
    return { tu: m, den: days[6], days: days, kho: weekly[m] ? 'tuan' : dKho.length ? dKho.length : 0, khoNgay: dKho.length, dt: doanhThu, ngayDt: mt && dDt.length < 7 ? 7 : dDt.length, don: soDon,
      chien: coMon ? tongChien : null, chienDs: chien, ngayMon: mt && dMon.length < 7 ? 7 : dMon.length };
  });
  var cd = sqKhoCD_(), r2 = function (v) { return Math.round(v * 100) / 100; }, nhomDs = sqKhoNhom_(), nhomIds = sqKhoIds_(nhomDs);
  var dungTuan = function (k, wi) { var T = tuanInfo[wi], m = weekly[T.tu]; if (m) { var x = m[k]; return x ? r2(-(x[3] + x[5])) : 0; }
    if (!T.khoNgay) return null; var s = 0; T.days.forEach(function (d) { var x = daily[d] && daily[d][k]; if (x) s += -(x[3] + x[5]); }); return r2(s); };
  var cur = SO - 1, items = [];
  Object.keys(ten).forEach(function (k) {
    var a0 = ten[k], s = cd[k] || {}, coBan = false;
    var w = tuanInfo.map(function (T, wi) { return dungTuan(k, wi); });
    Object.keys(weekly).concat(Object.keys(daily)).forEach(function (d) { var x = (weekly[d] || daily[d])[k]; if (x && x[3] < 0) coBan = true; });
    var loai = s.loai || sqKhoLoaiDoan_(a0[0], coBan);
    if (loai === 'ban') return;
    if (!w.some(function (v) { return v > 0; })) return;
    var per = tuanInfo.map(function (T, wi) { return w[wi] != null && T.dt > 0 ? w[wi] / (T.dt / 1e7) : null; });
    var truoc = per.slice(cur - 4, cur).filter(function (v) { return v != null; });
    var base = truoc.length ? truoc.reduce(function (x, y) { return x + y; }, 0) / truoc.length : null;
    var wPrev = w.slice(cur - 4, cur).filter(function (v) { return v != null; }), avgW = wPrev.length ? wPrev.reduce(function (x, y) { return x + y; }, 0) / wPrev.length : null;
    var dung = w[cur], T = tuanInfo[cur], tile = null, exp = null;
    if (dung != null && base > 0 && per[cur] != null) { tile = per[cur] / base; exp = base * T.dt / 1e7; }
    else if (dung != null && avgW > 0) { tile = dung / avgW; exp = avgW; }
    var dau = KHO_DAU.indexOf(k) >= 0, ppc = null, ppcBase = null, ppcKy = null;
    if (dau) {
      var pc = tuanInfo.map(function (T2, wi) { return T2.chien > 0 && w[wi] > 0 ? T2.chien / w[wi] : null; });
      ppc = pc[cur]; var pb = pc.slice(cur - 4, cur).filter(function (v) { return v != null; }); ppcBase = pb.length ? pb.reduce(function (x, y) { return x + y; }, 0) / pb.length : null;
      // chưa có món chiên theo tuần → lấy mốc từ file Mặt hàng gộp nhiều tuần (cả kỳ) + dầu dùng cùng kỳ
      if (ppcBase == null) monGop.sort(function (a, b) { return a.den < b.den ? 1 : -1; }).some(function (g) {
        var ws = []; tuanInfo.forEach(function (T2, wi) { if (T2.tu >= g.tu && T2.den <= g.den) ws.push(wi); });
        if (ws.length * 7 !== sqKhoSoNgay_(g.tu, g.den) + 1 || ws.some(function (wi) { return w[wi] == null; })) return false;
        var dauKy = ws.reduce(function (s2, wi) { return s2 + w[wi]; }, 0), chienKy = 0;
        g.mon.forEach(function (m) { if (sqKhoLaChien_(m.ten, chienDs)) chienKy += so_(m.sl); });
        if (!(dauKy > 0) || !(chienKy > 0)) return false;
        ppcBase = chienKy / dauKy; ppcKy = { tu: g.tu, den: g.den, chien: chienKy, dau: r2(dauKy) }; return true;
      });
      if (ppc != null && ppcBase) { tile = ppcBase / ppc; exp = T.chien / ppcBase; }
    }
    var muc = '';
    // món dùng ít (< 1 đơn vị/tuần trung bình) thì không báo – số nhỏ lệch là chuyện thường
    if (tile != null && dung != null && avgW >= 1 && dung - exp >= 1) muc = tile >= 2 ? 'r' : tile >= 1.3 ? 'y' : '';
    var d7 = T.days.map(function (d) { var x = daily[d] && daily[d][k]; return x ? r2(-(x[3] + x[5])) : null; });
    items.push({ k: k, t: a0[0], dv: a0[1], loai: loai, nhom: sqKhoNhomCua_(a0[0], s, nhomIds), w: w, per: per.map(function (v) { return v == null ? null : Math.round(v * 1000) / 1000; }),
      dung: dung, truoc: w[cur - 1], per10: per[cur] == null ? null : Math.round(per[cur] * 1000) / 1000, base10: base == null ? null : Math.round(base * 1000) / 1000, avgW: avgW == null ? null : r2(avgW),
      tile: tile == null ? null : r2(tile), exp: exp == null ? null : r2(exp), muc: muc, d7: d7, dau: dau, ppc: ppc == null ? null : Math.round(ppc), ppcBase: ppcBase == null ? null : Math.round(ppcBase), ppcKy: ppcKy });
  });
  var thu = {}; nhomDs.forEach(function (g, i) { thu[g.id] = i; });
  items.sort(function (a, b) { return (thu[a.nhom] - thu[b.nhom]) || ((b.muc === 'r') - (a.muc === 'r')) || ((b.muc === 'y') - (a.muc === 'y')) || a.t.localeCompare(b.t); });
  var monList = Object.keys(tenMon).sort();
  return { tuan: tuan, tuNay: tuNay, homNay: hn, tuanInfo: tuanInfo, items: items, chienDs: chienDs, monList: monList, nhomDs: nhomDs };
}

/* ---------- ĐỊNH LƯỢNG & TỒN BẾP (đã xuất kho – chưa dùng) ----------
 * Định lượng (món bán → nguyên liệu kho) là BẢO MẬT: chỉ nằm trong tab SQ_DinhLuong của file Google Sheet quán,
 * anh Đạt tự tải file Excel lên trong Điều hành. Không có con số định lượng nào trong mã nguồn.
 * Còn ở bếp = hôm trước còn + xuất kho (cột Giảm trên Sapo) − món bán × định lượng − hủy hàng.
 */
var SQ_DL = 'SQ_DinhLuong';
var H_SQ_DL = ['NVL kho', 'Món bán', 'Lượng mỗi phần', 'Đơn vị', 'Bật', 'Nguồn', 'Sửa lúc'];
var BEP_BAT_DAU = '2026-11-10';   // ngày bắt đầu có file kho + mặt hàng từng ngày
var BEP_NHOM_CHINH = ['thit', 'ga', 'chan', 'hs', 'cha', 'trung'];
function sqTabDL_() { return sqTab_(SQ_DL, H_SQ_DL, function (sh) { sh.getRange('A:B').setNumberFormat('@'); sh.getRange('D:D').setNumberFormat('@'); sh.getRange('F:G').setNumberFormat('@'); }); }
function sqDLRows_() {
  var sh = sqTabDL_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, H_SQ_DL.length).getValues().map(function (r) {
    return { nvl: String(r[0] || ''), mon: String(r[1] || ''), sl: so_(r[2]), dv: String(r[3] || ''), bat: !(r[4] === 0 || r[4] === '0' || r[4] === false), src: String(r[5] || 'file'), luc: String(r[6] || '') };
  }).filter(function (x) { return x.nvl && x.mon; });
}
function sqDLGhi_(rows) {
  var sh = sqTabDL_(), n = sh.getLastRow();
  if (n >= 2) sh.getRange(2, 1, n - 1, H_SQ_DL.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, H_SQ_DL.length).setValues(rows.map(function (x) { return [x.nvl, x.mon, x.sl, x.dv, x.bat ? 1 : 0, x.src, x.luc]; }));
}
function sqDL_() { var m = {}; sqDLRows_().forEach(function (x) { (m[sqKhoK_(x.nvl)] = m[sqKhoK_(x.nvl)] || []).push(x); }); return m; }
// tải file định lượng: p.ds = [{nvl, quy, dvg, hao, mon, sl}] – dòng anh đã sửa tay (nguồn "anh") được giữ nguyên
function sqDLFile_(p) {
  var ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 1500).map(function (x) {
    return { nvl: chu_(x.nvl, 70), mon: chu_(x.mon, 80), sl: Math.round(so_(x.sl) * 10000) / 10000, dv: chu_(x.dvg, 12), quy: so_(x.quy), hao: so_(x.hao) };
  }).filter(function (x) { return x.nvl && x.mon && x.sl > 0; });
  if (!ds.length) throw new Error('File không có dòng định lượng nào (cần sheet "App").');
  return khoa_(function () {
    var luc = bayGio_(), cu = sqDLRows_(), giu = cu.filter(function (x) { return x.src === 'anh'; }), co = {}, out = giu.slice(), nMoi = 0;
    giu.forEach(function (x) { co[sqKhoK_(x.nvl) + '|' + khongDau_(x.mon)] = 1; });
    ds.forEach(function (x) { var k = sqKhoK_(x.nvl) + '|' + khongDau_(x.mon); if (co[k]) return; co[k] = 1; out.push({ nvl: x.nvl, mon: x.mon, sl: x.sl, dv: x.dv, bat: true, src: 'file', luc: luc }); nMoi++; });
    out.sort(function (a, b) { return a.nvl.localeCompare(b.nvl) || a.mon.localeCompare(b.mon); });
    sqDLGhi_(out); sqBepCPXoa_('');
    // quy đổi đơn vị kho: chỉ điền khi anh chưa đặt
    var sh = sqTabKhoCD_(), cd = sqKhoCD_(), xong = {}, nNvl = 0;
    ds.forEach(function (x) {
      var k = sqKhoK_(x.nvl); if (xong[k]) return; xong[k] = 1; nNvl++;
      var o = cd[k] || sqKhoCDMoi_(x.nvl), doi = false;
      if (o.quy === '' && x.quy > 0) { o.quy = x.quy; doi = true; }
      if (!o.dvg && x.dv) { o.dvg = x.dv; doi = true; }
      if (o.hao === '' && x.hao > 0 && x.hao < 0.95) { o.hao = x.hao; doi = true; }
      if (!doi) return;
      var row = cd[k] ? cd[k].row : sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, H_SQ_KHO_CD.length).setValues([sqKhoCDDong_(o.ten || x.nvl, o)]);
      if (!cd[k]) { o.row = row; cd[k] = o; }
    });
    return { ok: true, dong: nMoi, giu: giu.length, nvl: nNvl };
  });
}
// anh sửa món liên quan của 1 nguyên liệu: p = {ten, quy, dvg, hao, ds:[{mon, sl, bat}]}
function sqDLSua_(p) {
  var ten = chu_(p.ten, 70); if (!ten) throw new Error('Thiếu tên nguyên liệu.');
  var dvg = chu_(p.dvg, 12) || 'g', co = {};
  var ds = (Array.isArray(p.ds) ? p.ds : []).slice(0, 80).map(function (x) { return { mon: chu_(x.mon, 80), sl: Math.round(so_(x.sl) * 10000) / 10000, bat: !!x.bat }; })
    .filter(function (x) { var k = khongDau_(x.mon); if (!x.mon || co[k]) return false; co[k] = 1; return true; });
  if (ds.some(function (x) { return x.bat && !(x.sl > 0); })) throw new Error('Món đã tick cần có lượng mỗi phần.');
  var r = khoa_(function () {
    var luc = bayGio_(), k = sqKhoK_(ten), out = sqDLRows_().filter(function (x) { return sqKhoK_(x.nvl) !== k; });
    ds.forEach(function (x) { out.push({ nvl: ten, mon: x.mon, sl: x.sl, dv: dvg, bat: x.bat, src: 'anh', luc: luc }); });
    sqDLGhi_(out); sqBepCPXoa_('');
    return { ok: true };
  });
  sqKhoCDLuu_({ ten: ten, quy: p.quy, dvg: dvg, hao: p.hao });
  return r;
}
var SQ_KT = 'SQ_KiemTu';
var H_SQ_KT = ['ID', 'Cuối ngày', 'Lúc đếm', 'Nguyên liệu', 'App tính', 'Đếm thật', 'Lệch', 'Lý do', 'Người kiểm', 'Tạo lúc'];
var KT_LYDO = { huy: 'Hư / đổ', nv: 'Nhân viên dùng', khong: 'Không rõ' };
function sqTabKT_() { return sqTab_(SQ_KT, H_SQ_KT, function (sh) { sh.getRange('A:D').setNumberFormat('@'); sh.getRange('H:J').setNumberFormat('@'); }); }
function sqKTRows_() {
  var sh = sqTabKT_(), n = sh.getLastRow();
  if (n < 2) return [];
  var sr = function (v) { return v === '' || v === null ? '' : Number(v); };
  return sh.getRange(2, 1, n - 1, H_SQ_KT.length).getValues().map(function (r, i) {
    return { row: i + 2, id: String(r[0] || ''), ngay: chuanNgay_(r[1]), luc: String(r[2] || ''), nvl: String(r[3] || ''), du: sr(r[4]), dem: sr(r[5]), lech: sr(r[6]), ly: String(r[7] || ''), nguoi: String(r[8] || ''), tao: String(r[9] || '') };
  }).filter(function (x) { return x.id && x.nvl && /^\d{4}-\d{2}-\d{2}$/.test(x.ngay) && x.dem !== ''; });
}
function sqBepCPXoa_(tu) { var cp = sqBepCP_(); if (cp && (!tu || cp.ngay >= tu)) PropertiesService.getScriptProperties().deleteProperty('KHO_BEP_CP'); }
// kiểm tủ: p = {ten, ngay (cuối ngày được tính), luc: 'sang'|'toi', dem, lyDo}
function sqBepKiem_(p, nguoi) {
  var ten = chu_(p.ten, 70); if (!ten) throw new Error('Thiếu tên nguyên liệu.');
  var ngay = sqKiemNgay_(p.ngay), hn = homNay_();
  if (ngay > hn || ngay < congNgay_(hn, -7)) throw new Error('Ngày kiểm tủ không hợp lệ.');
  var dem = Number(String(p.dem === undefined ? '' : p.dem).replace(',', '.'));
  if (!isFinite(dem) || dem < 0 || String(p.dem).trim() === '') throw new Error('Điền số thực còn trong tủ.');
  dem = Math.round(dem * 1000) / 1000;
  var ly = KT_LYDO[String(p.lyDo || '')] ? String(p.lyDo) : '', luc = p.luc === 'toi' ? 'toi' : 'sang', k = sqKhoK_(ten);
  // app tính (trước khi lấy số đếm) – nếu đã có file kho + mặt hàng của ngày đó
  var du = '';
  try { var b = sqKhoBep_({ ngay: ngay, boQua: k + '|' + ngay }), it = b.items.filter(function (x) { return x.k === k; })[0];
    if (it && b.ngay === ngay && it.bep != null) du = it.bep; } catch (e) { }
  return khoa_(function () {
    var sh = sqTabKT_(), cu = sqKTRows_().filter(function (x) { return x.ngay === ngay && sqKhoK_(x.nvl) === k; })[0];
    var row = cu ? cu.row : sh.getLastRow() + 1, id = cu ? cu.id : 'KT' + Date.now().toString(36);
    var lech = du === '' ? '' : Math.round((dem - du) * 100) / 100;
    sh.getRange(row, 1, 1, H_SQ_KT.length).setValues([[id, ngay, luc, ten, du, dem, lech, ly, chu_(nguoi, 40), bayGio_()]]);
    sqBepCPXoa_(ngay);
    return { ok: true, du: du, dem: dem, lech: lech, cho: du === '' };
  });
}
function sqBepDat_(p) { return sqKhoCDLuu_({ ten: p.ten, bepSl: p.sl, bepNgay: p.sl === '' || p.sl === null || p.sl === undefined ? '' : p.ngay }); }
function sqBepCP_() { var v = PropertiesService.getScriptProperties().getProperty('KHO_BEP_CP'), o = v ? sqJ_(v, null) : null; return o && /^\d{4}-\d{2}-\d{2}$/.test(String(o.ngay || '')) && o.v ? o : null; }
function sqMonGoc_(k) { return k.replace(/\s*\([^()]*\)\s*$/, ''); }

function sqKhoBep_(p) {
  p = p || {};
  var meta = sqKhoMeta_().filter(function (x) { return x.loai === 'kho' && x.ngay >= BEP_BAT_DAU; }).sort(function (a, b) { return a.ngay < b.ngay ? -1 : 1; });
  var dl = sqDL_(), coDL = Object.keys(dl).length > 0, nhomDs = sqKhoNhom_(), nhomIds = sqKhoIds_(nhomDs), hn = homNay_();
  if (!meta.length) return { trong: true, coDL: coDL, homNay: hn, items: [], monList: [], nhomDs: nhomDs };
  var ngays = meta.map(function (x) { return x.ngay; }), ngay = /^\d{4}-\d{2}-\d{2}$/.test(String(p.ngay || '')) ? String(p.ngay) : ngays[ngays.length - 1];
  var tr = ngays.filter(function (d) { return d <= ngay; }); ngay = tr.length ? tr[tr.length - 1] : ngays[0];
  var laMoi = ngay === ngays[ngays.length - 1], cp = sqBepCP_(), tu = BEP_BAT_DAU, mo = {};
  if (cp && cp.ngay >= BEP_BAT_DAU && cp.ngay < ngay) { tu = congNgay_(cp.ngay, 1); mo = cp.v; }
  else if (sqKhoSoNgay_(tu, ngay) > 40) tu = congNgay_(ngay, -40);
  var can = meta.filter(function (x) { return x.ngay >= tu && x.ngay <= ngay; }), snap = {};
  if (can.length) {
    var r0 = Math.min.apply(null, can.map(function (x) { return x.row; })), r1 = Math.max.apply(null, can.map(function (x) { return x.row; }));
    var cot = sqTabKho_().getRange(r0, 5, r1 - r0 + 1, 1).getValues();
    can.forEach(function (x) { var m = {}; sqKhoMon_(sqJ_(String(cot[x.row - r0][0] || ''), [])).forEach(function (a) { m[sqKhoK_(a[0])] = a; }); snap[x.ngay] = m; });
  }
  // món bán từng ngày
  var sold = {}, goc = {}, tenMon = {}, tu30 = congNgay_(ngay, -30);
  sqSapoRows_().forEach(function (x) {
    if (x.loai !== 'mon' || x.ngay > ngay || x.ngay < (tu < tu30 ? tu : tu30)) return;
    var s = {}, g = {};
    (x.d.mon || []).forEach(function (m) { var k = khongDau_(m.ten), b = sqMonGoc_(k); s[k] = (s[k] || 0) + so_(m.sl); (g[b] = g[b] || { sl: 0, n: 0, ten: m.ten }); if (k !== b) g[b].n++; else g[b].sl += so_(m.sl); tenMon[m.ten] = 1; });
    if (x.ngay >= tu) { sold[x.ngay] = s; goc[x.ngay] = g; }
  });
  // hủy hàng từng ngày
  var huyNgay = {};
  try { sqHuyRows_().forEach(function (h) { if (h.ngay >= tu && h.ngay <= ngay) (huyNgay[h.ngay] = huyNgay[h.ngay] || []).push({ tok: sqKhoTok_(h.ten), sl: h.sl, dv: khongDau_(h.dv) }); }); } catch (e) { }
  var days = sqKhoNgayDs_(ngay, sqKhoSoNgay_(tu, ngay) + 1), cd = sqKhoCD_(), cur = snap[ngay] || {};
  var ktAll = sqKTRows_().filter(function (x) { return sqKhoK_(x.nvl) + '|' + x.ngay !== p.boQua; }), ktMap = {}, ktDs = {};
  ktAll.forEach(function (x) { var kk = sqKhoK_(x.nvl); ktMap[kk + '|' + x.ngay] = x; (ktDs[kk] = ktDs[kk] || []).push(x); });
  var ngMoi = ngays[ngays.length - 1];
  var r2 = function (v) { return Math.round(v * 100) / 100; }, soV = function (v) { return String(r2(v)).replace('.', ','); };
  var items = [], cpMoi = {}, cpNgay = days.length > 24 && laMoi ? congNgay_(ngay, -14) : '';
  Object.keys(cur).forEach(function (k) {
    var a0 = cur[k], s = cd[k] || {}, tatCa = dl[k] || [], links = tatCa.filter(function (x) { return x.bat && x.sl > 0; });
    var nhom = sqKhoNhomCua_(a0[0], s, nhomIds), tuTru = false, coGiam = false;
    days.forEach(function (d) { var x = snap[d] && snap[d][k]; if (x) { if (x[3] < 0) tuTru = true; if (x[5] < 0) coGiam = true; } });
    if (!tatCa.length && !(BEP_NHOM_CHINH.indexOf(nhom) >= 0 && coGiam && !tuTru)) return;
    var quy = s.quy > 0 ? s.quy : null, hao = s.hao > 0 ? s.hao : 0, dvg = s.dvg || (tatCa[0] && tatCa[0].dv) || 'g', dvK = khongDau_(a0[1]), tok = sqKhoTok_(a0[0]);
    var bep = mo[k] !== undefined ? so_(mo[k]) : 0, ser = [];
    days.forEach(function (d) {
      var x = snap[d] && snap[d][k], xuat = x ? r2(-x[5]) : null, ban = x ? r2(-x[3]) : null, dung = null, ds = [], sd = sold[d];
      if (sd && links.length && quy) {
        var tong = 0, daGop = {};
        links.forEach(function (l) {
          var lk = khongDau_(l.mon), sl = sd[lk];
          if (sl === undefined) { // file Mặt hàng cũ gộp biến thể (mẹt nhỏ + lớn) → ước theo định lượng trung bình
            var b = sqMonGoc_(lk), g = goc[d][b];
            if (b === lk || !g || g.n || !(g.sl > 0) || daGop[b]) return;
            daGop[b] = 1; var cung = links.filter(function (y) { return sqMonGoc_(khongDau_(y.mon)) === b; }), tb = cung.reduce(function (t, y) { return t + y.sl; }, 0) / cung.length;
            tong += g.sl * tb; ds.push({ t: g.ten + ' (gộp các cỡ)', sl: r2(g.sl), dl: r2(tb), g: r2(g.sl * tb), uoc: true }); return;
          }
          if (!(sl > 0)) return;
          tong += sl * l.sl; ds.push({ t: l.mon, sl: r2(sl), dl: l.sl, g: r2(sl * l.sl) });
        });
        dung = r2(tong / quy / (1 - hao));
      }
      var huy = 0; (huyNgay[d] || []).forEach(function (h) { if (sqKhoDiem_(tok, h.tok) < 2 && sqKhoDiem_(h.tok, tok) < 2) return;
        if (h.dv === dvK) huy += h.sl; else if (quy && (h.dv === 'g' || h.dv === 'gam' || h.dv === 'gr') && khongDau_(dvg) === 'g') huy += h.sl / quy; else if (quy && h.dv === 'kg' && khongDau_(dvg) === 'g') huy += h.sl * 1000 / quy; });
      huy = r2(huy);
      var prev = r2(bep), dat = false, kiem = null, kt = ktMap[k + '|' + d];
      if (!tuTru) { if (xuat != null) bep += xuat; if (dung != null) bep -= dung; bep -= huy;
        if (kt) { var du0 = r2(bep); kt.du = du0; kt.lech = r2(kt.dem - du0); kiem = { dem: kt.dem, du: du0, lech: kt.lech, ly: kt.ly, nguoi: kt.nguoi, luc: kt.luc }; bep = kt.dem; dat = true; }
        else if (s.bepNgay === d && s.bepSl !== '' && s.bepSl !== undefined) { bep = s.bepSl; dat = true; } }
      ser.push({ d: d, prev: prev, xuat: xuat, ban: ban, dung: dung, huy: huy, bep: r2(bep), dat: dat, kiem: kiem, thieu: !x ? 'kho' : !sd ? 'mon' : '', ds: ds });
      if (d === cpNgay) cpMoi[k] = r2(bep);
    });
    var coDung = ser.filter(function (z) { return z.dung != null; }).slice(-7), avg = coDung.length ? coDung.reduce(function (t, z) { return t + z.dung; }, 0) / coDung.length : null;
    var dv = String(a0[1]).toLowerCase();
    // đánh giá từng ngày
    ser.forEach(function (z) {
      var muc = 'm', nhan = '';
      if (!links.length) nhan = 'cần tick món';
      else if (!quy) nhan = 'cần quy đổi';
      else if (z.thieu === 'kho') nhan = 'chưa có file kho';
      else if (z.thieu === 'mon') nhan = 'chưa có file Mặt hàng';
      else if (tuTru) { var le = r2(z.ban - z.dung); muc = Math.abs(le) < Math.max(0.5, 0.1 * z.dung) ? 'g' : 'y'; nhan = muc === 'g' ? 'khớp Sapo' : 'lệch ' + soV(Math.abs(le)); }
      else if (z.kiem) { var nguong = Math.max(0.25, 0.1 * (avg || 0)), lc = z.kiem.lech;
        if (lc <= -nguong) { muc = z.kiem.ly && z.kiem.ly !== 'khong' ? 'y' : 'r'; nhan = 'hụt ' + soV(-lc) + ' ' + dv; }
        else if (lc >= nguong) { muc = 'y'; nhan = 'thừa ' + soV(lc) + ' ' + dv; }
        else { muc = 'g'; nhan = 'kiểm tủ khớp'; } }
      else if (z.dat) { muc = 'g'; nhan = 'anh đặt lại'; }
      else if (z.bep < -Math.max(0.1, 0.1 * (avg || 0))) { muc = 'r'; nhan = 'âm ' + soV(-z.bep) + ' ' + dv; }
      else if (avg > 0 && z.bep > 0.5 && z.bep > 1.5 * avg) { muc = 'y'; nhan = 'tủ đủ ~' + soV(Math.round(z.bep / avg * 10) / 10) + ' ngày'; }
      else { muc = 'g'; nhan = Math.abs(z.bep) < 0.05 ? 'khớp' : 'hợp lý'; }
      z.muc = muc; z.nhan = nhan;
    });
    var L = ser[ser.length - 1], gan = ser.slice(-7), mucL = L.muc, nhanL = L.nhan;
    var hut = ser.slice(-3).filter(function (z) { return z.kiem && z.muc !== 'g' && /^hụt/.test(z.nhan); }).pop();
    if (hut && mucL !== 'r' && !(L.kiem)) { mucL = hut.muc; nhanL = hut.nhan + ' (kiểm ' + sqDM_(hut.d) + ')'; }
    var kiemDs = (ktDs[k] || []).slice().sort(function (a, b) { return a.ngay < b.ngay ? 1 : -1; }).slice(0, 8).map(function (x) {
      return { ngay: x.ngay, luc: x.luc, du: x.du, dem: x.dem, lech: x.lech, ly: x.ly, nguoi: x.nguoi, cho: x.ngay > ngMoi }; });
    items.push({ k: k, t: a0[0], dv: a0[1], nhom: nhom, quy: quy, dvg: dvg, hao: hao, tuTru: tuTru, cuoi: a0[7],
      links: tatCa.map(function (x) { return { mon: x.mon, sl: x.sl, bat: x.bat, src: x.src }; }),
      prev: L.prev, xuat: L.xuat, ban: L.ban, dung: L.dung, huy: L.huy, bep: tuTru ? null : L.bep, muc: mucL, nhan: nhanL, kiem: L.kiem, kiemDs: kiemDs,
      avg: avg == null ? null : r2(avg), duNgay: !tuTru && avg > 0 && L.bep > 0 ? Math.round(L.bep / avg * 10) / 10 : null,
      tongXuat: r2(gan.reduce(function (t, z) { return t + (z.xuat || 0); }, 0)), tongDung: r2(gan.reduce(function (t, z) { return t + (z.dung || 0); }, 0)),
      dat: s.bepNgay && s.bepSl !== '' ? { ngay: s.bepNgay, sl: s.bepSl } : null, days: gan.slice().reverse() });
  });
  if (cpNgay && !p.boQua && Object.keys(cpMoi).length) { try { PropertiesService.getScriptProperties().setProperty('KHO_BEP_CP', JSON.stringify({ ngay: cpNgay, v: cpMoi })); } catch (e) { } }
  var thu = {}, tm = { r: 0, y: 1, g: 2, m: 3 }; nhomDs.forEach(function (g, i) { thu[g.id] = i; });
  items.sort(function (a, b) { return ((a.tuTru ? 1 : 0) - (b.tuTru ? 1 : 0)) || (tm[a.muc] - tm[b.muc]) || (thu[a.nhom] - thu[b.nhom]) || a.t.localeCompare(b.t); });
  sqDLRows_().forEach(function (x) { tenMon[x.mon] = 1; });
  return { ngay: ngay, ngayMoi: ngMoi, tu: tu, homNay: hn, coDL: coDL, items: items, monList: Object.keys(tenMon).sort(function (a, b) { return a.localeCompare(b); }), nhomDs: nhomDs };
}


/* ---------- SỔ QUÁN: quản lý xem kho (chức năng "kho", anh Đạt bật trong Điều hành → Nhân viên) ---------- */
function sqKhoQuyen_(nv) { if ((nv.chucNang || []).indexOf('kho') < 0) throw new Error('Em chưa được mở chức năng Xem kho. Nhờ anh Đạt bật trong Điều hành → Nhân viên.'); }
function sqKhoNV_(nv, p) { sqKhoQuyen_(nv); var j = sqKhoTinh_(p || {}); delete j.doiChieu; return j; }
function sqKhoBepNV_(nv, p) { sqKhoQuyen_(nv); return sqKhoBep_({ ngay: p && p.ngay }); }
function sqKhoTuanNV_(nv, p) {
  sqKhoQuyen_(nv);
  var j = sqKhoTuan_(p || {}), I = j.tuanInfo;
  j.tuanInfo = I.map(function (T, i) { var P = I[i - 1];
    return { tu: T.tu, den: T.den, days: T.days, kho: T.kho, khoNgay: T.khoNgay, ngayDt: T.ngayDt, ngayMon: T.ngayMon, chien: T.chien, chienDs: T.chienDs,
      dtPct: T.dt > 0 && P && P.dt > 0 ? Math.round((T.dt / P.dt - 1) * 100) : null }; });
  j.items.forEach(function (x) { delete x.per; delete x.per10; delete x.base10; delete x.exp; });
  return j;
}
function sqKhoDatNV_(nv, p) { sqKhoQuyen_(nv); return sqKhoCDLuu_({ ten: p.ten, dat: p.dat }); }
function sqBepKiemNV_(nv, p) { sqKhoQuyen_(nv); return sqBepKiem_(p, nv.ten); }
