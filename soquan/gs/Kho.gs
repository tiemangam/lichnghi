/**
 * KHO – số tồn kho Sapo (file "Báo cáo kho hàng") Sơn gửi mỗi tối + file nhiều ngày (tháng) để tính trung bình.
 * Điều hành → tab Kho: cần đặt (đã hết / sắp hết theo mức tối thiểu), xuất dùng so với định mức (báo dùng đột biến),
 * nhập – xuất trong ngày, đối chiếu "Thêm" với hóa đơn mua và Hủy hàng với "Giảm".
 * Anh chỉnh từng món ở tab SQ_KhoCaiDat (qua nút ✎): loại, mức tối thiểu, định mức dùng (vd 1 can trong 4–5 ngày).
 */
var DH_KHO = 'DH_Kho';
var H_DH_KHO = ['Ngày', 'Loại', 'Dữ liệu tới ngày', 'Số ngày', 'Số liệu (JSON)', 'Tải lúc'];
var SQ_KHO_CD = 'SQ_KhoCaiDat';
var H_SQ_KHO_CD = ['Món', 'Loại', 'Mức tối thiểu', 'Định mức SL', 'Trong (ngày) từ', 'đến', 'Đã đặt ngày', 'Sửa lúc', 'Nhóm'];
var KHO_LOAI = { tuoi: 'Đồ tươi', kho: 'Gia vị – xuất dùng', bb: 'Bao bì – vật dụng', ban: 'Bán sẵn', tp: 'Thành phẩm bếp' };
var KHO_THEO_DOI = ['tuoi', 'kho', 'bb'];   // loại được theo dõi tốc độ dùng
var KHO_CUA_SO = 3;                          // xét 3 ngày gần nhất
var KHO_SO_NGAY = 14;

function sqTabKho_() { return sqTab_(DH_KHO, H_DH_KHO, function (sh) { sh.getRange('A:D').setNumberFormat('@'); sh.getRange('F:F').setNumberFormat('@'); }); }
function sqTabKhoCD_() {
  var sh = sqTab_(SQ_KHO_CD, H_SQ_KHO_CD, function (sh) { sh.getRange('A:B').setNumberFormat('@'); sh.getRange('G:I').setNumberFormat('@'); });
  if (!sh.getRange(1, 9).getValue()) sh.getRange(1, 9).setValue('Nhóm').setFontWeight('bold').setBackground('#445803').setFontColor('#F8E8C4');
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
    m[sqKhoK_(t)] = { row: i + 2, ten: t, loai: KHO_LOAI[String(r[1])] ? String(r[1]) : '', min: sr(r[2]), sl: sr(r[3]), n1: sr(r[4]), n2: sr(r[5]), dat: chuanNgay_(r[6]) || '', nhom: String(r[8] || '') };
  });
  return m;
}
function sqKhoCDLuu_(p) {
  var ten = chu_(p.ten, 70); if (!ten) throw new Error('Thiếu tên món.');
  var so0 = function (v, am) { if (v === '' || v === null || v === undefined) return ''; var x = Number(v); if (!isFinite(x) || x < 0 || (!am && x === 0)) return ''; return Math.round(x * 1000) / 1000; };
  return khoa_(function () {
    var sh = sqTabKhoCD_(), cd = sqKhoCD_(), x = cd[sqKhoK_(ten)];
    var o = x || { ten: ten, loai: '', min: '', sl: '', n1: '', n2: '', dat: '', nhom: '' };
    if ('loai' in p) o.loai = KHO_LOAI[String(p.loai)] ? String(p.loai) : '';
    if ('min' in p) o.min = so0(p.min, true);
    if ('sl' in p) o.sl = so0(p.sl);
    if ('n1' in p) o.n1 = so0(p.n1);
    if ('n2' in p) o.n2 = so0(p.n2);
    if ('nhom' in p) o.nhom = /^[a-z0-9_]{1,20}$/.test(String(p.nhom || '')) ? String(p.nhom) : '';
    if ('dat' in p) o.dat = /^\d{4}-\d{2}-\d{2}$/.test(String(p.dat || '')) ? String(p.dat) : '';
    if (o.sl === '' || o.n1 === '') { o.sl = ''; o.n1 = ''; o.n2 = ''; }
    if (o.n2 !== '' && o.n2 < o.n1) { var t = o.n1; o.n1 = o.n2; o.n2 = t; }
    var row = x ? x.row : sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, H_SQ_KHO_CD.length).setValues([[ten, o.loai, o.min, o.sl, o.n1, o.n2, o.dat, bayGio_(), o.nhom || '']]);
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
      nhom: sqKhoNhomCua_(a[0], s, nhomIds), nhomDoan: sqKhoNhomDoan_(a[0]), d14: d14, n14: n14, soNgay3: w.length, dung3: dung3, exp3: exp3, tile: tile, muc: muc, duNgay: rate > 0 ? Math.round(cuoi / rate * 10) / 10 : null, dat: !!(s.dat && s.dat >= ngay) };
  });

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
    out.push({ muc: 'r', loai: 'Kho – dùng nhiều', tab: 'kho', ngay: j.ngay, text: x.t + ': dùng ' + String(x.dung3).replace('.', ',') + ' ' + String(x.dv).toLowerCase() + ' trong ' + x.soNgay3 + ' ngày – gấp ' + String(Math.round(x.tile * 10) / 10).replace('.', ',') + ' lần định mức' });
  });
  try {
    var t = sqKhoTuan_({});
    t.items.filter(function (x) { return x.muc === 'r'; }).slice(0, 4).forEach(function (x) {
      var T = t.tuanInfo[t.tuanInfo.length - 1];
      out.push({ muc: 'r', loai: 'Kho – tuần ' + sqDM_(T.tu) + '–' + sqDM_(T.den), tab: 'kho', ngay: T.den, text: x.t + ': dùng ' + String(x.dung).replace('.', ',') + ' ' + String(x.dv).toLowerCase() + ' – gấp ' + String(Math.round(x.tile * 10) / 10).replace('.', ',') + ' lần so với ' + (x.dau && x.ppcBase ? 'số món chiên bán' : 'doanh thu') + ' (4 tuần trước)' });
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
