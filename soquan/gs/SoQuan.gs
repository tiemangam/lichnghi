/**
 * SỔ QUÁN – TIỆM ĂN GẶM  (Đợt 1: chụp hóa đơn, hàng tự đặt, bill không tính tiền, duyệt phiếu, giá nhập thay đổi)
 *
 * File này THÊM VÀO dự án Apps Script "Phiếu chi đi chợ Gặm" (tài khoản quán), cùng file dữ liệu với app đi chợ.
 * Dùng lại các hàm có sẵn trong Code.gs: layFile_, apiLuu, docDanhMuc_, khongDau_, chu_, so_, soHoacRong_, chuanNgay_,
 * homNay_, bayGio_, khoa_, timDong_, congNgay_.
 * Phiếu hóa đơn ảnh + hàng tự đặt được ghi vào Phieu/ChiTiet như phiếu đi chợ → P&L tự kéo về như cũ.
 * Bill không tính tiền KHÔNG phải chi phí → chỉ ghi ở tab SQ_Phieu.
 *
 * Thuộc tính tập lệnh (Cài đặt dự án → Thuộc tính tập lệnh):
 *   GEMINI_KEY   – mã Gemini (aistudio.google.com)            [bắt buộc để đọc ảnh]
 *   GEMINI_MODEL – tên model, mặc định gemini-3.6-flash       [không bắt buộc]
 *   CHU_PASS     – mật khẩu của anh Đạt (mở phần Điều hành)  [bắt buộc]
 *   QL_KEY       – mã quản lý app Nhà Gặm (để app quản lý vào được phần cũ khi anh Đạt đăng nhập)
 *   SQ_FOLDER_ID – tự tạo khi chạy sqCaiDat
 */

var SQ_NV = 'SQ_NhanVien', SQ_META = 'SQ_Phieu', SQ_ALIAS = 'SQ_TenGoiKhac';
var H_SQ_NV = ['Tên', 'Mã PIN', 'Vai trò', 'Nghỉ làm (gõ x)', 'Chức năng', 'Sửa lúc'];
var H_SQ_META = ['ID phiếu', 'Loại', 'Ngày', 'Người nhập', 'Mối', 'Ảnh (file ID)', 'AI đọc (JSON)', 'Các dòng (JSON)',
  'Tổng trên ảnh', 'Cộng các dòng', 'Ghi chú NV', 'Trạng thái', 'Câu hỏi của anh', 'Trả lời NV', 'Duyệt lúc', 'Tạo lúc', 'Thông tin thêm (JSON)', 'Tổng'];
var M = { ID: 0, LOAI: 1, NGAY: 2, NGUOI: 3, MOI: 4, ANH: 5, AI: 6, DONG: 7, TONGANH: 8, CONG: 9, GHICHU: 10, TT: 11, HOI: 12, TRALOI: 13, DUYETLUC: 14, TAOLUC: 15, THEM: 16, TONG: 17 };
var H_SQ_ALIAS = ['Mối', 'Chữ trên hóa đơn', 'Tên món đúng', 'Người dạy', 'Lúc'];
var LOAI = { HD: 'Hóa đơn ảnh', TD: 'Hàng tự đặt', BILL: 'Bill không tính tiền', DC: 'Đi chợ' };
var TT = { CHO: 'Chờ duyệt', OK: 'Đã duyệt', HOI: 'Hỏi lại', TRALOI: 'Đã trả lời', TUCHOI: 'Từ chối', DOC: 'Đang đọc', KIEM: 'Chờ kiểm tra', LOI: 'Lỗi đọc' };
var TT_CHUA_XAC_NHAN = ['Đang đọc', 'Chờ kiểm tra', 'Lỗi đọc'];
var NHOM_TU_DAT = 'Hàng tự đặt';

/* ======================= CÀI ĐẶT (chạy 1 lần) ======================= */
function sqCaiDat() {
  var ss = layFile_();
  var mk = function (ten, head) {
    var sh = ss.getSheetByName(ten);
    if (sh) return sh;
    sh = ss.insertSheet(ten);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#445803').setFontColor('#F8E8C4');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@');
    return sh;
  };
  var nv = mk(SQ_NV, H_SQ_NV);
  if (nv.getRange(1, 5).getValue() !== 'Chức năng') nv.getRange(1, 1, 1, H_SQ_NV.length).setValues([H_SQ_NV]);
  mk(SQ_META, H_SQ_META);
  mk(SQ_ALIAS, H_SQ_ALIAS);
  sqHdCaiDat_(mk);
  sqHuyCaiDat_(mk);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SQ_FOLDER_ID')) {
    var f = DriveApp.createFolder('Sổ Quán – ảnh hóa đơn');
    props.setProperty('SQ_FOLDER_ID', f.getId());
  }
  var thieu = ['GEMINI_KEY', 'QL_KEY'].filter(function (k) { return !props.getProperty(k); });
  var msg = 'Sổ Quán đã sẵn sàng. Tab nhân viên: ' + SQ_NV + '. Thư mục ảnh: https://drive.google.com/drive/folders/' + props.getProperty('SQ_FOLDER_ID') +
    (thieu.length ? '\nCÒN THIẾU thuộc tính: ' + thieu.join(', ') : '\nĐủ thuộc tính.');
  Logger.log(msg);
  return msg;
}

/* ======================= CỔNG API ======================= */
function sqGet_(e) { return sqOut_(sqXuLy_(e.parameter.api, e.parameter)); }
function doPost(e) {
  var b = {};
  try { b = JSON.parse(e.postData.contents || '{}'); } catch (err) { return sqOut_({ ok: false, error: 'Dữ liệu gửi lên bị lỗi.' }); }
  return sqOut_(sqXuLy_(b.api, b));
}
function sqOut_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sqXuLy_(api, p) {
  try {
    var f = SQ_API[api];
    if (!f) throw new Error('Không có lệnh ' + api);
    var r = f(p || {});
    r.ok = true;
    return r;
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

var SQ_API = {
  /* ---- nhân viên ---- */
  nv_vao: function (p) { var nv = sqNV_(p.pin); return { ten: nv.ten, vai: nv.vai, chucNang: nv.chucNang, homNay: homNay_() }; },
  nv_dulieu: function (p) { var nv = sqNV_(p.pin); return sqDuLieuNV_(nv); },
  nv_doc: function (p) { sqNV_(p.pin); return sqDocAnh_(p); },
  nv_luu_hd: function (p) { var nv = sqNV_(p.pin); return sqLuuHD_(nv, p); },
  nv_luu_tudat: function (p) { var nv = sqNV_(p.pin); return sqLuuTuDat_(nv, p); },
  nv_luu_bill: function (p) { var nv = sqNV_(p.pin); return sqLuuBill_(nv, p); },
  nv_tra_loi: function (p) { var nv = sqNV_(p.pin); return sqTraLoi_(nv, p); },
  nv_huy_luu: function (p) { var nv = sqNV_(p.pin); return sqHuyLuu_(nv, p); },
  nv_gui_hd: function (p) { var nv = sqNV_(p.pin); return sqGuiHD_(nv, p); },
  nv_doc_nen: function (p) { sqNV_(p.pin); return sqDocNen_(p.ids || []); },
  nv_doc_lai: function (p) { sqNV_(p.pin); return sqDocLai_(p.id); },
  nv_hd_ds: function (p) { var nv = sqNV_(p.pin); return sqHdDs_(nv, p); },
  nv_hd_xem: function (p) { sqNV_(p.pin); var m = sqTimMeta_(sqKiemId_(p.id)); if (!m) throw new Error('Không tìm thấy hóa đơn.'); return { hd: m }; },
  nv_hd_xoa: function (p) { var nv = sqNV_(p.pin); return sqHdXoa_(nv, p.id); },
  nv_anh: function (p) { sqNV_(p.pin); sqKiemAnhCuaPhieu_(p.fid); return sqAnh_(p.fid); },
  nv_quen_luu: function (p) { var nv = sqNV_(p.pin); return sqQuenLuu_(nv, p); },
  nv_quen_xoa: function (p) { sqNV_(p.pin); return sqQuenXoa_(p.id); },
  nv_tn_ngay: function (p) { var nv = sqNV_(p.pin); return sqTnNgay_(nv, p); },
  nv_tn_luu: function (p) { var nv = sqNV_(p.pin); return sqTnLuu_(nv, p); },
  nv_tn_thang: function (p) { var nv = sqNV_(p.pin); return sqTnThang_(nv, p); },
  nv_sapo_ds: function (p) { var nv = sqNV_(p.pin); return sqSapoNV_(nv); },
  nv_sapo_luu: function (p) { var nv = sqNV_(p.pin); sqSapoQuyen_(nv); var r = sqSapoLuu_(p, nv.ten); r.ds = sqSapoNV_(nv).ds; return r; },
  nv_phieu: function (p) { var nv = sqNV_(p.pin); var x = sqChiTietPhieu_(p.id); if (x.meta && x.meta.nguoi !== nv.ten) throw new Error('Phiếu này không phải của em.'); return { phieu: x }; },
  /* ---- anh Đạt ---- */
  chu_vao: function (p) { sqChu_(p.pass); return { qlKey: PropertiesService.getScriptProperties().getProperty('QL_KEY') || '' }; },
  chu_trang_thai: function () { return { coMatKhau: !!PropertiesService.getScriptProperties().getProperty('CHU_PASS') }; },
  chu_tao_pass: function (p) { return sqTaoPass_(p); },
  chu_doi_pass: function (p) { sqChu_(p.pass); return sqDoiPass_(p); },
  chu_nv_ds: function (p) { sqChu_(p.pass); return sqNVDs_(); },
  chu_nv_luu: function (p) { sqChu_(p.pass); return sqNVLuu_(p); },
  chu_duyet_ds: function (p) { sqChu_(p.pass); return sqDuyetDs_(p); },
  chu_duyet: function (p) { sqChu_(p.pass); return sqDuyet_(p); },
  chu_phieu: function (p) { sqChu_(p.pass); return { phieu: sqChiTietPhieu_(p.id) }; },
  chu_anh: function (p) { sqChu_(p.pass); return sqAnh_(p.fid); },
  chu_gia: function (p) { sqChu_(p.pass); return sqGiaThayDoi_(p); },
  chu_mon: function (p) { sqChu_(p.pass); return sqLichSuMon_(p.khoa); },
  chu_huy: function (p) { sqChu_(p.pass); return sqHuyBaoCao_(p); },
  chu_huy_gia: function (p) { sqChu_(p.pass); return sqHuyDatGia_(p); },
  chu_tn: function (p) { sqChu_(p.pass); return sqTnChu_(p); },
  chu_tn_xuly: function (p) { sqChu_(p.pass); return sqTnXuLy_(p); },
  chu_tong: function (p) { sqChu_(p.pass); return sqTongQuan_(p); },
  chu_cp: function (p) { sqChu_(p.pass); return sqCpNgay_(p); },
  chu_cp_moi: function (p) { sqChu_(p.pass); return sqCpMoi_(p); },
  chu_cp_sua: function (p) { sqChu_(p.pass); return sqCpSua_(p); },
  chu_cp_quay: function (p) { sqChu_(p.pass); return sqCpQuay_(p); },
  chu_sapo_luu: function (p) { sqChu_(p.pass); return sqSapoLuu_(p, 'anh Đạt'); },
  chu_kho: function (p) { sqChu_(p.pass); return sqKhoChu_(p); },
  chu_sapo_ds: function (p) { sqChu_(p.pass); return sqSapoDs_(14); },
  chu_kho_cd: function (p) { sqChu_(p.pass); return sqKhoCDLuu_(p); },
  chu_kho_tuan: function (p) { sqChu_(p.pass); return sqKhoTuan_(p); },
  chu_kho_chien: function (p) { sqChu_(p.pass); return sqKhoChienLuu_(p); },
  chu_kho_nhom: function (p) { sqChu_(p.pass); return sqKhoNhomLuu_(p); },
  chu_kho_bep: function (p) { sqChu_(p.pass); return sqKhoBep_(p); },
  chu_dl_file: function (p) { sqChu_(p.pass); return sqDLFile_(p); },
  chu_dl_sua: function (p) { sqChu_(p.pass); return sqDLSua_(p); },
  chu_bep_dat: function (p) { sqChu_(p.pass); return sqBepDat_(p); }
};

/* ======================= ĐĂNG NHẬP ======================= */
function sqNV_(pin) {
  pin = String(pin || '').trim();
  if (!/^\d{4,6}$/.test(pin)) throw new Error('Sai mã PIN.');
  var rows = sqCacheGet_('sq_nv');
  if (!rows) { rows = layFile_().getSheetByName(SQ_NV).getDataRange().getValues().map(function (r) { return r.map(String); }); sqCachePut_('sq_nv', rows, 300); }
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === pin && !/x/i.test(String(rows[i][3]))) return { ten: String(rows[i][0]).trim(), vai: String(rows[i][2]).trim(), chucNang: sqTachCN_(rows[i][4], rows[i][2]) };
  }
  throw new Error('Sai mã PIN.');
}
function sqChu_(pass) {
  var dung = PropertiesService.getScriptProperties().getProperty('CHU_PASS');
  if (!dung) throw new Error('Chưa tạo mật khẩu.');
  if (String(pass || '') !== dung) throw new Error('Sai mật khẩu.');
}

/* ======================= ĐỌC / GHI TAB ======================= */
// Đọc tab SQ_Phieu 1 lần mỗi lượt gọi (nhớ tạm trong lượt), ghi xong thì đọc lại
var _SQ_META = null;
function sqMetaMoi_() { _SQ_META = null; }
function sqMetaRows_() {
  if (_SQ_META) return _SQ_META;
  var sh = layFile_().getSheetByName(SQ_META);
  var v = sh.getDataRange().getValues();
  return _SQ_META = v.slice(1).map(function (r, i) { var o = sqMetaObj_(r); o._row = i + 2; return o; }).filter(function (o) { return o.id; });
}
// Bộ nhớ đệm của Google (CacheService) cho dữ liệu ít đổi: danh sách nhân viên, danh mục món
function sqCache_() { try { return CacheService.getScriptCache(); } catch (e) { return null; } }
function sqCacheGet_(k) { var c = sqCache_(); if (!c) return null; try { var v = c.get(k); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function sqCachePut_(k, v, giay) { var c = sqCache_(); if (!c) return; try { var s = JSON.stringify(v); if (s.length < 95000) c.put(k, s, giay || 300); } catch (e) { } }
function sqCacheXoa_(k) { var c = sqCache_(); if (c) try { c.remove(k); } catch (e) { } }
function sqJ_(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }
function sqMetaObj_(r) {
  return {
    id: String(r[M.ID]), loai: String(r[M.LOAI]), ngay: chuanNgay_(r[M.NGAY]), nguoi: String(r[M.NGUOI]), moi: String(r[M.MOI]),
    anh: String(r[M.ANH] || '').split(',').filter(String), ai: sqJ_(r[M.AI], null), dong: sqJ_(r[M.DONG], []),
    tongAnh: soHoacRong_(r[M.TONGANH]), cong: so_(r[M.CONG]), ghiChu: String(r[M.GHICHU]), tt: String(r[M.TT]) || TT.CHO,
    hoi: String(r[M.HOI]), traLoi: String(r[M.TRALOI]), duyetLuc: String(r[M.DUYETLUC]), taoLuc: String(r[M.TAOLUC]),
    them: sqJ_(r[M.THEM], {}), tong: so_(r[M.TONG])
  };
}
function sqMetaRow_(o) {
  var r = [];
  r[M.ID] = o.id; r[M.LOAI] = o.loai; r[M.NGAY] = o.ngay; r[M.NGUOI] = o.nguoi || ''; r[M.MOI] = o.moi || '';
  r[M.ANH] = (o.anh || []).join(','); r[M.AI] = o.ai ? JSON.stringify(o.ai).slice(0, 45000) : ''; r[M.DONG] = JSON.stringify(o.dong || []).slice(0, 45000);
  r[M.TONGANH] = o.tongAnh === undefined ? '' : o.tongAnh; r[M.CONG] = o.cong || 0; r[M.GHICHU] = o.ghiChu || ''; r[M.TT] = o.tt || TT.CHO;
  r[M.HOI] = o.hoi || ''; r[M.TRALOI] = o.traLoi || ''; r[M.DUYETLUC] = o.duyetLuc || ''; r[M.TAOLUC] = o.taoLuc || bayGio_();
  r[M.THEM] = JSON.stringify(o.them || {}); r[M.TONG] = o.tong || 0;
  return r;
}
function sqGhiMeta_(o) {
  var sh = layFile_().getSheetByName(SQ_META);
  var row = timDong_(sh, o.id);
  if (row < 0) row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, H_SQ_META.length).setValues([sqMetaRow_(o)]);
  sqMetaMoi_();
}
function sqTimMeta_(id) {
  var ds = sqMetaRows_();
  for (var i = ds.length - 1; i >= 0; i--) if (ds[i].id === id) return ds[i];
  return null;
}
function sqAlias_() {
  var v = layFile_().getSheetByName(SQ_ALIAS).getDataRange().getValues().slice(1);
  return v.filter(function (r) { return String(r[1]).trim() && String(r[2]).trim(); })
    .map(function (r) { return { moi: String(r[0]).trim(), chu: String(r[1]).trim(), ten: String(r[2]).trim() }; });
}
function sqNhoAlias_(moi, ds, nguoi) {
  if (!ds || !ds.length) return;
  var sh = layFile_().getSheetByName(SQ_ALIAS);
  var cu = sqAlias_();
  ds.forEach(function (a) {
    var chu = chu_(a.chu, 80), ten = chu_(a.ten, 80);
    if (!chu || !ten || khongDau_(chu) === khongDau_(ten)) return;
    var co = cu.some(function (x) { return khongDau_(x.moi) === khongDau_(moi) && khongDau_(x.chu) === khongDau_(chu); });
    if (!co) sh.appendRow([moi || '', chu, ten, nguoi, bayGio_()]);
  });
}

/* ======================= DỮ LIỆU CHO APP NHÂN VIÊN ======================= */
function sqDanhMucGon_() {
  var c = sqCacheGet_('sq_dm');
  if (c) return c;
  var dm = docDanhMuc_(), gop = {};
  dm.forEach(function (d) {
    var k = khongDau_(d.ten);
    if (!gop[k] || d.ngay > gop[k].ngay) gop[k] = { ten: d.ten, dv: d.dv, gia: d.gia, ngay: d.ngay, nhom: d.nhom, lan: d.lan };
  });
  var out = { danhMuc: Object.keys(gop).map(function (k) { return gop[k]; }),
    tuDat: dm.filter(function (d) { return d.nhom === NHOM_TU_DAT; }).sort(function (a, b) { return b.lan - a.lan; }).slice(0, 24) };
  sqCachePut_('sq_dm', out, 600);
  return out;
}
function sqDuLieuNV_(nv) {
  sqDungTongHopNeuCan_();
  var dmg = sqDanhMucGon_();
  var metas = sqMetaRows_();
  var moiSet = {};
  metas.forEach(function (m) { if (m.moi) moiSet[m.moi] = (moiSet[m.moi] || 0) + 1; });
  var tu = congNgay_(homNay_(), -3);
  var cuaToi = metas.filter(function (m) { return m.nguoi === nv.ten && TT_CHUA_XAC_NHAN.indexOf(m.tt) < 0 && (m.ngay >= tu || m.tt === TT.HOI); })
    .map(function (m) { return { id: m.id, loai: m.loai, ngay: m.ngay, moi: m.moi, tong: m.tong || m.cong, tt: m.tt, hoi: m.hoi, traLoi: m.traLoi }; })
    .reverse().slice(0, 30);
  // Hàng tự đặt: nhớ lần giao gần nhất của từng mối (món, số lượng, giá) để lần sau bấm 1 cái là điền lại
  var tuDatMoi = {};
  metas.forEach(function (m) {
    if (m.loai !== LOAI.TD || !m.moi) return;
    var k = khongDau_(m.moi), cu = tuDatMoi[k];
    if (!cu || m.ngay >= cu.ngay) tuDatMoi[k] = { moi: m.moi, ngay: m.ngay, traLien: !!(m.them && m.them.traLien),
      mon: (m.dong || []).map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia }; }) };
  });
  return {
    ten: nv.ten, vai: nv.vai, chucNang: nv.chucNang, homNay: homNay_(),
    huy: sqHuyChoNV_(nv),
    tuDatMoi: Object.keys(tuDatMoi).map(function (k) { return tuDatMoi[k]; }).sort(function (a, b) { return a.ngay < b.ngay ? 1 : -1; }).slice(0, 20),
    danhMuc: dmg.danhMuc,
    tuDat: dmg.tuDat,
    alias: sqAlias_(),
    moi: Object.keys(moiSet).sort(function (a, b) { return moiSet[b] - moiSet[a]; }).slice(0, 20),
    cuaToi: cuaToi,
    quen: sqQuenDs_(metas),
    choKiem: metas.filter(function (m) { return m.loai === LOAI.HD && (m.tt === TT.KIEM || m.tt === TT.LOI || m.tt === TT.DOC); }).length,
    diCho: ScriptApp.getService().getUrl()
  };
}

/* ======================= ĐỌC ẢNH BẰNG GEMINI ======================= */
function sqThuMuc_() {
  var id = PropertiesService.getScriptProperties().getProperty('SQ_FOLDER_ID');
  if (!id) throw new Error('Chưa chạy sqCaiDat.');
  return DriveApp.getFolderById(id);
}
function sqDocAnh_(p) {
  var anh = (p.anh || []).slice(0, 4);
  if (!anh.length) throw new Error('Chưa có ảnh.');
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_KEY');
  if (!key) throw new Error('Chưa cài mã Gemini (GEMINI_KEY).');
  var model = props.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
  // 1. lưu ảnh lên Drive (để anh Đạt xem lại)
  var folder = sqThuMuc_(), fids = [];
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss');
  anh.forEach(function (a, i) {
    var blob = Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'image/jpeg', 'HD-' + stamp + '-' + (i + 1) + '.jpg');
    fids.push(folder.createFile(blob).getId());
  });
  // 2. gợi ý danh mục + tên gọi khác của mối
  var dm = docDanhMuc_(), tenDm = {}, ds = [];
  dm.forEach(function (d) { var k = khongDau_(d.ten); if (!tenDm[k]) { tenDm[k] = d.ten; ds.push(d.ten); } });
  var moi = chu_(p.moi, 60);
  var alias = sqAlias_().filter(function (a) { return !moi || !a.moi || khongDau_(a.moi) === khongDau_(moi); });
  var prompt = [
    'Đây là ' + anh.length + ' ảnh của MỘT hóa đơn mua hàng của quán ăn Việt Nam (có thể viết tay, có thể là hóa đơn in, có thể chụp nghiêng).',
    'Hãy đọc TỪNG dòng hàng hóa và trả về JSON đúng schema.',
    'Quy ước:',
    '- Số tiền trên hóa đơn viết tay thường ghi theo NGHÌN đồng (vd "950" = 950.000đ, "1.2tr" = 1.200.000đ). Luôn trả gia và tien theo ĐỒNG.',
    '- sl là số lượng (số thập phân dùng dấu chấm), dv là đơn vị (kg, con, bó, quả, hũ, gói, thùng, lít…).',
    '- chu = chữ đúng y như viết trên hóa đơn. ten = tên món chuẩn: nếu khớp món trong DANH MỤC bên dưới thì ghi đúng tên đó, không thì viết lại cho rõ nghĩa.',
    '- Nếu chữ hoặc số không chắc chắn (mờ, khó đọc, sl×gia không ra tien) thì chacChan=false và lyDo ngắn bằng tiếng Việt. Ngược lại chacChan=true, lyDo rỗng.',
    '- anh = số thứ tự ảnh chứa dòng đó (ảnh đầu = 0). vung = [ymin, xmin, ymax, xmax] theo thang 0–1000 bao quanh cả dòng đó trên ảnh.',
    '- Phí giao hàng, phí ship cũng là 1 dòng. Không bỏ sót dòng nào.',
    '- tongTrenAnh = tổng tiền ghi trên hóa đơn (đồng), không có thì 0. ngay = ngày trên hóa đơn dạng yyyy-mm-dd, không có thì rỗng. moi = tên người bán / cửa hàng nếu có.',
    'DANH MỤC món quán hay mua: ' + ds.slice(0, 400).join('; '),
    alias.length ? 'TÊN GỌI KHÁC đã biết (chữ trên hóa đơn → tên đúng): ' + alias.slice(0, 200).map(function (a) { return a.chu + ' → ' + a.ten; }).join('; ') : '',
    moi ? 'Nhân viên cho biết hóa đơn này của mối: ' + moi : ''
  ].filter(String).join('\n');
  var parts = [{ text: prompt }];
  anh.forEach(function (a) { parts.push({ inline_data: { mime_type: a.mime || 'image/jpeg', data: a.b64 } }); });
  var schema = {
    type: 'OBJECT',
    properties: {
      moi: { type: 'STRING' }, ngay: { type: 'STRING' }, tongTrenAnh: { type: 'NUMBER' },
      dong: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
        chu: { type: 'STRING' }, ten: { type: 'STRING' }, sl: { type: 'NUMBER' }, dv: { type: 'STRING' },
        gia: { type: 'NUMBER' }, tien: { type: 'NUMBER' }, chacChan: { type: 'BOOLEAN' }, lyDo: { type: 'STRING' },
        anh: { type: 'INTEGER' }, vung: { type: 'ARRAY', items: { type: 'INTEGER' } }
      }, required: ['chu', 'ten', 'tien', 'chacChan'] } }
    },
    required: ['dong']
  };
  var body = sqGoiGemini_(key, model, parts, schema), dung = body.model;
  body = body.body;
  var j = JSON.parse(body), txt = '';
  try { txt = j.candidates[0].content.parts.map(function (x) { return x.text || ''; }).join(''); } catch (e) { throw new Error('Gemini không trả kết quả.'); }
  var doc = sqJ_(txt, null);
  if (!doc || !doc.dong) throw new Error('Không đọc được hóa đơn, chụp lại rõ hơn nhé.');
  // 3. áp tên gọi khác + đối chiếu danh mục
  doc.dong = (doc.dong || []).map(function (d) { return sqChuanDong_(d, alias, tenDm); });
  if (!moi && doc.moi) moi = doc.moi;
  return { fids: fids, moi: moi, ngay: /^\d{4}-\d{2}-\d{2}$/.test(doc.ngay || '') ? doc.ngay : '', tongAnh: so_(doc.tongTrenAnh), dong: doc.dong, model: dung };
}
// Gọi Gemini: thử lại khi Google quá tải (429/500/503), rồi chuyển sang model nhẹ hơn.
var SQ_MODEL_DU_PHONG = ['gemini-3.5-flash-lite', 'gemini-flash-lite-latest', 'gemini-flash-latest'];
function sqGoiGemini_(key, model, parts, schema) {
  var ds = [model].concat(SQ_MODEL_DU_PHONG.filter(function (m) { return m !== model; }));
  var batDau = Date.now(), loiCuoi = '', quaTai = false;
  for (var i = 0; i < ds.length; i++) {
    var gc = { temperature: 0, responseMimeType: 'application/json', responseSchema: schema, thinkingConfig: { thinkingLevel: 'low' } };
    for (var lan = 0; lan < 2; lan++) {
      if (Date.now() - batDau > 100000) break; // tối đa ~100 giây
      var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + ds[i] + ':generateContent', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-goog-api-key': key },
        payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: gc })
      });
      var code = res.getResponseCode(), body = res.getContentText();
      if (code === 200) return { body: body, model: ds[i] };
      loiCuoi = code + ': ' + body.slice(0, 200);
      if (code === 400 && gc.thinkingConfig && /thinking/i.test(body)) { delete gc.thinkingConfig; lan--; continue; }
      if (code === 400 || code === 401 || code === 403) {
        if (/API key|API_KEY|PERMISSION/i.test(body)) throw new Error('Mã Gemini (GEMINI_KEY) không dùng được, anh Đạt kiểm tra lại giúp.');
        break; // model này không nhận, thử model khác
      }
      if (code === 404) break;
      if (code === 429 || code >= 500) { quaTai = true; Utilities.sleep(lan === 0 ? 1500 : 3000); continue; }
      break;
    }
  }
  if (quaTai) throw new Error('Máy đọc hóa đơn của Google đang quá tải. Đợi 1–2 phút rồi bấm đọc lại nhé (ảnh không bị mất).');
  throw new Error('Gemini báo lỗi ' + loiCuoi);
}
function sqChuanDong_(d, alias, tenDm) {
  var o = {
    chu: chu_(d.chu, 80), ten: chu_(d.ten || d.chu, 80), sl: soHoacRong_(d.sl), dv: chu_(d.dv, 15),
    gia: soHoacRong_(d.gia), tien: Math.round(so_(d.tien)), chacChan: d.chacChan !== false, lyDo: chu_(d.lyDo, 120),
    anh: so_(d.anh), vung: Array.isArray(d.vung) && d.vung.length === 4 ? d.vung.map(so_) : null, daNho: false
  };
  var kc = khongDau_(o.chu);
  for (var i = 0; i < alias.length; i++) if (khongDau_(alias[i].chu) === kc) { o.ten = alias[i].ten; o.daNho = true; break; }
  var kt = khongDau_(o.ten);
  if (tenDm[kt]) o.ten = tenDm[kt];
  else if (!o.daNho && o.ten && !/ph[ií] (giao|ship)|ship|v[aậ]n chuy[eể]n/i.test(o.ten)) {
    o.moiMon = true;
    if (o.chacChan) { o.chacChan = false; o.lyDo = 'Tên mới, chưa có trong danh mục'; }
  }
  if (o.tien && o.sl !== '' && o.gia !== '' && Math.abs(o.sl * o.gia - o.tien) > Math.max(1000, o.tien * 0.02)) {
    if (o.chacChan) { o.chacChan = false; o.lyDo = 'Số lượng × đơn giá không ra thành tiền'; }
  }
  if (o.gia === '' && o.sl && o.tien) o.gia = Math.round(o.tien / o.sl);
  return o;
}

/* ======================= LƯU PHIẾU ======================= */
function sqMon_(dong) {
  return (dong || []).slice(0, 200).map(function (d) {
    var sl = soHoacRong_(d.sl), gia = soHoacRong_(d.gia), tien = soHoacRong_(d.tien);
    if (tien === '' && sl !== '' && gia !== '') tien = sl * gia;
    return { ten: chu_(d.ten, 80), sl: sl, dv: chu_(d.dv, 15), gia: gia === '' ? '' : Math.round(gia), tien: Math.round(so_(tien)),
      chu: chu_(d.chu, 80), sua: !!d.sua, aiTen: chu_(d.aiTen, 80), aiSl: soHoacRong_(d.aiSl), aiGia: soHoacRong_(d.aiGia), aiTien: soHoacRong_(d.aiTien),
      anh: so_(d.anh), vung: Array.isArray(d.vung) ? d.vung.map(so_) : null };
  }).filter(function (m) { return m.ten; });
}
function sqKiemId_(id) { if (!/^[A-Za-z0-9_-]{4,40}$/.test(String(id || ''))) throw new Error('Mã phiếu không hợp lệ.'); return String(id); }
function sqKiemNgay_(s) { if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) throw new Error('Ngày không hợp lệ.'); return String(s); }
function sqGiuDuyet_(cu, nv) {
  if (cu && cu.nguoi && cu.nguoi !== nv.ten && cu.loai !== LOAI.HD && cu.loai !== LOAI.TD) throw new Error('Phiếu này không phải của em.');
  if (cu && cu.tt === TT.OK) throw new Error('Phiếu đã được anh duyệt, không sửa được nữa.');
  return cu ? (cu.tt === TT.HOI ? TT.TRALOI : cu.tt === TT.TUCHOI ? TT.TRALOI : TT.CHO) : TT.CHO;
}

// Lưu phiếu vào Phieu/ChiTiet bằng apiLuu của app đi chợ, nhưng KHÔNG dựng lại tab TongHopNgay ngay
// (việc này đọc cả tab Phieu và ghi lại cả bảng, tốn vài giây). Đánh dấu để dựng lại lúc app tải dữ liệu nền.
function sqApiLuu_(p) {
  var co = typeof dungTongHop_ === 'function', goc = co ? dungTongHop_ : null;
  if (co) dungTongHop_ = function () { };
  sqCacheXoa_('sq_dm');
  try { return apiLuu(p); }
  finally {
    if (co) { dungTongHop_ = goc; PropertiesService.getScriptProperties().setProperty('SQ_TONGHOP_CU', '1'); }
  }
}
function sqDungTongHopNeuCan_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SQ_TONGHOP_CU') !== '1' || typeof dungTongHop_ !== 'function') return;
  try { khoa_(function () { dungTongHop_(layFile_()); }); props.deleteProperty('SQ_TONGHOP_CU'); } catch (e) { }
}

function sqLuuHD_(nv, p) {
  var id = sqKiemId_(p.id), ngay = sqKiemNgay_(p.ngay);
  var moi = chu_(p.moi, 60) || 'Chưa rõ mối';
  var dong = sqMon_(p.dong);
  if (!dong.length) throw new Error('Hóa đơn chưa có dòng nào.');
  return khoa_(function () {
    var cu = sqTimMeta_(id), tt = sqGiuDuyet_(cu, nv);
    var cong = dong.reduce(function (s, d) { return s + d.tien; }, 0);
    sqApiLuu_({ id: id, ngay: ngay, nhom: 'HĐ ' + moi, nguoiMua: nv.ten, ghiChu: chu_(p.ghiChu, 200),
      mon: dong.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) });
    var anh = (cu && cu.anh && cu.anh.length) ? cu.anh : (p.fids || []).filter(function (x) { return /^[\w-]{10,}$/.test(x); });
    var them = Object.assign({}, (cu && cu.them) || {}, { suaLuc: bayGio_(), kiemBoi: nv.ten });
    var o = {
      id: id, loai: LOAI.HD, ngay: ngay, nguoi: cu && cu.nguoi ? cu.nguoi : nv.ten, moi: moi, anh: anh,
      ai: (cu && cu.ai) || p.ai || null, dong: dong, tongAnh: soHoacRong_(p.tongAnh), cong: cong, ghiChu: chu_(p.ghiChu, 200),
      tt: tt, hoi: cu ? cu.hoi : '', traLoi: cu ? cu.traLoi : '', taoLuc: cu ? cu.taoLuc : bayGio_(), them: them, tong: cong
    };
    sqGhiMeta_(o);
    sqNhoAlias_(moi, p.nho, nv.ten);
    sqHocTuSua_(moi, o.ai, dong, nv.ten);
    return { id: id, tong: cong, tt: tt };
  });
}

function sqLuuTuDat_(nv, p) {
  var id = sqKiemId_(p.id), ngay = sqKiemNgay_(p.ngay);
  var moi = chu_(p.moi, 60);
  if (!moi) throw new Error('Chưa ghi mối giao.');
  var dong = sqMon_(p.mon);
  if (!dong.length) throw new Error('Phiếu chưa có món nào.');
  var traLien = p.traLien === true;
  return khoa_(function () {
    var cu = sqTimMeta_(id), tt = sqGiuDuyet_(cu, nv);
    var cong = dong.reduce(function (s, d) { return s + d.tien; }, 0);
    var anh = cu ? cu.anh : [];
    if (p.anh && p.anh.b64) anh = [sqThuMuc_().createFile(Utilities.newBlob(Utilities.base64Decode(p.anh.b64), p.anh.mime || 'image/jpeg', 'TD-' + id + '.jpg')).getId()];
    var quenId = sqQuenSauLuu_(nv, p, moi, dong, traLien);
    var ghiChu = [moi, traLien ? 'Trả liền' : 'Ghi nợ mối', chu_(p.ghiChu, 150)].filter(String).join(' · ');
    sqApiLuu_({ id: id, ngay: ngay, nhom: NHOM_TU_DAT, nguoiMua: nv.ten, ghiChu: ghiChu,
      mon: dong.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) });
    sqGhiMeta_({ id: id, loai: LOAI.TD, ngay: ngay, nguoi: cu && cu.nguoi ? cu.nguoi : nv.ten, moi: moi, anh: anh, dong: dong, cong: cong, ghiChu: chu_(p.ghiChu, 150),
      tt: tt, hoi: cu ? cu.hoi : '', traLoi: cu ? cu.traLoi : '', taoLuc: cu ? cu.taoLuc : bayGio_(), them: { traLien: traLien, suaLuc: bayGio_(), quenId: quenId || (cu && cu.them && cu.them.quenId) || '' }, tong: cong });
    return { id: id, tong: cong, tt: tt, quenId: quenId };
  });
}

var LY_DO_BILL = ['Đãi khách / người quen', 'Bồi thường sự cố / voucher', 'Nhân viên / gia đình ăn'];
function sqLuuBill_(nv, p) {
  var id = sqKiemId_(p.id), ngay = sqKiemNgay_(p.ngay);
  var lyDo = chu_(p.lyDo, 60), choPhep = chu_(p.choPhep, 40);
  if (!lyDo) throw new Error('Chưa chọn lý do.');
  if (!choPhep) throw new Error('Chưa chọn ai cho phép.');
  var dong = sqMon_(p.mon);
  if (!dong.length) throw new Error('Bill chưa có món nào.');
  return khoa_(function () {
    var cu = sqTimMeta_(id), tt = sqGiuDuyet_(cu, nv);
    var anh = cu ? cu.anh : [];
    if (p.anhBill && p.anhBill.b64) {
      var blob = Utilities.newBlob(Utilities.base64Decode(p.anhBill.b64), p.anhBill.mime || 'image/jpeg', 'BILL-' + id + '.jpg');
      anh = [sqThuMuc_().createFile(blob).getId()];
    }
    var tong = dong.reduce(function (s, d) { return s + d.tien; }, 0);
    sqGhiMeta_({ id: id, loai: LOAI.BILL, ngay: ngay, nguoi: nv.ten, moi: '', anh: anh, dong: dong, cong: tong, ghiChu: chu_(p.ghiChu, 200),
      tt: tt, hoi: cu ? cu.hoi : '', traLoi: cu ? cu.traLoi : '', taoLuc: cu ? cu.taoLuc : bayGio_(),
      them: { lyDo: lyDo, choPhep: choPhep, ban: chu_(p.ban, 20), khach: chu_(p.khach, 60), soHD: chu_(p.soHD, 20), suaLuc: bayGio_() }, tong: tong });
    return { id: id, tong: tong, tt: tt };
  });
}

function sqTraLoi_(nv, p) {
  var id = sqKiemId_(p.id);
  return khoa_(function () {
    var cu = sqTimMeta_(id);
    if (!cu || cu.nguoi !== nv.ten) throw new Error('Không tìm thấy phiếu của em.');
    cu.traLoi = chu_(p.traLoi, 300); cu.tt = TT.TRALOI;
    sqGhiMeta_(cu);
    return { id: id };
  });
}

/* ======================= PHẦN CỦA ANH ĐẠT ======================= */
// Phiếu (Phieu + ChiTiet) trong khoảng ngày, kèm meta Sổ Quán nếu có
function sqPhieuTrong_(tu, den) {
  var ds = docPhieu_(function (p) { return p.ngay >= tu && p.ngay <= den; });
  var metas = {};
  sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  var out = ds.map(function (p) {
    var m = metas[p.id];
    var x = { id: p.id, ngay: p.ngay, nhom: p.nhom, nguoi: p.nguoiMua, tong: p.tong, tienNhan: p.tienNhan, ghiChu: p.ghiChu, mon: p.mon,
      loai: m ? m.loai : LOAI.DC, moi: m ? m.moi : '', tt: m ? m.tt : TT.CHO, hoi: m ? m.hoi : '', traLoi: m ? m.traLoi : '',
      anh: m ? m.anh : [], tongAnh: m ? m.tongAnh : '', soSua: m ? m.dong.filter(function (d) { return d.sua; }).length : 0, suaLuc: p.suaLuc };
    delete metas[p.id];
    return x;
  });
  // bill không tính tiền (chỉ có trong meta)
  Object.keys(metas).forEach(function (k) {
    var m = metas[k];
    if (m.loai !== LOAI.BILL || m.ngay < tu || m.ngay > den) return;
    out.push({ id: m.id, ngay: m.ngay, nhom: LOAI.BILL, nguoi: m.nguoi, tong: m.tong, mon: m.dong, loai: m.loai, moi: '', tt: m.tt, hoi: m.hoi,
      traLoi: m.traLoi, anh: m.anh, them: m.them, ghiChu: m.ghiChu, suaLuc: m.taoLuc });
  });
  return out;
}

function sqDuyetDs_(p) {
  var den = homNay_(), tu = congNgay_(den, -(Math.min(so_(p.soNgay) || 21, 120)));
  var ds = sqPhieuTrong_(tu, den);
  var gia = sqTinhGia_(sqDongMua_());
  // gắn cảnh báo cho từng phiếu
  var billTuan = {};
  ds.forEach(function (x) { if (x.loai === LOAI.BILL) { var k = (x.them && x.them.lyDo) + '|' + sqTuan_(x.ngay); billTuan[k] = (billTuan[k] || 0) + 1; } });
  ds.forEach(function (x) {
    var cb = [];
    (gia.theoPhieu[x.id] || []).forEach(function (g) {
      if (Math.abs(g.pct) >= 0.05) cb.push({ muc: Math.abs(g.pct) >= 0.15 ? 'r' : 'y', text: g.ten + ' ' + sqTien_(g.gia) + '/' + (g.dv || 'đv') + ', ' + (g.pct > 0 ? 'tăng ' : 'giảm ') +
        Math.round(Math.abs(g.pct) * 100) + '% so với lần mua ' + sqDM_(g.truoc.ngay) + ' (' + sqTien_(g.truoc.gia) + ')' });
    });
    if (x.loai === LOAI.HD && x.tongAnh !== '' && x.tongAnh > 0 && Math.abs(x.tongAnh - x.tong) > 1000)
      cb.push({ muc: 'r', text: 'Cộng các dòng lệch ' + sqTien_(x.tong - x.tongAnh) + ' so với tổng trên ảnh (' + sqTien_(x.tongAnh) + ')' });
    if (x.soSua) cb.push({ muc: 'i', text: x.soSua + ' dòng nhân viên sửa tay so với AI đọc' });
    if (x.loai === LOAI.BILL) {
      if (!x.anh || !x.anh.length) cb.push({ muc: 'r', text: 'Không có ảnh bill Sapo' });
      var n = billTuan[(x.them && x.them.lyDo) + '|' + sqTuan_(x.ngay)];
      if (n >= 3) cb.push({ muc: 'y', text: 'Tuần này đã có ' + n + ' bill cùng lý do "' + (x.them && x.them.lyDo) + '"' });
    }
    if (x.tong && x.tong % 100000 === 0 && x.loai !== LOAI.BILL) cb.push({ muc: 'y', text: 'Tổng phiếu tròn ' + sqTien_(x.tong) });
    x.canhBao = cb;
  });
  ds.sort(function (a, b) { return a.ngay < b.ngay ? 1 : a.ngay > b.ngay ? -1 : (String(a.suaLuc) < String(b.suaLuc) ? 1 : -1); });
  return { ds: ds, tu: tu, den: den };
}
function sqTuan_(s) { var p = s.split('-'), d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])); var dow = (d.getUTCDay() + 6) % 7; d.setUTCDate(d.getUTCDate() - dow); return d.toISOString().slice(0, 10); }
function sqTien_(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.') + 'đ'; }
function sqDM_(s) { var p = String(s).split('-'); return p[2] + '/' + p[1]; }

function sqDuyet_(p) {
  var id = sqKiemId_(p.id);
  var tt = [TT.OK, TT.HOI, TT.TUCHOI, TT.CHO].indexOf(p.tt) >= 0 ? p.tt : null;
  if (!tt) throw new Error('Trạng thái không hợp lệ.');
  return khoa_(function () {
    var m = sqTimMeta_(id);
    if (!m) { // phiếu đi chợ từ app cũ → tạo dòng meta
      var ph = docPhieu_(function (x) { return x.id === id; })[0];
      if (!ph) throw new Error('Không tìm thấy phiếu.');
      m = { id: id, loai: LOAI.DC, ngay: ph.ngay, nguoi: ph.nguoiMua, moi: ph.nhom, anh: [], dong: [], cong: ph.tong, tong: ph.tong, taoLuc: bayGio_(), them: {} };
    }
    m.tt = tt;
    if (tt === TT.HOI || tt === TT.TUCHOI) { m.hoi = chu_(p.hoi, 300); m.traLoi = ''; }
    m.duyetLuc = bayGio_();
    sqGhiMeta_(m);
    return { id: id, tt: tt };
  });
}

function sqChiTietPhieu_(id) {
  id = sqKiemId_(id);
  var m = sqTimMeta_(id);
  var ph = docPhieu_(function (x) { return x.id === id; })[0] || null;
  if (!m && !ph) throw new Error('Không tìm thấy phiếu.');
  return { id: id, meta: m, phieu: ph };
}

function sqAnh_(fid) {
  if (!/^[\w-]{10,}$/.test(String(fid || ''))) throw new Error('Ảnh không hợp lệ.');
  var f = DriveApp.getFileById(fid);
  var b = f.getBlob();
  return { b64: Utilities.base64Encode(b.getBytes()), mime: b.getContentType() };
}

/* ======================= GIÁ NHẬP THAY ĐỔI ======================= */
// Mỗi dòng mua: {id, ngay, nhom, stt, ten, sl, dv, gia, tien}
function sqDongMua_() {
  var sh = layFile_().getSheetByName('ChiTiet');
  var n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 9).getValues().map(function (r) {
    return { id: String(r[0]), ngay: chuanNgay_(r[1]), nhom: String(r[2]), stt: so_(r[3]), ten: String(r[4]), sl: soHoacRong_(r[5]), dv: String(r[6]), gia: soHoacRong_(r[7]), tien: so_(r[8]) };
  }).filter(function (r) { return r.id && r.ten; });
}
function sqDonVi_(dv) {
  var d = khongDau_(dv).replace(/\./g, '');
  if (/^(kg|ky|ki|kilo|kgs)$/.test(d)) return 'kg';
  if (/^(trai|qua)$/.test(d)) return 'quả';
  if (/^(l|lit|lit\.)$/.test(d)) return 'lít';
  return d;
}
var SQ_NHOM_HANG = [
  ['Gà', /(^| )(ga|chan ga|sun ga|canh ga|dui ga|uc ga|vit)( |$)/],
  ['Thịt heo', /(heo|3 roi|ba roi|ba chi|nac|suon|gio heo|mo heo|^mo$|da heo|thit xay|cot let)/],
  ['Bò', /(^| )bo( |$)/],
  ['Cá – hải sản', /(^| )(ca|tom|tep|muc|ech|cua|ghe|oc|so|hen|nghieu)( |$)/],
  ['Rau', /(rau|cai|xa lach|ngo|hanh la|hanh|gia|he|bac ha|mong toi|den|muong|tia to|kinh gioi|la chanh|la chuoi|ot|toi|gung|sa|me|bi|muop|dau bap|nam|ca chua|khoai|cu)/],
  ['Trái cây', /(xoai|coc|thom|mit|chuoi|dua|tac|chanh|khe|cam|buoi|oi|trai cay)/],
  ['Trứng – sữa', /(trung|sua chua|sua)/]
];
function sqNhomHang_(ten) {
  var k = khongDau_(ten);
  for (var i = 0; i < SQ_NHOM_HANG.length; i++) if (SQ_NHOM_HANG[i][1].test(k)) return SQ_NHOM_HANG[i][0];
  return 'Khác';
}
// Tính giá mỗi đơn vị, so với lần mua gần nhất TRƯỚC ĐÓ của cùng mặt hàng + cùng đơn vị
function sqTinhGia_(dongs) {
  var theoMon = {};
  dongs.forEach(function (r) {
    var sl = r.sl === '' ? 0 : Number(r.sl), dv = sqDonVi_(r.dv);
    var gia = r.gia !== '' && Number(r.gia) > 0 ? Number(r.gia) : (sl > 0 && r.tien > 0 ? r.tien / sl : 0);
    if (!(gia > 0) || !(sl > 0)) return;
    var khoa = khongDau_(r.ten) + '|' + dv;
    (theoMon[khoa] = theoMon[khoa] || []).push({ id: r.id, ngay: r.ngay, nhom: r.nhom, stt: r.stt, ten: r.ten, sl: sl, dv: r.dv || dv, gia: Math.round(gia), tien: r.tien });
  });
  var bienDong = [], theoPhieu = {};
  Object.keys(theoMon).forEach(function (khoa) {
    var ds = theoMon[khoa].sort(function (a, b) { return a.ngay < b.ngay ? -1 : a.ngay > b.ngay ? 1 : 0; });
    for (var i = 1; i < ds.length; i++) {
      var cur = ds[i], prev = null;
      for (var j = i - 1; j >= 0; j--) if (ds[j].ngay < cur.ngay) { prev = ds[j]; break; }
      if (!prev) continue;
      var pct = (cur.gia - prev.gia) / prev.gia;
      var g = { khoa: khoa, ten: cur.ten, nhomHang: sqNhomHang_(cur.ten), dv: cur.dv, id: cur.id, stt: cur.stt, ngay: cur.ngay, nhom: cur.nhom, sl: cur.sl, gia: cur.gia, tien: cur.tien,
        truoc: { id: prev.id, ngay: prev.ngay, gia: prev.gia, sl: prev.sl, tien: prev.tien, nhom: prev.nhom }, chenh: cur.gia - prev.gia, pct: pct, traThem: Math.round((cur.gia - prev.gia) * cur.sl) };
      if (Math.abs(pct) >= 0.005) bienDong.push(g);
      (theoPhieu[cur.id] = theoPhieu[cur.id] || []).push(g);
    }
  });
  return { bienDong: bienDong, theoMon: theoMon, theoPhieu: theoPhieu };
}
function sqGiaThayDoi_(p) {
  var den = /^\d{4}-\d{2}-\d{2}$/.test(p.den || '') ? p.den : homNay_();
  var tu = /^\d{4}-\d{2}-\d{2}$/.test(p.tu || '') ? p.tu : congNgay_(den, -6);
  var g = sqTinhGia_(sqDongMua_());
  var nguoi = {};
  docPhieu_(function (x) { return x.ngay >= congNgay_(tu, -120); }).forEach(function (x) { nguoi[x.id] = x.nguoiMua; });
  var ds = g.bienDong.filter(function (x) { return x.ngay >= tu && x.ngay <= den; });
  ds.forEach(function (x) { x.nguoi = nguoi[x.id] || ''; x.truoc.nguoi = nguoi[x.truoc.id] || ''; });
  ds.sort(function (a, b) { return Math.abs(b.pct) - Math.abs(a.pct); });
  var traThem = ds.reduce(function (s, x) { return s + (x.traThem > 0 ? x.traThem : 0); }, 0);
  var botDuoc = ds.reduce(function (s, x) { return s + (x.traThem < 0 ? -x.traThem : 0); }, 0);
  return { tu: tu, den: den, ds: ds, traThem: traThem, botDuoc: botDuoc };
}
function sqLichSuMon_(khoa) {
  khoa = String(khoa || '');
  var g = sqTinhGia_(sqDongMua_());
  var ds = (g.theoMon[khoa] || []).slice(-40);
  var metas = {};
  sqMetaRows_().forEach(function (m) { metas[m.id] = m; });
  var nguoi = {};
  var ids = {};
  ds.forEach(function (x) { ids[x.id] = true; });
  docPhieu_(function (x) { return ids[x.id]; }).forEach(function (x) { nguoi[x.id] = x.nguoiMua; });
  ds.forEach(function (x) {
    x.nguoi = nguoi[x.id] || '';
    var m = metas[x.id];
    if (m) {
      x.loai = m.loai; x.tt = m.tt; x.fids = m.anh;
      var d = (m.dong || [])[x.stt - 1];
      if (d && khongDau_(d.ten) === khongDau_(x.ten)) { x.anhSo = d.anh; x.vung = d.vung; x.chu = d.chu; }
    } else { x.loai = LOAI.DC; x.tt = TT.CHO; x.fids = []; }
  });
  return { khoa: khoa, ds: ds };
}

/* =====================================================================
 * HÓA ĐƠN BẢN 2 (09/11/2026)
 * - Gửi nhiều ảnh 1 lần: mỗi ảnh (hoặc vài ảnh gộp) = 1 hóa đơn, lưu ngay với trạng thái "Đang đọc".
 * - Đọc ở nền: app gọi nv_doc_nen, Gemini đọc song song (fetchAll) → "Chờ kiểm tra".
 * - Nhân viên mở "Hóa đơn đã gửi", sửa, bấm Xác nhận đúng (nv_luu_hd) → mới ghi Phieu/ChiTiet (P&L).
 * - Mỗi chỗ sửa được nhớ theo mối (tab SQ_HocMoi) và đưa vào lời nhắc lần sau.
 * - Hóa đơn quen (tab SQ_HoaDonQuen): bấm 1 nút là điền sẵn phiếu hàng tự đặt.
 * ===================================================================== */
var SQ_HOC = 'SQ_HocMoi', SQ_QUEN = 'SQ_HoaDonQuen';
var H_SQ_HOC = ['Mối', 'Chữ trên ảnh', 'Tên đúng', 'SL', 'ĐV', 'Đơn giá', 'Thành tiền', 'Ghi chú', 'Người dạy', 'Lúc'];
var H_SQ_QUEN = ['ID', 'Tên', 'Biểu tượng', 'Mối', 'Món (JSON)', 'Trả liền', 'Người tạo', 'Sửa lúc'];
// Đáp án anh Đạt xác nhận 09/11 – dùng làm mẫu học ban đầu
var SQ_HOC_MAU = [
  ['Mối cá (Zalo)', 'Cá lóc năm ký nhân giá 75.000 bằng 375.000', 'Cá lóc', 5, 'kg', 75000, 375000, ''],
  ['Mối cá (Zalo)', 'Hông 5x75=375', 'Cá diêu hồng', 5, 'kg', 75000, 375000, '"Hông" / "hồng" của mối cá là CÁ DIÊU HỒNG. "+375=750" là tổng, không phải món'],
  ['Gas Petrolimex', 'Khí dầu mỏ hóa lỏng bình 12 · 1.084.000 · Bên bán giao 2', 'Gas bình 12kg', 2, 'bình', 542000, 1084000, 'Số ghi tay ở cột "Số lượng nguyên niêm" là TỔNG TIỀN; số bình lấy ở "Số lượng Bên bán giao"'],
  ['Mối trứng (Zalo)', 'Trứng gà: 5 vĩ* 22 = 330k', 'Trứng gà', 5, 'vỉ', 66000, 330000, '"22" là 2.200đ một quả, 1 vỉ 30 quả'],
  ['Mối trứng (Zalo)', 'Trứng làm bánh: 20 trứng bỏ riêng* 22 = 44', 'Trứng làm bánh', 20, 'quả', 2200, 44000, ''],
  ['Mối trứng (Zalo)', 'Trứng cút: 30 trứng = 15k', 'Trứng cút', 30, 'quả', 500, 15000, '']
];
function sqTab_(ten, head, napMau) {
  var ss = layFile_(), sh = ss.getSheetByName(ten);
  if (sh) return sh;
  sh = ss.insertSheet(ten);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#445803').setFontColor('#F8E8C4');
  sh.setFrozenRows(1);
  if (napMau) napMau(sh);
  return sh;
}
function sqHdCaiDat_() { sqTabHoc_(); sqTabQuen_(); }
function sqTabHoc_() {
  return sqTab_(SQ_HOC, H_SQ_HOC, function (sh) {
    var now = bayGio_();
    sh.getRange(2, 1, SQ_HOC_MAU.length, H_SQ_HOC.length).setValues(SQ_HOC_MAU.map(function (r) { return r.concat(['anh Đạt', now]); }));
  });
}
function sqTabQuen_() { return sqTab_(SQ_QUEN, H_SQ_QUEN); }

/* ---------- bộ nhớ theo mối ---------- */
function sqHoc_() {
  var v = sqTabHoc_().getDataRange().getValues().slice(1), gap = {}, out = [];
  for (var i = v.length - 1; i >= 0; i--) { // dòng mới nhất thắng
    var r = v[i], chu = String(r[1]).trim(), ten = String(r[2]).trim();
    if (!chu || !ten) continue;
    var k = khongDau_(r[0]) + '|' + khongDau_(chu);
    if (gap[k]) continue; gap[k] = 1;
    out.push({ moi: String(r[0]).trim(), chu: chu, ten: ten, sl: soHoacRong_(r[3]), dv: String(r[4]), gia: soHoacRong_(r[5]), tien: soHoacRong_(r[6]), ghi: String(r[7]) });
  }
  return out; // mới nhất trước
}
function sqHocTuSua_(moi, ai, dong, nguoi) {
  var them = [];
  (dong || []).forEach(function (d) {
    if (!d.sua || !d.chu) return;
    var doi = khongDau_(d.ten) !== khongDau_(d.aiTen) || String(d.sl) !== String(d.aiSl) || String(d.gia) !== String(d.aiGia) || String(d.tien) !== String(d.aiTien);
    if (!doi) return;
    them.push([moi || '', d.chu, d.ten, d.sl, d.dv, d.gia, d.tien, '', nguoi, bayGio_()]);
  });
  if (!them.length) return 0;
  var sh = sqTabHoc_();
  sh.getRange(sh.getLastRow() + 1, 1, them.length, H_SQ_HOC.length).setValues(them);
  return them.length;
}

/* ---------- lời nhắc đọc hóa đơn ---------- */
function sqPromptHD_(ds, alias, hoc, moiBiet) {
  var nam = homNay_().slice(0, 4);
  var L = [
    'Bạn là kế toán của quán ăn Gặm ở Sài Gòn. Hôm nay là ' + homNay_() + '. Ảnh là 1 chứng từ MUA HÀNG của quán. Có thể là: hóa đơn viết tay, hóa đơn hoặc biên bản giao hàng in sẵn có ghi tay, hoặc ảnh chụp TIN NHẮN (Zalo) người bán báo hàng và tiền. Nếu có nhiều ảnh thì đó là các trang của CÙNG 1 chứng từ.',
    'Việc của bạn: liệt kê TỪNG món hàng quán đã mua, mỗi món 1 dòng: số lượng (sl), đơn vị (dv), đơn giá (gia), thành tiền (tien). Tiền luôn trả theo ĐỒNG.',
    '',
    'CÁCH ĐỌC SỐ (rất quan trọng):',
    '1. Người bán hay viết tắt tiền theo NGHÌN: "75" = 75.000đ, "375" = 375.000đ, "15k" = 15.000đ, "1.2tr" = 1.200.000đ.',
    '2. "k" hoặc "ký" đứng NGAY SAU SỐ LƯỢNG hàng là KÝ (kg): "chân gà 10k x 95" = 10 kg × 95.000đ. "năm ký" = 5 kg.',
    '3. Phép tính "a x b = c" hoặc "a * b = c": a = số lượng, b = đơn giá, c = thành tiền. THÀNH TIỀN ghi trên chứng từ là chuẩn. Chọn đơn vị tiền của b sao cho a × b = c. Ví dụ "10 bịch * 35 = 350" là 10 bịch × 35.000đ; "40 cái * 15 = 60" là 40 cái × 1.500đ.',
    '4. Nếu số lượng ghi theo đơn vị lớn (vỉ, thùng, bao, két) mà đơn giá lại là giá 1 cái lẻ, thì giữ đơn vị lớn, đơn giá = thành tiền ÷ số lượng, và ghi cách hiểu vào lyDo.',
    '5. Cộng dồn kiểu "...=375+375=750" là TỔNG, không phải món mới. Mỗi phép tính số lượng × giá là 1 món.',
    '6. Trên biểu mẫu in có cột: số tiền ghi tay (dạng 1.084.000 hoặc có dấu chấm ngăn nghìn) là THÀNH TIỀN dù nằm ở cột nào. Số lượng lấy ở chỗ ghi số bình / số cái / số kg thực giao (vd "Số lượng Bên bán giao: 2" → 2 bình). Tên hàng in sẵn thì đọc cả phần ghi tay bổ sung (vd "bình 12" = bình 12kg).',
    '7. Kiểm tra hợp lý: quán mua vài kg đến vài chục kg. Nếu ra hàng trăm, hàng nghìn kg thì chắc chắn đã đọc nhầm tiền thành số lượng – đọc lại.',
    '8. Không có đơn giá thì gia = thành tiền ÷ số lượng. Không có số lượng thì để trống sl, chỉ ghi tien.',
    '',
    'GHI KẾT QUẢ:',
    '- chu = chữ đúng y như trên ảnh cho món đó. ten = tên món chuẩn: nếu khớp món trong DANH MỤC hoặc MẪU ĐÃ HỌC thì ghi đúng tên đó, không thì viết rõ nghĩa.',
    '- cachTinh = phép tính đã hiểu, viết gọn, vd "5 kg × 75.000 = 375.000".',
    '- chacChan=false khi chữ/số mờ, khi phải đoán tên món, hoặc khi phải tự suy đơn vị tiền; lyDo ngắn bằng tiếng Việt.',
    '- anh = số thứ tự ảnh chứa dòng đó (ảnh đầu = 0). vung = [ymin, xmin, ymax, xmax] theo thang 0–1000 bao quanh dòng đó.',
    '- tongTrenAnh = tổng tiền ghi trên ảnh (đồng), không có thì 0. ngay = ngày trên chứng từ yyyy-mm-dd (chỉ ghi ngày/tháng thì năm là ' + nam + '), không có thì rỗng. moi = tên người bán/cửa hàng/công ty nếu thấy, không thì rỗng.',
    '- Phí giao, phí ship là 1 dòng. Không bỏ sót, không thêm dòng không có trên ảnh.',
    '',
    'DANH MỤC món quán hay mua: ' + ds.slice(0, 400).join('; ')
  ];
  if (moiBiet && moiBiet.length) L.push('CÁC MỐI QUEN của quán (nếu chứng từ là của mối nào thì ghi moi đúng tên đó): ' + moiBiet.slice(0, 40).join('; '));
  if (alias && alias.length) L.push('TÊN GỌI KHÁC đã biết (chữ trên hóa đơn → tên đúng): ' + alias.slice(0, 150).map(function (a) { return a.chu + ' → ' + a.ten; }).join('; '));
  if (hoc && hoc.length) {
    L.push('', 'MẪU ĐÃ HỌC (nhân viên đã sửa đúng ở các lần trước – gặp chữ giống vậy thì đọc theo đúng cách này):');
    hoc.slice(0, 50).forEach(function (h) {
      L.push('- Mối «' + (h.moi || '?') + '»: trên ảnh ghi «' + h.chu + '» → món «' + h.ten + '»' + (h.sl !== '' ? ', ' + h.sl + ' ' + h.dv : '') +
        (h.gia !== '' ? ', đơn giá ' + h.gia + 'đ' : '') + (h.tien !== '' ? ', thành tiền ' + h.tien + 'đ' : '') + (h.ghi ? ' (' + h.ghi + ')' : ''));
    });
  }
  return L.join('\n');
}
function sqSchemaHD_() {
  return { type: 'OBJECT', properties: {
    moi: { type: 'STRING' }, ngay: { type: 'STRING' }, tongTrenAnh: { type: 'NUMBER' },
    dong: { type: 'ARRAY', items: { type: 'OBJECT', properties: {
      chu: { type: 'STRING' }, ten: { type: 'STRING' }, sl: { type: 'NUMBER' }, dv: { type: 'STRING' }, gia: { type: 'NUMBER' }, tien: { type: 'NUMBER' },
      chacChan: { type: 'BOOLEAN' }, lyDo: { type: 'STRING' }, cachTinh: { type: 'STRING' }, anh: { type: 'INTEGER' }, vung: { type: 'ARRAY', items: { type: 'INTEGER' } }
    }, required: ['chu', 'ten', 'tien', 'chacChan'] } } }, required: ['dong'] };
}
function sqMoiBiet_(metas) {
  var dem = {};
  (metas || []).forEach(function (m) { if ((m.loai === LOAI.HD || m.loai === LOAI.TD) && m.moi && m.moi !== 'Chưa rõ mối') dem[m.moi] = (dem[m.moi] || 0) + 1; });
  sqQuenRows_().forEach(function (q) { if (q.moi) dem[q.moi] = (dem[q.moi] || 0) + 1; });
  return Object.keys(dem).sort(function (a, b) { return dem[b] - dem[a]; });
}
// Ghép tên mối AI đọc với mối quen; không thấy tên thì đoán theo mẫu đã học
function sqGhepMoi_(aiMoi, dong, moiBiet, hoc) {
  var k = khongDau_(aiMoi);
  if (k) {
    for (var i = 0; i < moiBiet.length; i++) {
      var b = khongDau_(moiBiet[i]);
      if (b === k || (b.length >= 4 && k.indexOf(b) >= 0) || (k.length >= 4 && b.indexOf(k) >= 0)) return moiBiet[i];
    }
    return chu_(aiMoi, 60);
  }
  for (var j = 0; j < (dong || []).length; j++) {
    var c = khongDau_(dong[j].chu);
    if (!c) continue;
    for (var h = 0; h < hoc.length; h++) {
      var hc = khongDau_(hoc[h].chu);
      if (hoc[h].moi && (hc === c || (hc.length >= 6 && (c.indexOf(hc) >= 0 || hc.indexOf(c) >= 0)))) return hoc[h].moi;
    }
  }
  return '';
}

/* ---------- gửi nhiều hóa đơn ---------- */
function sqGuiHD_(nv, p) {
  var ngay = sqKiemNgay_(p.ngay);
  var ds = (p.hd || []).slice(0, 12);
  if (!ds.length) throw new Error('Chưa có ảnh.');
  var folder = sqThuMuc_(), stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss'), now = bayGio_(), ms = Date.now();
  var rows = [], ids = [];
  ds.forEach(function (h, i) {
    var id = sqKiemId_(h.id);
    var fids = (h.anh || []).slice(0, 4).map(function (a, j) {
      return folder.createFile(Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'image/jpeg', 'HD-' + stamp + '-' + (i + 1) + (j ? '-' + (j + 1) : '') + '.jpg')).getId();
    });
    if (!fids.length) return;
    var thumb = /^[A-Za-z0-9+/=]+$/.test(h.thumb || '') && h.thumb.length < 16000 ? h.thumb : '';
    rows.push({ id: id, loai: LOAI.HD, ngay: ngay, nguoi: nv.ten, moi: '', anh: fids, ai: null, dong: [], tongAnh: '', cong: 0, ghiChu: '',
      tt: TT.DOC, taoLuc: now, them: { thumb: thumb, guiLuc: ms }, tong: 0 });
    ids.push(id);
  });
  khoa_(function () {
    var sh = layFile_().getSheetByName(SQ_META);
    var moi = rows.filter(function (o) { return timDong_(sh, o.id) < 0; });
    if (moi.length) sh.getRange(sh.getLastRow() + 1, 1, moi.length, H_SQ_META.length).setValues(moi.map(sqMetaRow_));
    sqMetaMoi_();
  });
  var fidMap = {}; rows.forEach(function (o) { fidMap[o.id] = o.anh; });
  return { ids: ids, fids: fidMap };
}

/* ---------- đọc ở nền ---------- */
function sqDocLai_(id) {
  id = sqKiemId_(id);
  khoa_(function () { var m = sqTimMeta_(id); if (!m) throw new Error('Không tìm thấy hóa đơn.'); if (TT_CHUA_XAC_NHAN.indexOf(m.tt) < 0) throw new Error('Hóa đơn đã xác nhận, không đọc lại được.'); m.tt = TT.DOC; m.them = m.them || {}; m.them.docLuc = 0; delete m.them.loi; sqGhiMeta_(m); });
  return sqDocNen_([id]);
}
function sqDocNen_(ids) {
  var muon = {}; (ids || []).slice(0, 12).forEach(function (x) { muon[String(x)] = 1; });
  var now = Date.now(), viec = [];
  khoa_(function () { // nhận việc: đánh dấu đang đọc để 2 máy không đọc trùng
    sqMetaRows_().forEach(function (m) {
      if (!muon[m.id] || (m.tt !== TT.DOC && m.tt !== TT.LOI)) return;
      var t = (m.them && m.them.docLuc) || 0;
      if (now - t < 150000) return;
      m.them = m.them || {}; m.them.docLuc = now; m.tt = TT.DOC;
      sqGhiMeta_(m); viec.push(m);
    });
  });
  if (!viec.length) return { xong: 0 };
  var props = PropertiesService.getScriptProperties(), key = props.getProperty('GEMINI_KEY');
  if (!key) throw new Error('Chưa cài mã Gemini (GEMINI_KEY).');
  var model = props.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';
  var dm = docDanhMuc_(), tenDm = {}, ds = [];
  dm.forEach(function (d) { var k = khongDau_(d.ten); if (!tenDm[k]) { tenDm[k] = d.ten; ds.push(d.ten); } });
  var alias = sqAlias_(), hoc = sqHoc_(), moiBiet = sqMoiBiet_(sqMetaRows_());
  hoc.forEach(function (h) { if (h.moi && moiBiet.indexOf(h.moi) < 0) moiBiet.push(h.moi); });
  var prompt = sqPromptHD_(ds, alias, hoc, moiBiet), schema = sqSchemaHD_();
  var partsDs = viec.map(function (m) {
    var parts = [{ text: prompt }];
    m.anh.forEach(function (fid) { var b = DriveApp.getFileById(fid).getBlob(); parts.push({ inline_data: { mime_type: b.getContentType() || 'image/jpeg', data: Utilities.base64Encode(b.getBytes()) } }); });
    return parts;
  });
  var reqs = partsDs.map(function (parts) {
    return { url: 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key }, payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: schema, thinkingConfig: { thinkingLevel: 'low' } } }) };
  });
  var res = UrlFetchApp.fetchAll(reqs);
  var ketQua = viec.map(function (m, i) {
    try {
      var body = res[i].getResponseCode() === 200 ? res[i].getContentText() : sqGoiGemini_(key, model, partsDs[i], schema).body; // lỗi/quá tải → thử lại + đổi model
      var j = JSON.parse(body), txt = j.candidates[0].content.parts.map(function (x) { return x.text || ''; }).join('');
      var doc = sqJ_(txt, null);
      if (!doc || !doc.dong) throw new Error('Không đọc được, chụp lại rõ hơn nhé.');
      var moi = sqGhepMoi_(doc.moi, doc.dong, moiBiet, hoc);
      var aliasMoi = alias.filter(function (a) { return !a.moi || !moi || khongDau_(a.moi) === khongDau_(moi); });
      var dong = doc.dong.map(function (d) { var o = sqChuanDong_(d, aliasMoi, tenDm); o.cachTinh = chu_(d.cachTinh, 80); return o; });
      return { ok: true, moi: moi, ngay: /^\d{4}-\d{2}-\d{2}$/.test(doc.ngay || '') ? doc.ngay : '', tongAnh: so_(doc.tongTrenAnh), dong: dong };
    } catch (e) { return { ok: false, loi: String(e && e.message || e).slice(0, 200) }; }
  });
  khoa_(function () {
    var theoId = {}; sqMetaRows_().forEach(function (m) { theoId[m.id] = m; });
    viec.forEach(function (v, i) {
      var m = theoId[v.id], kq = ketQua[i];
      if (!m || m.tt !== TT.DOC) return; // đã bị xóa / đã xác nhận trong lúc đọc
      m.them = m.them || {}; m.them.docLuc = Date.now();
      if (!kq.ok) { m.tt = TT.LOI; m.them.loi = kq.loi; sqGhiMeta_(m); return; }
      delete m.them.loi;
      m.ai = { moi: kq.moi, ngay: kq.ngay, tongAnh: kq.tongAnh, dong: kq.dong, model: model };
      m.dong = kq.dong; if (!m.moi) m.moi = kq.moi; m.tongAnh = kq.tongAnh || '';
      m.cong = kq.dong.reduce(function (s, d) { return s + so_(d.tien); }, 0); m.tong = m.cong; m.tt = TT.KIEM;
      sqGhiMeta_(m);
    });
  });
  return { xong: viec.length, loi: ketQua.filter(function (k) { return !k.ok; }).length };
}

/* ---------- danh sách, xóa ---------- */
function sqHdDs_(nv, p) {
  var den = homNay_(), tu = congNgay_(den, -(Math.min(so_(p.soNgay) || 14, 60)));
  var quen = {}; sqQuenRows_().forEach(function (q) { quen[q.id] = q.ten; });
  var gan = congNgay_(den, -2);
  var ds = sqMetaRows_().filter(function (m) {
    if (m.loai !== LOAI.HD && m.loai !== LOAI.TD) return false;
    return m.ngay >= tu || TT_CHUA_XAC_NHAN.indexOf(m.tt) >= 0;
  }).map(function (m) {
    var t = m.them || {};
    return { id: m.id, loai: m.loai, ngay: m.ngay, nguoi: m.nguoi, moi: m.moi, tong: m.tong || m.cong, tt: m.tt, gio: String(m.taoLuc).slice(11, 16), taoLuc: m.taoLuc,
      thumb: t.thumb || '', soAnh: (m.anh || []).length, soDong: (m.dong || []).length, loi: t.loi || '', kiemBoi: t.kiemBoi || '', guiLuc: t.guiLuc || 0, docLuc: t.docLuc || 0,
      quen: t.quenId ? (quen[t.quenId] || '') : '', vang: (m.dong || []).filter(function (d) { return d.chacChan === false && !d.sua; }).length, hoi: m.hoi,
      ct: (TT_CHUA_XAC_NHAN.indexOf(m.tt) >= 0 || m.tt === TT.HOI || m.ngay >= gan) ? { dong: m.dong, tongAnh: m.tongAnh, ghiChu: m.ghiChu, anh: m.anh, aiNgay: m.ai && m.ai.ngay || '', loi: t.loi || '', traLien: !!t.traLien, quenId: t.quenId || '' } : null };
  });
  ds.sort(function (a, b) { return a.ngay < b.ngay ? 1 : a.ngay > b.ngay ? -1 : (a.guiLuc || 0) < (b.guiLuc || 0) ? 1 : (String(a.taoLuc) < String(b.taoLuc) ? 1 : -1); });
  return { ds: ds, tu: tu, den: den, bayGio: Date.now() };
}
function sqHdXoa_(nv, id) {
  id = sqKiemId_(id);
  return khoa_(function () {
    var m = sqTimMeta_(id);
    if (!m) throw new Error('Không tìm thấy hóa đơn.');
    if (m.tt === TT.OK) throw new Error('Hóa đơn anh Đạt đã duyệt, không xóa được.');
    layFile_().getSheetByName(SQ_META).deleteRow(m._row); sqMetaMoi_();
    if (TT_CHUA_XAC_NHAN.indexOf(m.tt) < 0) {
      var ss = layFile_(), shP = ss.getSheetByName('Phieu'), r = timDong_(shP, id);
      if (r > 0) shP.deleteRow(r);
      xoaChiTiet_(ss, id);
      PropertiesService.getScriptProperties().setProperty('SQ_TONGHOP_CU', '1');
    }
    return { id: id };
  });
}
function sqKiemAnhCuaPhieu_(fid) {
  fid = String(fid || '');
  var co = sqMetaRows_().some(function (m) { return (m.anh || []).indexOf(fid) >= 0; });
  if (!co) throw new Error('Ảnh không thuộc hóa đơn nào.');
}

/* ---------- hóa đơn quen ---------- */
function sqQuenRows_() {
  return sqTabQuen_().getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]) && String(r[1]); }).map(function (r, i) {
    return { id: String(r[0]), ten: String(r[1]), icon: String(r[2]) || '📦', moi: String(r[3]), mon: sqJ_(r[4], []), traLien: String(r[5]) === 'x', nguoi: String(r[6]), suaLuc: String(r[7]), _row: i + 2 };
  });
}
function sqQuenDs_(metas) {
  var lan = {}, homNay = homNay_(), hn = [];
  (metas || []).forEach(function (m) {
    if (m.loai !== LOAI.TD) return;
    var q = m.them && m.them.quenId;
    if (q && (!lan[q] || m.ngay > lan[q])) lan[q] = m.ngay;
    if (m.ngay === homNay) hn.push({ id: m.id, moi: m.moi, nguoi: m.nguoi, tong: m.tong || m.cong, gio: String(m.taoLuc).slice(11, 16), quenId: q || '',
      mon: (m.dong || []).map(function (d) { return d.ten + ' ' + (d.sl === '' ? '' : d.sl) + ' ' + (d.dv || ''); }).join(', ') });
  });
  return { ds: sqQuenRows_().map(function (q) { delete q._row; q.lanCuoi = lan[q.id] || ''; return q; }), homNay: hn.reverse() };
}
function sqIcon_(s) {
  var k = ' ' + khongDau_(s).replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ') + ' ';
  var bang = [['sua chua', '🥛'], ['sua', '🥛'], ['trung', '🥚'], ['muoi', '🧂'], ['cha', '🥟'], ['nem', '🥟'], ['gas', '🔥'], ['binh', '🔥'], ['da', '🧊'], ['bia', '🍺'], ['nuoc', '💧'],
    ['rau', '🥬'], ['ca', '🐟'], ['tom', '🦐'], ['muc', '🦑'], ['ga', '🐔'], ['vit', '🦆'], ['thit', '🥩'], ['heo', '🥩'], ['bo', '🥩'], ['banh', '🥖'], ['bun', '🍜'], ['dau', '🫗'], ['duong', '🍬'], ['hop', '📦']];
  for (var i = 0; i < bang.length; i++) if (k.indexOf(' ' + bang[i][0] + ' ') >= 0) return bang[i][1];
  return '📦';
}
function sqMonQuen_(mon) {
  return sqMon_(mon).map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia }; });
}
function sqQuenLuu_(nv, p) {
  var ten = chu_(p.ten, 40), mon = sqMonQuen_(p.mon);
  if (!ten) throw new Error('Chưa đặt tên hóa đơn quen.');
  if (!mon.length) throw new Error('Hóa đơn quen chưa có món nào.');
  return khoa_(function () {
    var sh = sqTabQuen_(), ds = sqQuenRows_(), id = String(p.id || ''), cu = id ? ds.filter(function (q) { return q.id === id; })[0] : null;
    if (!cu) { var trung = ds.filter(function (q) { return khongDau_(q.ten) === khongDau_(ten); })[0]; if (trung) cu = trung; }
    var r = [cu ? cu.id : 'Q' + Date.now().toString(36).toUpperCase(), ten, chu_(p.icon, 4) || (cu && cu.icon) || sqIcon_(ten + ' ' + mon[0].ten), chu_(p.moi, 60),
      JSON.stringify(mon), p.traLien ? 'x' : '', nv.ten, bayGio_()];
    if (cu) sh.getRange(cu._row, 1, 1, H_SQ_QUEN.length).setValues([r]); else sh.appendRow(r);
    return { quen: { id: r[0], ten: r[1], icon: r[2], moi: r[3], mon: mon, traLien: !!p.traLien } };
  });
}
function sqQuenXoa_(id) {
  return khoa_(function () {
    var q = sqQuenRows_().filter(function (x) { return x.id === String(id || ''); })[0];
    if (!q) throw new Error('Không tìm thấy hóa đơn quen.');
    sqTabQuen_().deleteRow(q._row);
    return { id: q.id };
  });
}
// Sau khi lưu phiếu hàng tự đặt: cập nhật hóa đơn quen đang dùng, hoặc tạo mới nếu nhân viên tick "Lưu làm hóa đơn quen"
function sqQuenSauLuu_(nv, p, moi, dong, traLien) {
  try {
    if (p.quenId && p.capNhatQuen) {
      var q = sqQuenRows_().filter(function (x) { return x.id === String(p.quenId); })[0];
      if (q) { sqTabQuen_().getRange(q._row, 4, 1, 5).setValues([[moi, JSON.stringify(sqMonQuen_(dong)), traLien ? 'x' : '', nv.ten, bayGio_()]]); return q.id; }
    }
    if (p.taoQuen && p.taoQuen.ten) {
      var sh = sqTabQuen_(), ten = chu_(p.taoQuen.ten, 40), id = 'Q' + Date.now().toString(36).toUpperCase();
      sh.appendRow([id, ten, sqIcon_(ten + ' ' + (dong[0] && dong[0].ten)), moi, JSON.stringify(sqMonQuen_(dong)), traLien ? 'x' : '', nv.ten, bayGio_()]);
      return id;
    }
  } catch (e) { }
  return p.quenId ? String(p.quenId) : '';
}
