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
var TT = { CHO: 'Chờ duyệt', OK: 'Đã duyệt', HOI: 'Hỏi lại', TRALOI: 'Đã trả lời', TUCHOI: 'Từ chối' };
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
  chu_huy_gia: function (p) { sqChu_(p.pass); return sqHuyDatGia_(p); }
};

/* ======================= ĐĂNG NHẬP ======================= */
function sqNV_(pin) {
  pin = String(pin || '').trim();
  if (!/^\d{4,6}$/.test(pin)) throw new Error('Sai mã PIN.');
  var rows = layFile_().getSheetByName(SQ_NV).getDataRange().getValues();
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
function sqMetaRows_() {
  var sh = layFile_().getSheetByName(SQ_META);
  var v = sh.getDataRange().getValues();
  return v.slice(1).map(function (r, i) { var o = sqMetaObj_(r); o._row = i + 2; return o; }).filter(function (o) { return o.id; });
}
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
function sqDuLieuNV_(nv) {
  var dm = docDanhMuc_();
  var gop = {};
  dm.forEach(function (d) {
    var k = khongDau_(d.ten);
    if (!gop[k] || d.ngay > gop[k].ngay) gop[k] = { ten: d.ten, dv: d.dv, gia: d.gia, ngay: d.ngay, nhom: d.nhom, lan: d.lan };
  });
  var metas = sqMetaRows_();
  var moiSet = {};
  metas.forEach(function (m) { if (m.moi) moiSet[m.moi] = (moiSet[m.moi] || 0) + 1; });
  var tu = congNgay_(homNay_(), -3);
  var cuaToi = metas.filter(function (m) { return m.nguoi === nv.ten && (m.ngay >= tu || m.tt === TT.HOI); })
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
    danhMuc: Object.keys(gop).map(function (k) { return gop[k]; }),
    tuDat: dm.filter(function (d) { return d.nhom === NHOM_TU_DAT; }).sort(function (a, b) { return b.lan - a.lan; }).slice(0, 24),
    alias: sqAlias_(),
    moi: Object.keys(moiSet).sort(function (a, b) { return moiSet[b] - moiSet[a]; }).slice(0, 20),
    cuaToi: cuaToi,
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
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({ contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: schema } })
  });
  var code = res.getResponseCode(), body = res.getContentText();
  if (code !== 200) throw new Error('Gemini báo lỗi ' + code + ': ' + body.slice(0, 300));
  var j = JSON.parse(body), txt = '';
  try { txt = j.candidates[0].content.parts.map(function (x) { return x.text || ''; }).join(''); } catch (e) { throw new Error('Gemini không trả kết quả.'); }
  var doc = sqJ_(txt, null);
  if (!doc || !doc.dong) throw new Error('Không đọc được hóa đơn, chụp lại rõ hơn nhé.');
  // 3. áp tên gọi khác + đối chiếu danh mục
  doc.dong = (doc.dong || []).map(function (d) { return sqChuanDong_(d, alias, tenDm); });
  if (!moi && doc.moi) moi = doc.moi;
  return { fids: fids, moi: moi, ngay: /^\d{4}-\d{2}-\d{2}$/.test(doc.ngay || '') ? doc.ngay : '', tongAnh: so_(doc.tongTrenAnh), dong: doc.dong, model: model };
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
  if (cu && cu.nguoi && cu.nguoi !== nv.ten) throw new Error('Phiếu này không phải của em.');
  if (cu && cu.tt === TT.OK) throw new Error('Phiếu đã được anh duyệt, không sửa được nữa.');
  return cu ? (cu.tt === TT.HOI ? TT.TRALOI : cu.tt === TT.TUCHOI ? TT.TRALOI : TT.CHO) : TT.CHO;
}

function sqLuuHD_(nv, p) {
  var id = sqKiemId_(p.id), ngay = sqKiemNgay_(p.ngay);
  var moi = chu_(p.moi, 60) || 'Chưa rõ mối';
  var dong = sqMon_(p.dong);
  if (!dong.length) throw new Error('Hóa đơn chưa có dòng nào.');
  return khoa_(function () {
    var cu = sqTimMeta_(id), tt = sqGiuDuyet_(cu, nv);
    var cong = dong.reduce(function (s, d) { return s + d.tien; }, 0);
    apiLuu({ id: id, ngay: ngay, nhom: 'HĐ ' + moi, nguoiMua: nv.ten, ghiChu: chu_(p.ghiChu, 200),
      mon: dong.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) });
    var o = {
      id: id, loai: LOAI.HD, ngay: ngay, nguoi: nv.ten, moi: moi, anh: (p.fids || []).filter(function (x) { return /^[\w-]{10,}$/.test(x); }),
      ai: p.ai || (cu && cu.ai) || null, dong: dong, tongAnh: soHoacRong_(p.tongAnh), cong: cong, ghiChu: chu_(p.ghiChu, 200),
      tt: tt, hoi: cu ? cu.hoi : '', traLoi: cu ? cu.traLoi : '', taoLuc: cu ? cu.taoLuc : bayGio_(), them: { suaLuc: bayGio_() }, tong: cong
    };
    sqGhiMeta_(o);
    sqNhoAlias_(moi, p.nho, nv.ten);
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
    var ghiChu = [moi, traLien ? 'Trả liền' : 'Ghi nợ mối', chu_(p.ghiChu, 150)].filter(String).join(' · ');
    apiLuu({ id: id, ngay: ngay, nhom: NHOM_TU_DAT, nguoiMua: nv.ten, ghiChu: ghiChu,
      mon: dong.map(function (d) { return { ten: d.ten, sl: d.sl, dv: d.dv, gia: d.gia, tien: d.tien }; }) });
    sqGhiMeta_({ id: id, loai: LOAI.TD, ngay: ngay, nguoi: nv.ten, moi: moi, dong: dong, cong: cong, ghiChu: chu_(p.ghiChu, 150),
      tt: tt, hoi: cu ? cu.hoi : '', traLoi: cu ? cu.traLoi : '', taoLuc: cu ? cu.taoLuc : bayGio_(), them: { traLien: traLien, suaLuc: bayGio_() }, tong: cong });
    return { id: id, tong: cong, tt: tt };
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
      var g = { khoa: khoa, ten: cur.ten, nhomHang: sqNhomHang_(cur.ten), dv: cur.dv, id: cur.id, ngay: cur.ngay, nhom: cur.nhom, sl: cur.sl, gia: cur.gia, tien: cur.tien,
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
