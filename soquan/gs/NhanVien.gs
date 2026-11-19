/**
 * NHÂN VIÊN & MẬT KHẨU – anh Đạt tự quản lý trong app (Điều hành → Nhân viên).
 * Mỗi nhân viên: tên, mã PIN 4 số (không trùng), vai trò, các chức năng được dùng trong Sổ Quán, nghỉ làm.
 * Mật khẩu của anh Đạt: tạo lần đầu bằng mã quản lý (QL_KEY), sau đó đổi trong app.
 */
var SQ_CHUC_NANG = { hd: 'Chụp hóa đơn', dicho: 'Đi chợ', td: 'Hóa đơn quen', tn: 'Thu ngân cuối ca', sapo: 'Gửi số Sapo', huy: 'Hủy hàng', kho: 'Xem kho (cần đặt, tủ bếp, kiểm tủ)', kiemchi: 'Kiểm chi phí ngày (dò hóa đơn, đi chợ, chi tại quầy)' };
var SQ_VAI_MAC_DINH = {
  'ke toan': ['hd', 'dicho', 'td', 'huy', 'kiemchi'], 'di cho': ['dicho', 'hd', 'td', 'huy'], 'thu ngan': ['tn', 'sapo', 'hd', 'td', 'huy'],
  'quan ly': ['tn', 'sapo', 'kho', 'hd', 'dicho', 'td', 'huy'], 'bep': ['huy', 'hd', 'td', 'dicho'], 'phuc vu': ['huy']
};
function sqTachCN_(s, vai) {
  var ds = String(s || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return SQ_CHUC_NANG[x]; });
  ds = ds.length ? ds : (SQ_VAI_MAC_DINH[khongDau_(vai)] || ['hd', 'dicho', 'td', 'huy']).slice();
  if (khongDau_(vai) === 'thu ngan' && ds.indexOf('tn') < 0) ds.unshift('tn'); // thu ngân luôn có ô Thu ngân cuối ca
  if (/^(thu ngan|quan ly)$/.test(khongDau_(vai)) && ds.indexOf('sapo') < 0) ds.push('sapo'); // thu ngân + quản lý gửi số Sapo mỗi tối
  return ds;
}
function sqTaoPass_(p) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('CHU_PASS')) throw new Error('Đã có mật khẩu rồi. Muốn đổi thì vào Điều hành → Nhân viên.');
  var ql = props.getProperty('QL_KEY');
  if (!ql || String(p.qlKey || '').trim() !== ql) throw new Error('Mã quản lý không đúng.');
  var moi = String(p.moi || '');
  if (moi.length < 6) throw new Error('Mật khẩu ít nhất 6 ký tự.');
  if (moi === ql) throw new Error('Mật khẩu phải khác mã quản lý.');
  props.setProperty('CHU_PASS', moi);
  return { qlKey: ql };
}
function sqDoiPass_(p) {
  var moi = String(p.moi || ''), props = PropertiesService.getScriptProperties();
  if (moi.length < 6) throw new Error('Mật khẩu ít nhất 6 ký tự.');
  if (moi === props.getProperty('QL_KEY')) throw new Error('Mật khẩu phải khác mã quản lý.');
  props.setProperty('CHU_PASS', moi);
  return {};
}
function sqNgaySinh_(v) { var s = (v instanceof Date) ? chuanNgay_(v) : chuanNgay_(String(v || '').trim()); return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : ''; }
function sqNVDs_() {
  var v = layFile_().getSheetByName(SQ_NV).getDataRange().getValues().slice(1);
  return {
    chucNang: SQ_CHUC_NANG, vaiMacDinh: SQ_VAI_MAC_DINH,
    ds: v.map(function (r, i) { return { row: i + 2, ten: String(r[0]).trim(), pin: String(r[1]).trim(), vai: String(r[2]).trim(), nghi: /x/i.test(String(r[3])), chucNang: sqTachCN_(r[4], r[2]), suaLuc: String(r[5] || ''), sdt: String(r[6] || '').trim(), sinh: sqNgaySinh_(r[7]) }; })
      .filter(function (x) { return x.ten; })
  };
}
function sqNVLuu_(p) {
  var ten = chu_(p.ten, 40), pin = String(p.pin || '').trim(), vai = chu_(p.vai, 20);
  if (!ten) throw new Error('Chưa ghi tên.');
  if (!/^\d{4}$/.test(pin)) throw new Error('Mã PIN phải đúng 4 số.');
  var cn = (p.chucNang || []).filter(function (x) { return SQ_CHUC_NANG[x]; });
  if (!cn.length) throw new Error('Chọn ít nhất 1 chức năng cho ' + ten + '.');
  var sdt = String(p.sdt || '').replace(/[^\d+ .\-]/g, '').trim().slice(0, 16), sinh = String(p.sinh || '').trim();
  if (sinh && !/^\d{4}-\d{2}-\d{2}$/.test(sinh)) throw new Error('Ngày sinh chưa đúng (cần chọn đủ ngày/tháng/năm).');
  return khoa_(function () {
    var sh = layFile_().getSheetByName(SQ_NV), ds = sqNVDs_().ds, row = so_(p.row);
    ds.forEach(function (x) {
      if (x.row === row) return;
      if (x.pin === pin && !x.nghi) throw new Error('Mã PIN ' + pin + ' đang là của ' + x.ten + ', chọn số khác.');
      if (khongDau_(x.ten) === khongDau_(ten) && !x.nghi) throw new Error('Đã có nhân viên tên ' + x.ten + '.');
    });
    if (sh.getRange(1, 7).getValue() !== 'Số điện thoại') { sh.getRange(1, 1, 1, H_SQ_NV.length).setValues([H_SQ_NV]); sh.getRange('G:H').setNumberFormat('@'); }
    var vals = [[ten, pin, vai, p.nghi ? 'x' : '', cn.join(','), bayGio_(), sdt, sinh]];
    if (row >= 2 && ds.some(function (x) { return x.row === row; })) sh.getRange(row, 1, 1, H_SQ_NV.length).setValues(vals);
    else { row = sh.getLastRow() + 1; sh.getRange(row, 1, 1, H_SQ_NV.length).setValues(vals); }
    sqCacheXoa_('sq_nv');
    return { row: row };
  });
}

// sinh nhật nhân viên: báo ở Tổng quan từ 3 ngày trước đến đúng ngày
function sqSinhNhatViec_() {
  var hn = homNay_(), y = +hn.slice(0, 4), out = [];
  sqNVDs_().ds.forEach(function (x) {
    if (x.nghi || !x.sinh) return;
    var m = +x.sinh.slice(5, 7), d = +x.sinh.slice(8, 10), pad = function (n) { return ('0' + n).slice(-2); };
    var ngay = y + '-' + pad(m) + '-' + pad(d), con = sqKhoSoNgay_(hn, ngay);
    if (con < 0) { ngay = (y + 1) + '-' + pad(m) + '-' + pad(d); con = sqKhoSoNgay_(hn, ngay); }
    if (con < 0 || con > 3) return;
    var tuoi = +ngay.slice(0, 4) - (+x.sinh.slice(0, 4)), dm = pad(d) + '/' + pad(m);
    out.push({ muc: 'i', loai: '🎂 Sinh nhật', tab: 'nv', ngay: ngay, con: con, text: (con === 0 ? 'Hôm nay là sinh nhật ' : con === 1 ? 'Ngày mai sinh nhật ' : 'Còn ' + con + ' ngày nữa là sinh nhật ') + x.ten + ' (' + dm + (tuoi > 0 && tuoi < 80 ? ', tròn ' + tuoi + ' tuổi' : '') + ')' + (x.sdt ? ' · ' + x.sdt : '') });
  });
  return out.sort(function (a, b) { return a.con - b.con; });
}
