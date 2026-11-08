// TẠM THỜI – thử luồng hóa đơn bản 2 trên ảnh thật (gửi → đọc nền → xem kết quả → xóa). Xóa tệp này sau khi thử xong.
function thuV2() {
  var anh = { 'Gas': '1MrNbHyMDuLbcmxOn07fRb4F30vYi1CeV', 'Cá': '1bHwA9HFHdtjDscH-6_5dhR56sZqalwgo', 'Trứng': '1ujtF6WkMcSOuhYcDvL6HcAFhlphbbJac' };
  var nv = { ten: 'Thử (Claude)' }, hd = [], t0 = Date.now();
  Object.keys(anh).forEach(function (k, i) { hd.push({ id: 'HDTHU' + Date.now().toString(36).toUpperCase() + i, anh: [{ b64: Utilities.base64Encode(DriveApp.getFileById(anh[k]).getBlob().getBytes()), mime: 'image/jpeg' }] }); });
  var g = sqGuiHD_(nv, { ngay: homNay_(), hd: hd });
  Logger.log('Gửi ' + g.ids.length + ' hóa đơn: ' + (Date.now() - t0) + 'ms');
  t0 = Date.now();
  var d = sqDocNen_(g.ids);
  Logger.log('Đọc nền song song: ' + (Date.now() - t0) + 'ms, xong ' + d.xong + ', lỗi ' + d.loi);
  g.ids.forEach(function (id) {
    var m = sqTimMeta_(id);
    Logger.log(m.tt + ' | mối: ' + m.moi + ' | tổng ảnh ' + m.tongAnh + ' | cộng ' + m.cong + (m.them.loi ? ' | LỖI ' + m.them.loi : '') + '\n' + (m.dong || []).map(function (x) {
      return '   • ' + x.ten + ' [' + x.chu + '] ' + x.sl + ' ' + x.dv + ' × ' + x.gia + ' = ' + x.tien + (x.chacChan ? '' : ' ⚠ ' + x.lyDo) + (x.cachTinh ? ' {' + x.cachTinh + '}' : '');
    }).join('\n'));
    sqHdXoa_(nv, id);
  });
  Logger.log('Đã xóa các hóa đơn thử.');
}
