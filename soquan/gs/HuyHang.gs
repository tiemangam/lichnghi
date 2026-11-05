/**
 * HỦY HÀNG (khấu hao NVL) – phần thêm của Sổ Quán.
 * Nhân viên: chọn món hủy + số lượng + lý do + BẮT BUỘC chụp ảnh. KHÔNG nhập giá.
 * Giá do anh Đạt đặt trong Điều hành (tab SQ_GiaHuy). Món mới nhân viên gõ vào → thêm vào bảng giá, chờ anh đặt giá;
 * khi anh đặt giá, các lần hủy trước đó của món này (đang chưa có giá) tự tính tiền luôn.
 * Đặt giá mới → lưu lịch sử; lần hủy cũ giữ giá cũ.
 */
var SQ_HUY = 'SQ_Huy', SQ_GIAHUY = 'SQ_GiaHuy';
var H_SQ_HUY = ['ID phiếu', 'STT', 'Ngày', 'Ca', 'Người nhập', 'Món', 'Số lượng', 'Đơn vị', 'Đơn giá', 'Thành tiền', 'Lý do', 'Ảnh (file ID)', 'Ghi chú', 'Tạo lúc'];
var HU = { ID: 0, STT: 1, NGAY: 2, CA: 3, NGUOI: 4, MON: 5, SL: 6, DV: 7, GIA: 8, TIEN: 9, LYDO: 10, ANH: 11, GHICHU: 12, TAO: 13 };
var H_SQ_GIAHUY = ['Món', 'Đơn vị', 'Đơn giá', 'Lịch sử giá (JSON)', 'Cập nhật lúc', 'Ghi chú'];
var LY_DO_HUY = ['Nấu / chiên hư', 'Để lâu, quá hạn', 'Rơi đổ', 'Khách trả lại', 'Khác'];

// Bảng giá khấu hao lấy từ Gặm Điều Hành (tháng 7–9/2026) – nạp 1 lần khi tab còn trống
var SQ_GIAHUY_MAU = [
  ['Chả giò cá trích', 'Cuộn', 2000], ['Bò áp chảo thành phẩm', 'Kg', 300000], ['Bò vụn', 'Kg', 250000], ['Bò nhúng', 'Phần', 26000],
  ['Đùi vịt tươi', 'Đùi (300gram)', 27000], ['Đùi vịt rim (1/2 má đùi)', 'Phần', 15000], ['Vịt nấu tiêu thành phẩm', 'Phần', 23000],
  ['Đùi gà tươi', 'Đùi', 8000], ['Đùi gà nhỏ', 'Đùi', 8000], ['Đùi gà mini', 'Đùi', 10000], ['Đùi gà rim', 'Phần', 12000], ['Đùi gà rim khô', 'Đùi', 12000],
  ['Đùi gà tàu xì thành phẩm', 'Đùi', 10000], ['Chân gà chiên thành phẩm', 'Phần', 21000], ['Chân gà thâm', 'Kg', 70000], ['Chân gà thâm đen', 'Kg', 70000],
  ['Chân gà lên men', 'Kg', 100000], ['Chân rút thâm đen', 'Kg', 140000], ['Chân rút còn lông da', 'Kg', 140000], ['Chân rút lên men nhớt', 'Kg', 170000],
  ['Sụn tỏi', 'Phần', 21000], ['Trứng mực Phan Thiết bán thành phẩm', 'Kg', 165000], ['Trứng mực Phan Thiết thành phẩm', 'Phần', 23000],
  ['Trứng mực hoành thánh', 'Phần', 23000], ['Răng mực thành phẩm', 'Phần', 21000], ['Hoành thánh thành phẩm', 'Phần', 23000], ['Vỏ hoành thánh', 'Kg', 15000],
  ['Hến bán thành phẩm', 'Kg', 140000], ['Thịt heo thành phẩm', 'Kg', 100000], ['Cari thành phẩm', 'Phần', 21000], ['Thố giấm thành phẩm', 'Nồi', 15000],
  ['Mì xào rau thành phẩm', 'Phần', 12000], ['Cơm chiên', 'Phần', 19000], ['Cơm nấu thành phẩm', 'Kg', 40000], ['Bún trụng bời rời', 'Kg', 11000],
  ['Mì gói nguyên cây', 'Cây', 55000], ['Nước lẩu', 'Phần', 5000], ['Kho quẹt chấm', 'Phần', 10000], ['Rau củ kho quẹt thành phẩm', 'Phần', 15000],
  ['Gỏi mít', 'Phần', 22000], ['Mít trộn', 'Chén', 5000], ['Xoài cóc thành phẩm', 'gram', 45000], ['Xoài', 'Kg', 25000], ['Cóc', 'Kg', 20000],
  ['Rau sống', 'Kg', 80000], ['Rau răm', 'Kg', 40000], ['Húng lủi', 'Kg', 70000], ['Cải thảo', 'Kg', 15000], ['Cà rốt nguyên củ', 'Củ', 6000],
  ['Bắp hạt', 'Kg', 20000], ['Nấm kim châm', 'Bịch', 15000], ['Nấm kim chi', 'Bịch', 10000], ['Đậu hũ', 'Miếng', 8000], ['Đậu phộng', 'Kg', 80000],
  ['Cốt dừa', 'Lít', 60000], ['Cốt dừa hỏng', 'Kg', 60000], ['Sữa tươi', 'Lít', 65000], ['Sữa chua', 'Hũ', 9000], ['Bánh Flan', 'Cái', 5000],
  ['Nước sâm', 'Chai', 40000], ['Trà tắc', 'Chai', 30000], ['Bột mì', 'Bịch', 25000]
];

function sqHuyCaiDat_(mk) {
  var h = mk(SQ_HUY, H_SQ_HUY);
  var g = mk(SQ_GIAHUY, H_SQ_GIAHUY);
  if (g.getLastRow() < 2) {
    var now = bayGio_();
    g.getRange(2, 1, SQ_GIAHUY_MAU.length, H_SQ_GIAHUY.length).setValues(SQ_GIAHUY_MAU.map(function (x) {
      return [x[0], x[1], x[2], JSON.stringify([{ tu: '2026-09-01', gia: x[2] }]), now, 'Từ Gặm Điều Hành'];
    }));
  }
}

/* ---------- đọc ---------- */
function sqGiaHuy_() {
  var v = layFile_().getSheetByName(SQ_GIAHUY).getDataRange().getValues().slice(1);
  return v.map(function (r, i) {
    return { ten: String(r[0]).trim(), dv: String(r[1]).trim(), gia: soHoacRong_(r[2]), lichSu: sqJ_(r[3], []), capNhat: String(r[4]), ghiChu: String(r[5] || ''), _row: i + 2 };
  }).filter(function (x) { return x.ten; });
}
function sqTimGiaHuy_(ds, ten) { var k = khongDau_(ten); for (var i = 0; i < ds.length; i++) if (khongDau_(ds[i].ten) === k) return ds[i]; return null; }
// giá áp dụng tại 1 ngày (theo lịch sử)
function sqGiaTaiNgay_(item, ngay) {
  if (!item) return '';
  var ls = (item.lichSu || []).slice().sort(function (a, b) { return a.tu < b.tu ? -1 : 1; }), g = '';
  ls.forEach(function (x) { if (x.tu <= ngay) g = x.gia; });
  if (g === '' && ls.length) g = ls[0].gia;
  return g === '' ? item.gia : g;
}
function sqHuyRows_() {
  var sh = layFile_().getSheetByName(SQ_HUY);
  return sh.getDataRange().getValues().slice(1).map(function (r, i) {
    return { id: String(r[HU.ID]), stt: so_(r[HU.STT]), ngay: chuanNgay_(r[HU.NGAY]), ca: String(r[HU.CA]), nguoi: String(r[HU.NGUOI]), ten: String(r[HU.MON]),
      sl: so_(r[HU.SL]), dv: String(r[HU.DV]), gia: soHoacRong_(r[HU.GIA]), tien: soHoacRong_(r[HU.TIEN]), lyDo: String(r[HU.LYDO]),
      anh: String(r[HU.ANH] || '').split(',').filter(String), ghiChu: String(r[HU.GHICHU]), tao: String(r[HU.TAO]), _row: i + 2 };
  }).filter(function (x) { return x.id && x.ten; });
}

/* ---------- cho nhân viên ---------- */
function sqHuyChoNV_(nv) {
  var gia = sqGiaHuy_(), rows = sqHuyRows_();
  var dem = {};
  rows.forEach(function (r) { var k = khongDau_(r.ten); dem[k] = (dem[k] || 0) + 1; });
  var tu = congNgay_(homNay_(), -3);
  return {
    mon: gia.map(function (x) { return { ten: x.ten, dv: x.dv, coGia: x.gia !== '', lan: dem[khongDau_(x.ten)] || 0 }; }),
    lyDo: LY_DO_HUY,
    gan: rows.filter(function (r) { return r.nguoi === nv.ten && r.ngay >= tu; }).slice(-30).reverse()
      .map(function (r) { return { id: r.id, ngay: r.ngay, ca: r.ca, ten: r.ten, sl: r.sl, dv: r.dv, lyDo: r.lyDo }; })
  };
}
function sqHuyLuu_(nv, p) {
  var id = sqKiemId_(p.id), ngay = sqKiemNgay_(p.ngay);
  var ca = chu_(p.ca, 20);
  var anh = (p.anh || []).slice(0, 3);
  if (!anh.length) throw new Error('Phải chụp ảnh đồ hủy.');
  var dong = (p.dong || []).slice(0, 40).map(function (d) {
    return { ten: chu_(d.ten, 80), sl: soHoacRong_(d.sl), dv: chu_(d.dv, 20), lyDo: chu_(d.lyDo, 40), ghiChu: chu_(d.ghiChu, 120) };
  }).filter(function (d) { return d.ten; });
  if (!dong.length) throw new Error('Chưa có món hủy nào.');
  dong.forEach(function (d) {
    if (d.sl === '' || d.sl <= 0) throw new Error('Món "' + d.ten + '" chưa ghi số lượng.');
    if (!d.lyDo) throw new Error('Món "' + d.ten + '" chưa chọn lý do.');
  });
  return khoa_(function () {
    var ss = layFile_();
    var folder = sqThuMuc_(), fids = [];
    anh.forEach(function (a, i) {
      fids.push(folder.createFile(Utilities.newBlob(Utilities.base64Decode(a.b64), a.mime || 'image/jpeg', 'HUY-' + id + '-' + (i + 1) + '.jpg')).getId());
    });
    var gia = sqGiaHuy_(), shG = ss.getSheetByName(SQ_GIAHUY), now = bayGio_(), moi = [];
    var shH = ss.getSheetByName(SQ_HUY);
    // sửa lại phiếu cũ: xóa dòng cũ cùng ID
    var cu = sqHuyRows_().filter(function (r) { return r.id === id; });
    if (cu.length && cu[0].nguoi !== nv.ten) throw new Error('Phiếu này không phải của em.');
    for (var i = cu.length - 1; i >= 0; i--) shH.deleteRow(cu[i]._row);
    var out = dong.map(function (d, i) {
      var it = sqTimGiaHuy_(gia, d.ten);
      if (!it) {
        it = { ten: d.ten, dv: d.dv, gia: '', lichSu: [] };
        shG.appendRow([d.ten, d.dv, '', '[]', now, 'Món mới – ' + nv.ten + ' thêm, chờ anh Đạt đặt giá']);
        gia.push(it); moi.push(d.ten);
      }
      var g = sqGiaTaiNgay_(it, ngay);
      var tien = g === '' ? '' : Math.round(g * d.sl);
      return [id, i + 1, ngay, ca, nv.ten, it.ten, d.sl, d.dv || it.dv, g, tien, d.lyDo, fids.join(','), d.ghiChu, now];
    });
    shH.getRange(shH.getLastRow() + 1, 1, out.length, H_SQ_HUY.length).setValues(out);
    var tong = out.reduce(function (s, r) { return s + (r[HU.TIEN] === '' ? 0 : r[HU.TIEN]); }, 0);
    return { id: id, tong: tong, chuaGia: out.filter(function (r) { return r[HU.TIEN] === ''; }).map(function (r) { return r[HU.MON]; }), monMoi: moi };
  });
}

/* ---------- cho anh Đạt ---------- */
function sqHuyDatGia_(p) {
  var ten = chu_(p.ten, 80), dv = chu_(p.dv, 20), g = soHoacRong_(p.gia);
  if (!ten) throw new Error('Thiếu tên món.');
  if (g === '' || g < 0) throw new Error('Giá không hợp lệ.');
  g = Math.round(g);
  var tu = /^\d{4}-\d{2}-\d{2}$/.test(p.tu || '') ? p.tu : homNay_();
  return khoa_(function () {
    var ss = layFile_(), sh = ss.getSheetByName(SQ_GIAHUY), ds = sqGiaHuy_(), it = sqTimGiaHuy_(ds, ten), now = bayGio_();
    var tenMoi = chu_(p.tenMoi, 80) || (it ? it.ten : ten);
    if (!it) {
      sh.appendRow([tenMoi, dv, g, JSON.stringify([{ tu: '2000-01-01', gia: g }]), now, 'Anh Đạt thêm']);
    } else {
      var ls = it.lichSu || [];
      if (it.gia === '' || !ls.length) ls = [{ tu: '2000-01-01', gia: g }];       // lần đầu đặt giá: áp cho cả các lần hủy trước đó
      else if (g !== it.gia) ls.push({ tu: tu, gia: g });
      sh.getRange(it._row, 1, 1, H_SQ_GIAHUY.length).setValues([[tenMoi, dv || it.dv, g, JSON.stringify(ls), now, it.ghiChu.replace(/,? ?chờ anh Đạt đặt giá/, '')]]);
      it.lichSu = ls; it.gia = g;
    }
    // tính lại tiền các dòng hủy của món này đang chưa có giá (hoặc từ ngày áp giá mới)
    var shH = ss.getSheetByName(SQ_HUY), rows = sqHuyRows_(), item = { gia: g, lichSu: it ? it.lichSu : [{ tu: '2000-01-01', gia: g }] }, n = 0;
    rows.forEach(function (r) {
      if (khongDau_(r.ten) !== khongDau_(ten)) return;
      if (r.gia !== '' && r.ngay < tu) return;
      var gg = sqGiaTaiNgay_(item, r.ngay);
      shH.getRange(r._row, HU.MON + 1, 1, 5).setValues([[tenMoi, r.sl, r.dv || dv, gg, Math.round(gg * r.sl)]]);
      n++;
    });
    return { ten: tenMoi, gia: g, capNhat: n };
  });
}
function sqHuyBaoCao_(p) {
  var den = /^\d{4}-\d{2}-\d{2}$/.test(p.den || '') ? p.den : homNay_();
  var tu = /^\d{4}-\d{2}-\d{2}$/.test(p.tu || '') ? p.tu : den.slice(0, 8) + '01';
  var soNgay = Math.round((Date.parse(den) - Date.parse(tu)) / 864e5) + 1;
  var truocDen = congNgay_(tu, -1), truocTu = congNgay_(tu, -soNgay);
  var rows = sqHuyRows_(), gia = sqGiaHuy_();
  var ky = rows.filter(function (r) { return r.ngay >= tu && r.ngay <= den; });
  var truoc = rows.filter(function (r) { return r.ngay >= truocTu && r.ngay <= truocDen; });
  function gom(ds, khoaFn) {
    var m = {};
    ds.forEach(function (r) {
      var k = khoaFn(r), o = m[k] || (m[k] = { ten: k, sl: 0, dv: r.dv, tien: 0, lan: 0, chuaGia: 0, lyDo: {}, nguoi: {} });
      o.sl += r.sl; o.tien += r.tien === '' ? 0 : r.tien; o.lan++; if (r.tien === '') o.chuaGia++;
      o.lyDo[r.lyDo] = (o.lyDo[r.lyDo] || 0) + 1; o.nguoi[r.nguoi] = (o.nguoi[r.nguoi] || 0) + 1;
    });
    return Object.keys(m).map(function (k) { return m[k]; }).sort(function (a, b) { return b.tien - a.tien || b.lan - a.lan; });
  }
  var theoMonTruoc = {};
  gom(truoc, function (r) { return r.ten; }).forEach(function (x) { theoMonTruoc[khongDau_(x.ten)] = x; });
  var theoMon = gom(ky, function (r) { return r.ten; });
  theoMon.forEach(function (x) { var t = theoMonTruoc[khongDau_(x.ten)]; x.slTruoc = t ? t.sl : 0; x.tienTruoc = t ? t.tien : 0; });
  var theoNgay = {};
  ky.forEach(function (r) { theoNgay[r.ngay] = (theoNgay[r.ngay] || 0) + (r.tien === '' ? 0 : r.tien); });
  var sum = function (ds) { return ds.reduce(function (s, r) { return s + (r.tien === '' ? 0 : r.tien); }, 0); };
  return {
    tu: tu, den: den, truocTu: truocTu, truocDen: truocDen,
    tong: sum(ky), tongTruoc: sum(truoc), soLan: ky.length,
    theoMon: theoMon, theoLyDo: gom(ky, function (r) { return r.lyDo || 'Khác'; }), theoNguoi: gom(ky, function (r) { return r.nguoi || '?'; }),
    theoNgay: theoNgay,
    dong: ky.slice().reverse().slice(0, 300).map(function (r) { return { id: r.id, ngay: r.ngay, ca: r.ca, nguoi: r.nguoi, ten: r.ten, sl: r.sl, dv: r.dv, gia: r.gia, tien: r.tien, lyDo: r.lyDo, anh: r.anh, ghiChu: r.ghiChu }; }),
    bangGia: gia.map(function (x) { return { ten: x.ten, dv: x.dv, gia: x.gia, lichSu: x.lichSu, ghiChu: x.ghiChu }; })
      .sort(function (a, b) { return (a.gia === '' ? 0 : 1) - (b.gia === '' ? 0 : 1) || a.ten.localeCompare(b.ten, 'vi'); })
  };
}
