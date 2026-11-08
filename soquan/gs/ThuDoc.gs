// TẠM THỜI – thử cách đọc hóa đơn trên ảnh thật trong thư mục "Sổ Quán – ảnh hóa đơn". Xóa tệp này sau khi thử xong.
var THU_ANH = {
  'Zalo cá lóc': '1bHwA9HFHdtjDscH-6_5dhR56sZqalwgo',
  'Gas Petrolimex': '1MrNbHyMDuLbcmxOn07fRb4F30vYi1CeV',
  'Zalo trứng': '1ujtF6WkMcSOuhYcDvL6HcAFhlphbbJac'
};
function thuPromptCu_(ds) {
  return [
    'Đây là 1 ảnh của MỘT hóa đơn mua hàng của quán ăn Việt Nam (có thể viết tay, có thể là hóa đơn in, có thể chụp nghiêng).',
    'Hãy đọc TỪNG dòng hàng hóa và trả về JSON đúng schema.',
    'Quy ước:',
    '- Số tiền trên hóa đơn viết tay thường ghi theo NGHÌN đồng (vd "950" = 950.000đ, "1.2tr" = 1.200.000đ). Luôn trả gia và tien theo ĐỒNG.',
    '- sl là số lượng (số thập phân dùng dấu chấm), dv là đơn vị (kg, con, bó, quả, hũ, gói, thùng, lít…).',
    '- chu = chữ đúng y như viết trên hóa đơn. ten = tên món chuẩn: nếu khớp món trong DANH MỤC bên dưới thì ghi đúng tên đó, không thì viết lại cho rõ nghĩa.',
    '- Nếu chữ hoặc số không chắc chắn (mờ, khó đọc, sl×gia không ra tien) thì chacChan=false và lyDo ngắn bằng tiếng Việt. Ngược lại chacChan=true, lyDo rỗng.',
    '- Phí giao hàng, phí ship cũng là 1 dòng. Không bỏ sót dòng nào.',
    '- tongTrenAnh = tổng tiền ghi trên hóa đơn (đồng), không có thì 0. ngay = ngày trên hóa đơn dạng yyyy-mm-dd, không có thì rỗng. moi = tên người bán / cửa hàng nếu có.',
    'DANH MỤC món quán hay mua: ' + ds.join('; ')
  ].join('\n');
}
function thuPromptMoi_(ds) {
  return [
    'Bạn là kế toán của quán ăn Gặm ở Sài Gòn. Ảnh là 1 chứng từ MUA HÀNG của quán. Có thể là: hóa đơn viết tay, hóa đơn hoặc biên bản giao hàng in sẵn có ghi tay, hoặc ảnh chụp TIN NHẮN (Zalo) người bán báo hàng và tiền.',
    'Việc của bạn: liệt kê TỪNG món hàng quán đã mua, mỗi món 1 dòng: số lượng (sl), đơn vị (dv), đơn giá (gia), thành tiền (tien). Tiền luôn trả theo ĐỒNG.',
    '',
    'CÁCH ĐỌC SỐ (rất quan trọng):',
    '1. Người bán hay viết tắt tiền theo NGHÌN: "75" = 75.000đ, "375" = 375.000đ, "15k" = 15.000đ, "1.2tr" = 1.200.000đ.',
    '2. "k" hoặc "ký" đứng NGAY SAU SỐ LƯỢNG hàng là KÝ (kg): "chân gà 10k x 95" = 10 kg × 95.000đ. "năm ký" = 5 kg.',
    '3. Phép tính "a x b = c" hoặc "a * b = c": a = số lượng, b = đơn giá, c = thành tiền. THÀNH TIỀN ghi trên chứng từ là chuẩn. Chọn đơn vị tiền của b sao cho a × b = c. Ví dụ "10 bịch * 35 = 350" là 10 bịch × 35.000đ; "40 cái * 15 = 60" là 40 cái × 1.500đ.',
    '4. Nếu số lượng ghi theo đơn vị lớn (vỉ, thùng, bao, két) mà đơn giá lại là giá 1 cái lẻ, thì vẫn giữ đơn vị lớn, đơn giá = thành tiền ÷ số lượng, và ghi cách hiểu vào lyDo.',
    '5. Cộng dồn kiểu "...=375+375=750" là TỔNG, không phải món mới. Mỗi phép tính số lượng × giá là 1 món.',
    '6. Trên biểu mẫu in có cột: số tiền ghi tay (dạng 1.084.000 hoặc có dấu chấm ngăn nghìn) là THÀNH TIỀN dù nằm ở cột nào. Số lượng lấy ở chỗ ghi số bình / số cái / số kg thực giao (vd mục "Số lượng Bên bán giao: 2" → 2 bình). Tên hàng in sẵn thì đọc cả phần ghi tay bổ sung (vd "bình 12" = bình 12kg).',
    '7. Kiểm tra hợp lý: quán mua vài kg đến vài chục kg. Nếu ra hàng trăm, hàng nghìn kg thì chắc chắn đã đọc nhầm tiền thành số lượng – đọc lại.',
    '8. Không có đơn giá thì gia = thành tiền ÷ số lượng. Không có số lượng thì để trống sl, chỉ ghi tien.',
    '',
    'GHI KẾT QUẢ:',
    '- chu = chữ đúng y như trên ảnh cho món đó. ten = tên món chuẩn: nếu khớp món trong DANH MỤC thì ghi đúng tên trong danh mục, không thì viết rõ nghĩa (vd "Hông" trong tin nhắn bán cá → "Cá hồng?" và chacChan=false).',
    '- cachTinh = phép tính đã hiểu, viết gọn, vd "5 kg × 75.000 = 375.000".',
    '- chacChan=false khi chữ/số mờ, khi phải đoán tên món, hoặc khi phải tự suy đơn vị tiền; lyDo ngắn bằng tiếng Việt.',
    '- tongTrenAnh = tổng tiền ghi trên ảnh (đồng), không có thì 0. ngay = ngày trên chứng từ yyyy-mm-dd (tin nhắn ghi "2/11" thì năm là năm nay), không có thì rỗng. moi = tên người bán/cửa hàng/công ty nếu thấy, không thì rỗng.',
    '- Phí giao, phí ship là 1 dòng. Không bỏ sót, không thêm dòng không có trên ảnh.',
    '',
    'DANH MỤC món quán hay mua: ' + ds.join('; ')
  ].join('\n');
}
function thuGoi_(key, model, prompt, b64, thinking, coCachTinh) {
  var it = { chu: { type: 'STRING' }, ten: { type: 'STRING' }, sl: { type: 'NUMBER' }, dv: { type: 'STRING' }, gia: { type: 'NUMBER' }, tien: { type: 'NUMBER' }, chacChan: { type: 'BOOLEAN' }, lyDo: { type: 'STRING' } };
  if (coCachTinh) it.cachTinh = { type: 'STRING' };
  var schema = { type: 'OBJECT', properties: { moi: { type: 'STRING' }, ngay: { type: 'STRING' }, tongTrenAnh: { type: 'NUMBER' },
    dong: { type: 'ARRAY', items: { type: 'OBJECT', properties: it, required: ['chu', 'ten', 'tien', 'chacChan'] } } }, required: ['dong'] };
  var gc = { temperature: 0, responseMimeType: 'application/json', responseSchema: schema, thinkingConfig: { thinkingLevel: thinking } };
  var t = Date.now();
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: b64 } }] }], generationConfig: gc })
  });
  var ms = Date.now() - t, code = res.getResponseCode();
  if (code !== 200) return code + ' ' + res.getContentText().slice(0, 120) + ' (' + ms + 'ms)';
  var j = JSON.parse(res.getContentText()), txt = j.candidates[0].content.parts.map(function (x) { return x.text || ''; }).join('');
  var d = JSON.parse(txt);
  return (ms / 1000).toFixed(1) + 's | mối: ' + (d.moi || '-') + ' | ngày: ' + (d.ngay || '-') + ' | tổng ảnh: ' + d.tongTrenAnh + '\n' + (d.dong || []).map(function (x) {
    return '   • ' + x.ten + ' [' + x.chu + '] ' + x.sl + ' ' + (x.dv || '') + ' × ' + x.gia + ' = ' + x.tien + (x.chacChan ? '' : ' ⚠ ' + x.lyDo) + (x.cachTinh ? ' {' + x.cachTinh + '}' : '');
  }).join('\n');
}
function thuDocMau() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  var ds = [], seen = {};
  docDanhMuc_().forEach(function (d) { var k = khongDau_(d.ten); if (!seen[k]) { seen[k] = 1; ds.push(d.ten); } });
  ds = ds.slice(0, 400);
  var cu = thuPromptCu_(ds), moi = thuPromptMoi_(ds);
  Object.keys(THU_ANH).forEach(function (ten) {
    var b64 = Utilities.base64Encode(DriveApp.getFileById(THU_ANH[ten]).getBlob().getBytes());
    Logger.log('===== ' + ten + ' =====');
    Logger.log('A. Cách cũ (3.6-flash, nghĩ ít): ' + thuGoi_(key, 'gemini-3.6-flash', cu, b64, 'low', false));
    Logger.log('B. Cách mới (3.6-flash, nghĩ ít): ' + thuGoi_(key, 'gemini-3.6-flash', moi, b64, 'low', true));
    Logger.log('C. Cách mới (3.6-flash, nghĩ kỹ): ' + thuGoi_(key, 'gemini-3.6-flash', moi, b64, 'high', true));
  });
}
