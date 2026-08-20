/**
 * TKSX - Theo dõi sản xuất (Thổi / In)
 * Dán toàn bộ code này vào Apps Script của Google Sheet sản xuất.
 *
 * CÁCH DÙNG:
 * 1) Chạy 1 lần hàm setupSheets()  -> tự tạo 7 tab: data, Config, Master, CN_In_Master, CN_Thoi_Master, May_In_Master, May_Thoi_Master
 * 2) Danh sách CN/Máy — chọn 1 trong 2 cách:
 *    a) Local: điền Mã (cột A) và Tên (cột B) vào 4 tab CN/Máy như cũ.
 *    b) Dùng chung sheet Master TRUNG TÂM: trong Config điền MASTER_URL = link sheet Master
 *       (có tab MA_CN, MA_MAY với cột Mã/Tên/Work station/Apply from/Apply to).
 *       WS_THOI = giá trị Work station của tổ Thổi (mặc định THO), WS_IN = của tổ In (mặc định INT).
 *       Apply from/to (dd/mm/yy): trống = hiệu lực vô hạn; chỉ lấy dòng còn hiệu lực hôm nay.
 *       Lưu ý: tài khoản Google chạy app này phải có quyền XEM sheet Master.
 *       Xóa trống MASTER_URL là quay về cách (a) ngay lập tức.
 * 3) Vào tab Config, dòng 2: đặt PIN (mặc định 1234).
 * 4) Deploy > New deployment > Web app:
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Copy URL kết thúc bằng /exec -> dán vào app (nút ⚙️).
 */

// ====== Tạo sẵn các tab (chạy 1 lần trong trình soạn thảo) ======
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {
    'data':           ['Thời gian', 'Loại', 'Công nhân', 'Máy', 'DATA', 'LSX', 'Khối lượng yêu cầu', 'Phương thức', 'Khối lượng cuộn', 'Kết quả sản xuất', 'Gửi lúc', 'Cảnh báo', 'Ngày SX', 'Ca', 'Mã CN', 'Mã máy'],
    'Config':         ['Key', 'Value'],
    'Master':         ['Loại', 'LSX', 'KLYC', 'Tổng SX', 'Còn lại', 'Số lần SX', 'LSX gốc', 'Loại bù'],
    'Phe':            ['Thời gian', 'Loại', 'Công nhân', 'Máy', 'Phế trơn', 'Phế màu', 'Gửi lúc', 'Ngày SX', 'Ca', 'Mã CN', 'Mã máy'],
    'ChuyenCuon':     ['Thời gian', 'LSX', 'KL chuyển (kg)', 'Gửi lúc'],
    'CN_In_Master':   ['Mã', 'Tên'],
    'CN_Thoi_Master': ['Mã', 'Tên'],
    'May_In_Master':  ['Mã', 'Tên'],
    'May_Thoi_Master':['Mã', 'Tên']
  };
  Object.keys(defs).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(defs[name]);
      sh.getRange(1, 1, 1, defs[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // Đặt PIN mặc định nếu Config trống
  var cfgSh = ss.getSheetByName('Config');
  if (cfgSh.getLastRow() < 2) {
    cfgSh.appendRow(['PIN', '1234']);
  }
  // Đặt PIN chuyển cuộn nếu chưa có
  var cfgVals = cfgSh.getRange(2, 1, cfgSh.getLastRow() - 1, 2).getValues();
  var hasPinChuyen = cfgVals.some(function (r) { return ('' + r[0]).trim().toUpperCase() === 'PIN_CHUYEN'; });
  if (!hasPinChuyen) {
    cfgSh.appendRow(['PIN_CHUYEN', '1111']);
  }
  // Thêm cấu hình Sheet đơn hàng (tra cứu khách hàng) nếu chưa có
  var donHangKeys = [
    ['DONHANG_URL', ''],       // Link Google Sheet đơn hàng (dán cả URL)
    ['DONHANG_TAB', ''],       // Tên tab chứa đơn hàng, VD: DonHang
    ['DONHANG_COL_LSX', ''],   // Cột chứa mã LSX, VD: B
    ['DONHANG_COL_KH', ''],    // Cột chứa tên khách hàng, VD: D
    // --- Sheet Master trung tâm (dùng chung CN/Máy với app thành phẩm) ---
    ['MASTER_URL', ''],        // Link Google Sheet Master trung tâm. ĐỂ TRỐNG = dùng 4 tab local như cũ
    ['MASTER_TAB_CN', 'MA_CN'],   // Tên tab công nhân trong sheet Master
    ['MASTER_TAB_MAY', 'MA_MAY'], // Tên tab máy trong sheet Master
    ['WS_THOI', 'THO'],        // Giá trị cột Work station của tổ Thổi
    ['WS_IN', 'INT']           // Giá trị cột Work station của tổ In
  ];
  var cfgAll = cfgSh.getRange(2, 1, Math.max(cfgSh.getLastRow() - 1, 1), 2).getValues();
  donHangKeys.forEach(function (dk) {
    var exists = cfgAll.some(function (r) { return ('' + r[0]).trim().toUpperCase() === dk[0]; });
    if (!exists) cfgSh.appendRow(dk);
  });
  // Xóa tab mặc định "Trang tính1"/"Sheet1" nếu còn trống
  var def = ss.getSheetByName('Trang tính1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  SpreadsheetApp.getActiveSpreadsheet().toast('Đã tạo xong các tab. Kiểm tra PIN / PIN_CHUYEN trong tab Config.');
}

// ====== Menu tùy chỉnh trong Google Sheet ======
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏭 TKSX')
    .addItem('🔄 Cập nhật Master', 'updateMaster_')
    .addItem('🔧 Cài đặt onChange trigger', 'installOnChangeTrigger')
    .addItem('🌙 Cài lịch tự tính lại Master mỗi đêm', 'installNightlyRebuild')
    .addItem('📋 Kiểm tra lịch đang cài', 'KIEM_TRA_LICH')
    .addItem('📅 Cập nhật cột Ngày SX + Ca', 'backfillNgayCa')
    .addSeparator()
    .addItem('⏱️ ĐO TỐC ĐỘ (chẩn đoán chậm)', 'DO_TOC_DO')
    .addItem('🔍 Kiểm tra bản code đang chạy', 'KIEM_TRA_PHIEN_BAN')
    .addItem('⏹️ Bật/tắt ghi nhật ký tốc độ', 'toggleTocDoLog')
    .addToUi();
}

// ============================================================================
// ⏱️  ĐO TỐC ĐỘ — chạy trực tiếp, KHÔNG cần tài khoản GCP
// ============================================================================
// Cách dùng: menu 🏭 TKSX > "⏱️ ĐO TỐC ĐỘ", hoặc mở Apps Script chọn hàm
// DO_TOC_DO rồi bấm ▶ Chạy. Kết quả hiện popup + ghi vào sheet "_TOCDO".
//
// AN TOÀN: không sửa dữ liệu thật.
//  - Các phép đo đọc: chỉ đọc.
//  - Phép đo ghi: thêm 3 dòng đánh dấu "___DO_TOC_DO___" vào cuối sheet data
//    rồi XÓA ngay trong khối finally.
//  - updateMaster_() chỉ tính lại đúng những gì vốn có (chạy lại nhiều lần vẫn ra y hệt).
// ============================================================================
function DO_TOC_DO() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var L = [];
  var add = function (s) { L.push(s); };
  var ms = function (fn) {
    var t = new Date().getTime();
    var r;
    try { r = fn(); } catch (err) { return { t: new Date().getTime() - t, err: '' + err }; }
    return { t: new Date().getTime() - t, v: r };
  };
  var fmt = function (x) {
    if (x.err) return 'LỖI: ' + x.err;
    return x.t + ' ms' + (x.t >= 5000 ? '   <<<<<< CHẬM' : (x.t >= 1500 ? '   <<< đáng chú ý' : ''));
  };

  add('=== BÁO CÁO ĐO TỐC ĐỘ TKSX ===');
  add('Thời điểm đo: ' + new Date());
  add('');

  // ---------- 1. Kích thước các sheet ----------
  add('--- 1. KÍCH THƯỚC CÁC SHEET ---');
  add('Tên sheet          | Dòng dùng | Dòng cấp phát | Cột cấp phát | Ô công thức');
  var all = ss.getSheets();
  var tongOTrong = 0;
  for (var i = 0; i < all.length; i++) {
    var s = all[i];
    var used = s.getLastRow(), maxR = s.getMaxRows(), maxC = s.getMaxColumns();
    // Đếm ô công thức trong tối đa 200 dòng đầu (đủ để phát hiện ARRAYFORMULA/QUERY)
    var nF = 0, mauF = '';
    try {
      var lim = Math.min(used > 0 ? used : 1, 200);
      var fs = s.getRange(1, 1, lim, maxC).getFormulas();
      for (var a = 0; a < fs.length; a++) {
        for (var b = 0; b < fs[a].length; b++) {
          if (fs[a][b]) { nF++; if (!mauF) mauF = fs[a][b].substring(0, 60); }
        }
      }
    } catch (err) { nF = -1; }
    tongOTrong += (maxR * maxC) - (used * maxC);
    add(pad_(s.getName(), 18) + ' | ' + pad_(used, 9) + ' | ' + pad_(maxR, 13) + ' | ' +
        pad_(maxC, 12) + ' | ' + nF + (mauF ? '   VD: ' + mauF : ''));
  }
  add('');
  add('Tổng số ô trống đã cấp phát (dòng thừa): ' + tongOTrong.toLocaleString());
  add('');

  // ---------- 2. Các phép đo ----------
  add('--- 2. ĐO TỪNG BƯỚC MÀ MỖI LẦN GỬI PHẢI LÀM ---');

  var m0 = ms(function () { return ss.getSheetByName('data').getRange(1, 1).getValue(); });
  add('Đọc 1 ô (đo độ trễ nền)            : ' + fmt(m0));

  var logSh = ss.getSheetByName('_log');
  var nLog = logSh ? logSh.getLastRow() : 0;
  var m1 = ms(function () {
    if (!logSh || nLog === 0) return null;
    return logSh.getRange(1, 1, nLog, 1)
      .createTextFinder('___KHONG_TON_TAI___').matchEntireCell(true).findNext();
  });
  add('Tra _log chống trùng (' + pad_(nLog, 6) + ' dòng)   : ' + fmt(m1));

  var m2 = ms(function () { return readDataVals_(ss); });
  var soDong = (m2.v && m2.v.length) ? m2.v.length : 0;
  add('Đọc CẢ sheet data (' + pad_(soDong, 6) + ' dòng)      : ' + fmt(m2));

  var m3 = ms(function () { SpreadsheetApp.flush(); return 1; });
  add('flush (đẩy thay đổi lên máy chủ)   : ' + fmt(m3));

  // Ghi thử 3 dòng vào cuối sheet data rồi xóa ngay -> đo đúng chi phí ghi + tính lại công thức
  var dataSh = ss.getSheetByName('data');
  var startTest = dataSh.getLastRow() + 1;
  var m4 = { t: 0 }, m5 = { t: 0 };
  try {
    var testRows = [];
    for (var k = 0; k < 3; k++) {
      testRows.push(['___DO_TOC_DO___', '', '', '', '', '___DO_TOC_DO___', '', '', '', '',
                     new Date(), '', '', '', '', '']);
    }
    m4 = ms(function () {
      dataSh.getRange(startTest, 1, 3, 16).setValues(testRows);
      SpreadsheetApp.flush();
      return 1;
    });
    add('Ghi 3 dòng vào sheet data          : ' + fmt(m4));
  } finally {
    // DỌN DẸP — phải chắc chắn, tuyệt đối không được xóa nhầm dữ liệu thật.
    // 2 lớp kiểm tra: (a) 3 dòng đó phải nằm đúng ở cuối sheet,
    //                 (b) cả 3 dòng phải mang dấu ___DO_TOC_DO___
    m5 = ms(function () {
      var last = dataSh.getLastRow();
      if (last !== startTest + 2) {
        return '⚠ VỊ TRÍ ĐỔI - hãy tự xóa tay các dòng có chữ ___DO_TOC_DO___';
      }
      var khop = true;
      var chk = dataSh.getRange(startTest, 1, 3, 1).getValues();
      for (var z = 0; z < 3; z++) {
        if (('' + chk[z][0]).indexOf('___DO_TOC_DO___') !== 0) khop = false;
      }
      if (!khop) {
        return '⚠ KHÔNG KHỚP DẤU - hãy tự xóa tay các dòng có chữ ___DO_TOC_DO___';
      }
      dataSh.deleteRows(startTest, 3);
      SpreadsheetApp.flush();
      return 'đã dọn sạch';
    });
    add('Dọn 3 dòng thử (' + (m5.v || m5.err) + ')  : ' + fmt(m5));
  }

  var m6 = ms(function () { updateMaster_(); return 1; });
  add('Tính lại toàn bộ Master            : ' + fmt(m6));

  add('');

  // ---------- 3. Trigger ----------
  add('--- 3. TRIGGER ĐANG CÀI ---');
  try {
    var trs = ScriptApp.getProjectTriggers();
    if (!trs.length) add('(không có trigger nào)');
    for (var q = 0; q < trs.length; q++) {
      add('- ' + trs[q].getHandlerFunction() + '  (' + trs[q].getEventType() + ')');
    }
  } catch (err) { add('Không đọc được trigger: ' + err); }
  add('');

  // ---------- 4. Tổng kết ----------
  var tongDoPost = (m1.t || 0) + (m2.t || 0) + (m4.t || 0) + (m6.t || 0);
  add('--- 4. TỔNG KẾT ---');
  add('Cộng 4 bước nặng nhất ≈ ' + tongDoPost + ' ms (' + Math.round(tongDoPost / 1000) + ' giây)');
  add('Nghi phạm lớn nhất: ' + nghiPham_([
    { ten: 'Tra _log', t: m1.t || 0 },
    { ten: 'Đọc cả sheet data', t: m2.t || 0 },
    { ten: 'Ghi dòng mới (gồm tính lại công thức)', t: m4.t || 0 },
    { ten: 'Tính lại Master', t: m6.t || 0 }
  ]));

  var text = L.join('\n');

  // Ghi ra sheet _TOCDO để copy dễ
  var out = ss.getSheetByName('_TOCDO') || ss.insertSheet('_TOCDO');
  out.clear();
  var lines = text.split('\n').map(function (x) { return [x]; });
  out.getRange(1, 1, lines.length, 1).setValues(lines);
  out.setColumnWidth(1, 900);

  try {
    SpreadsheetApp.getUi().alert('Kết quả đo tốc độ\n\n' + text +
      '\n\n(Đã ghi vào sheet "_TOCDO" để bạn copy gửi đi)');
  } catch (err) { }
  return text;
}

// ====== Kiểm tra: bản triển khai (web app) có đang chạy code mới không? ======
function KIEM_TRA_PHIEN_BAN() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var msg = 'PHIÊN BẢN CODE TRONG TRÌNH SOẠN THẢO:\n' + TKSX_VERSION + '\n\n';

  var sh = ss.getSheetByName('_TOCDO_LOG');
  if (!sh || sh.getLastRow() < 2) {
    msg += '❌ CHƯA CÓ dòng nào trong sheet _TOCDO_LOG.\n\n' +
           'Nghĩa là: kể từ khi dán code mới, CHƯA có lượt gửi nào chạy qua code mới.\n' +
           'Hãy quét thử 1 cuộn từ app rồi bấm lại menu này.\n\n' +
           'Nếu quét rồi mà vẫn không có dòng nào -> BẢN TRIỂN KHAI ĐANG CHẠY CODE CŨ.\n' +
           'Phải vào Triển khai > Quản lý triển khai > bút chì > Phiên bản: Mới > Triển khai.';
  } else {
    var n = Math.min(5, sh.getLastRow() - 1);
    var v = sh.getRange(sh.getLastRow() - n + 1, 1, n, 6).getValues();
    msg += n + ' LƯỢT GỬI GẦN NHẤT:\n\n';
    for (var i = v.length - 1; i >= 0; i--) {
      msg += '• ' + v[i][0] + '  |  ' + v[i][1] + '  |  ' + v[i][2] + ' giây\n' +
             '   ' + v[i][4] + '\n' +
             '   bản code: ' + v[i][5] + '\n\n';
    }
    msg += 'Nếu "bản code" ở trên KHÁC dòng đầu tiên của thông báo này\n' +
           '-> bản triển khai đang chạy code cũ, phải triển khai lại phiên bản mới.';
  }
  SpreadsheetApp.getUi().alert(msg);
  return msg;
}

function pad_(v, n) {
  var s = '' + v;
  while (s.length < n) s += ' ';
  return s;
}

function nghiPham_(arr) {
  arr.sort(function (a, b) { return b.t - a.t; });
  return arr[0].ten + ' (' + arr[0].t + ' ms)';
}

// ====== Backfill: tính lại cột Ngày SX (M/H) + Ca (N/I) cho TOÀN BỘ dữ liệu cũ ======
// Chạy 1 lần qua menu 🏭 TKSX > "📅 Cập nhật cột Ngày SX + Ca"
// Sau đó dữ liệu mới sẽ tự ghi 2 cột này khi nhận từ app.
function backfillNgayCa() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Tab data: cột M (13) = Ngày SX, cột N (14) = Ca ---
  var dataSh = ss.getSheetByName('data');
  if (dataSh && dataSh.getLastRow() >= 2) {
    // Đảm bảo header có 2 cột mới
    var hdr = dataSh.getRange(1, 13, 1, 2).getValues()[0];
    if (('' + hdr[0]).trim() !== 'Ngày SX') dataSh.getRange(1, 13).setValue('Ngày SX').setFontWeight('bold');
    if (('' + hdr[1]).trim() !== 'Ca') dataSh.getRange(1, 14).setValue('Ca').setFontWeight('bold');

    var nRows = dataSh.getLastRow() - 1;
    var times = dataSh.getRange(2, 1, nRows, 1).getValues();
    var out = [];
    for (var i = 0; i < nRows; i++) {
      out.push([prodDateKey_(times[i][0]), caKey_(times[i][0])]);
    }
    dataSh.getRange(2, 13, nRows, 2).setValues(out);
  }

  // --- Tab Phe: cột H (8) = Ngày SX, cột I (9) = Ca ---
  var pheSh = ss.getSheetByName('Phe');
  if (pheSh && pheSh.getLastRow() >= 2) {
    var pHdr = pheSh.getRange(1, 8, 1, 2).getValues()[0];
    if (('' + pHdr[0]).trim() !== 'Ngày SX') pheSh.getRange(1, 8).setValue('Ngày SX').setFontWeight('bold');
    if (('' + pHdr[1]).trim() !== 'Ca') pheSh.getRange(1, 9).setValue('Ca').setFontWeight('bold');

    var pn = pheSh.getLastRow() - 1;
    var pTimes = pheSh.getRange(2, 1, pn, 1).getValues();
    var pOut = [];
    for (var pi = 0; pi < pn; pi++) {
      pOut.push([prodDateKey_(pTimes[pi][0]), caKey_(pTimes[pi][0])]);
    }
    pheSh.getRange(2, 8, pn, 2).setValues(pOut);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Đã cập nhật Ngày SX + Ca cho ' +
    (dataSh ? (dataSh.getLastRow() - 1) : 0) + ' dòng data và ' +
    (pheSh ? (pheSh.getLastRow() - 1) : 0) + ' dòng Phế.');
}

// ====== App vừa ghi xong chưa? ======
// doPost đã tự cập nhật Master rồi, nên trigger không cần tính lại toàn bộ nữa
// (mỗi lần tính lại toàn bộ mất ~25 giây và ăn vào hạn mức chạy trigger của Google).
// Vẫn tính lại bình thường khi NGƯỜI sửa tay sheet data.
function vuaGhiBoiApp_() {
  try {
    var t = PropertiesService.getScriptProperties().getProperty('APP_WRITE_AT');
    if (!t) return false;
    return (new Date().getTime() - Number(t)) < 30000;   // trong vòng 30 giây
  } catch (err) { return false; }
}

// ====== Trigger: tự cập nhật Master khi sửa ô trong tab data ======
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() === 'data') {
      if (vuaGhiBoiApp_()) return;
      updateMaster_();
    }
  } catch (err) {}
}

// ====== Trigger: bắt cả xóa dòng, thêm dòng (onChange) ======
function onChange(e) {
  try {
    if (e.changeType === 'REMOVE_ROW' || e.changeType === 'INSERT_ROW' || e.changeType === 'OTHER') {
      var active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      if (active && active.getName() === 'data') {
        if (vuaGhiBoiApp_()) return;
        updateMaster_();
      }
    }
  } catch (err) {}
}

// ====== Cài đặt onChange trigger (chạy 1 lần) ======
function installOnChangeTrigger() {
  // Xóa trigger onChange cũ nếu có (tránh trùng)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'onChange') ScriptApp.deleteTrigger(t);
  });
  // Tạo trigger mới
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onChange()
    .create();
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Đã cài trigger onChange — Master sẽ tự cập nhật khi xóa/thêm dòng trong data.');
}

// ====== Đọc 1 sheet master -> [{code, name}] ======
function readMaster_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var out = [];
  vals.forEach(function (r) {
    var code = ('' + (r[0] === null ? '' : r[0])).trim();
    var nm = ('' + (r[1] === null ? '' : r[1])).trim();
    if (code || nm) out.push({ code: code, name: nm });
  });
  return out;
}

// Danh sách CN/Máy: ưu tiên sheet Master TRUNG TÂM (nếu Config có MASTER_URL),
// lỗi hoặc để trống thì dùng 4 tab local như cũ -> app không bao giờ chết vì thiếu master.
function getMasters_() {
  var cfg = getConfigMap_();
  if (cfg['MASTER_URL']) {
    var central = getCentralMastersCached_(cfg);
    if (central) return central;
  }
  return {
    cn_in:    readMaster_('CN_In_Master'),
    cn_thoi:  readMaster_('CN_Thoi_Master'),
    may_in:   readMaster_('May_In_Master'),
    may_thoi: readMaster_('May_Thoi_Master')
  };
}

// ====== Parse ngày dd/mm/yy hoặc dd/mm/yyyy (hoặc Date) -> Date 00:00, sai/trống -> null ======
function parseDMY_(v) {
  if (v instanceof Date) { var d = new Date(v); d.setHours(0, 0, 0, 0); return d; }
  var s = ('' + (v || '')).trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (!m) return null;
  var y = Number(m[3]);
  if (y < 100) y += 2000;
  return new Date(y, Number(m[2]) - 1, Number(m[1]));
}

// ====== Đọc 1 tab của sheet Master trung tâm theo TÊN CỘT (không phụ thuộc thứ tự cột) ======
// Nhận diện: Mã, Tên, Work station, Apply from, Apply to (from/to có thì lọc hiệu lực, không có thì bỏ qua)
// Trả về [{code, name, ws}] — chỉ gồm dòng còn hiệu lực tại hôm nay.
function readCentralTab_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var head = vals[0].map(function (h) { return ('' + h).trim().toLowerCase(); });
  var idx = function (names) {
    for (var i = 0; i < head.length; i++) { if (names.indexOf(head[i]) !== -1) return i; }
    return -1;
  };
  var iCode = idx(['mã', 'ma', 'mã nv', 'ma nv', 'code']);
  var iName = idx(['tên', 'ten', 'tên hiển thị', 'name']);
  var iWS   = idx(['work station', 'workstation', 'ws', 'tổ', 'to']);
  var iFrom = idx(['apply from', 'từ ngày', 'tu ngay']);
  var iTo   = idx(['apply to', 'đến ngày', 'den ngay']);
  if (iCode === -1 && iName === -1) return [];
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var code = iCode === -1 ? '' : ('' + (vals[r][iCode] == null ? '' : vals[r][iCode])).trim();
    var name = iName === -1 ? '' : ('' + (vals[r][iName] == null ? '' : vals[r][iName])).trim();
    if (!code && !name) continue;
    // Lọc hiệu lực: from > hôm nay hoặc to < hôm nay -> bỏ. Trống = vô hạn.
    if (iFrom !== -1) { var f = parseDMY_(vals[r][iFrom]); if (f && f > today) continue; }
    if (iTo !== -1)   { var t = parseDMY_(vals[r][iTo]);   if (t && t < today) continue; }
    var ws = iWS === -1 ? '' : ('' + (vals[r][iWS] == null ? '' : vals[r][iWS])).trim().toUpperCase();
    out.push({ code: code, name: name, ws: ws });
  }
  return out;
}

// ====== Đọc CN/Máy từ sheet Master trung tâm, tách theo Work station (Thổi/In) ======
// Lỗi (sai link / chưa share quyền / thiếu tab) -> trả null để getMasters_ fallback về tab local.
function getCentralMasters_(cfg) {
  var url = cfg['MASTER_URL'] || '';
  var wsThoi = ('' + (cfg['WS_THOI'] || 'THO')).trim().toUpperCase();
  var wsIn   = ('' + (cfg['WS_IN'] || 'INT')).trim().toUpperCase();
  var tabCN  = cfg['MASTER_TAB_CN'] || 'MA_CN';
  var tabMay = cfg['MASTER_TAB_MAY'] || 'MA_MAY';
  var idMatch = url.match(/\/d\/([a-zA-Z0-9\-_]+)/);
  var id = idMatch ? idMatch[1] : url;
  try {
    var oss = SpreadsheetApp.openById(id);
    var cn  = readCentralTab_(oss.getSheetByName(tabCN));
    var may = readCentralTab_(oss.getSheetByName(tabMay));
    var pick = function (rows, ws) {
      return rows
        .filter(function (r) { return r.ws === ws; })
        .map(function (r) { return { code: r.code, name: r.name }; });
    };
    var result = {
      cn_thoi:  pick(cn, wsThoi),
      cn_in:    pick(cn, wsIn),
      may_thoi: pick(may, wsThoi),
      may_in:   pick(may, wsIn)
    };
    // Không có dòng nào khớp Work station -> coi như cấu hình sai, fallback local
    if (!result.cn_thoi.length && !result.cn_in.length && !result.may_thoi.length && !result.may_in.length) return null;
    return result;
  } catch (err) {
    return null;
  }
}

// ====== Cache 2 phút cho master trung tâm (mở sheet ngoài hơi chậm) ======
function getCentralMastersCached_(cfg) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('central_masters_v1');
  if (hit) {
    try { return JSON.parse(hit); } catch (err) {}
  }
  var m = getCentralMasters_(cfg);
  if (m) { try { cache.put('central_masters_v1', JSON.stringify(m), 120); } catch (err) {} }
  return m;
}

// ====== Đọc PIN từ Config ======
function getPIN_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Config');
  if (!sh || sh.getLastRow() < 2) return { pin: '1234', pinChuyen: '1111' };
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var pin = '1234', pinChuyen = '1111';
  for (var i = 0; i < vals.length; i++) {
    var k = ('' + vals[i][0]).trim().toUpperCase();
    if (k === 'PIN') pin = ('' + vals[i][1]).trim();
    if (k === 'PIN_CHUYEN') pinChuyen = ('' + vals[i][1]).trim();
  }
  return { pin: pin, pinChuyen: pinChuyen };
}

// ====== Phân tích LSX: tách tiền tố bù ======
// N1-KH240615-01 -> {goc: 'KH240615-01', loaiBu: 'Nội bộ'}
// B2-KH240615-01 -> {goc: 'KH240615-01', loaiBu: 'Khách hàng'}
// KH240615-01    -> {goc: 'KH240615-01', loaiBu: ''}
function parseLSX_(lsx) {
  lsx = ('' + (lsx || '')).trim();
  var m = lsx.match(/^([NB])(\d+)-(.+)$/);
  if (m) {
    var prefix = m[1];
    var goc = m[3];
    var loai = prefix === 'N' ? 'Nội bộ' : 'Khách hàng';
    return { goc: goc, loaiBu: loai };
  }
  return { goc: lsx, loaiBu: '' };
}

// ====== Đọc toàn bộ Master data cho dashboard ======
function getMasterData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Master');
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues();
  var out = [];
  vals.forEach(function (r) {
    if (('' + r[1]).trim()) {
      out.push({
        loai: r[0], lsx: r[1], klyc: r[2], tongSX: r[3],
        conLai: r[4], soLanSX: r[5], lsxGoc: r[6], loaiBu: r[7]
      });
    }
  });
  return out;
}

// ====== Đọc toàn bộ Config -> object {KEY: value} ======
function getConfigMap_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Config');
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  vals.forEach(function (r) {
    var k = ('' + r[0]).trim().toUpperCase();
    if (k) out[k] = ('' + r[1]).trim();
  });
  return out;
}

// ====== Chuyển giá trị Thời gian -> chuỗi ngày 'yyyy-MM-dd' (ngày lịch) ======
// Hỗ trợ cả Date (Sheets tự nhận dạng) lẫn chuỗi 'dd/mm/yyyy hh:mm:ss' do app gửi
function dateKey_(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2) + '-' + ('0' + v.getDate()).slice(-2);
  }
  var s = ('' + (v || '')).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) {
    return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  return '';
}

// ====== Ngày SẢN XUẤT theo phân ca: Ca đêm (19h-7h) → ghi về ngày bắt đầu ca ======
// Ví dụ: 2:00 sáng ngày 14/7 → ngày SX = 13/7 (thuộc ca đêm ngày 13)
//        8:00 sáng ngày 14/7 → ngày SX = 14/7 (thuộc ca ngày 14)
function prodDateKey_(v) {
  var hour = getHour_(v);
  if (hour < 0) return dateKey_(v); // không lấy được giờ → fallback ngày lịch
  if (hour < 7) {
    // 0:00 - 6:59 → thuộc ca đêm hôm trước → lùi 1 ngày
    if (v instanceof Date) {
      var d2 = new Date(v.getTime() - 86400000); // trừ 24h
      return d2.getFullYear() + '-' + ('0' + (d2.getMonth() + 1)).slice(-2) + '-' + ('0' + d2.getDate()).slice(-2);
    }
    // Chuỗi: parse ngày rồi lùi 1
    var s = ('' + (v || '')).trim();
    var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      var d3 = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      d3.setDate(d3.getDate() - 1);
      return d3.getFullYear() + '-' + ('0' + (d3.getMonth() + 1)).slice(-2) + '-' + ('0' + d3.getDate()).slice(-2);
    }
    return dateKey_(v);
  }
  return dateKey_(v); // 7:00 trở đi → giữ nguyên ngày lịch
}

// ====== Xác định ca: 7h-19h = "Ngày", 19h-7h = "Đêm" ======
function caKey_(v) {
  var hour = getHour_(v);
  if (hour < 0) return '';
  return (hour >= 7 && hour < 19) ? 'Ngày' : 'Đêm';
}

// ====== Lấy giờ (số nguyên 0-23) từ Date hoặc chuỗi ======
function getHour_(v) {
  if (v instanceof Date) return v.getHours();
  var s = ('' + (v || '')).trim();
  var m = s.match(/(\d{1,2}):(\d{2})/);
  return m ? Number(m[1]) : -1;
}

// ====== Lấy giờ 'hh:mm' từ giá trị Thời gian (Date hoặc chuỗi 'dd/mm/yyyy hh:mm:ss') ======
function gioStr_(v) {
  if (v instanceof Date) {
    var p = function (n) { return ('0' + n).slice(-2); };
    return p(v.getHours()) + ':' + p(v.getMinutes());
  }
  var s = ('' + (v || '')).trim();
  var m = s.match(/(\d{1,2}):(\d{2})/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

// ====== Tổng hợp sản lượng + phế theo Tổ → Ngày SX → Ca → Máy, kèm chi tiết ======
// from, to: chuỗi 'yyyy-MM-dd' (ngày SẢN XUẤT, đã tính phân ca)
// Ca Ngày: 7h-19h, Ca Đêm: 19h-7h (phần 0h-7h tính về ngày hôm trước)
// Trả về: { sx, phe, sxd (chi tiết SX), phed (chi tiết phế) }
function getDailyData_(from, to) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var r2 = function (n) { return Math.round((Number(n) || 0) * 100) / 100; };

  // Mở rộng khoảng đọc: ca đêm kéo sang ngày lịch kế tiếp (0h-7h)
  // Nên cần đọc thêm 1 ngày sau "to" để bắt phần ca đêm
  var toPlus1 = (function(d) {
    var p = d.split('-');
    var dt = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    dt.setDate(dt.getDate() + 1);
    return dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
  })(to);

  // --- Sản xuất: gom theo Loại + Ngày SX + Ca + Máy, kèm chi tiết ---
  var sxMap = {}, sxd = [];
  var dataSh = ss.getSheetByName('data');
  if (dataSh && dataSh.getLastRow() >= 2) {
    var nRows = dataSh.getLastRow() - 1;
    // Bước 1: đọc cột Thời gian để khoanh vùng — mở rộng 1 ngày để bắt ca đêm
    var timeVals = dataSh.getRange(2, 1, nRows, 1).getValues();
    var minR = -1, maxR = -1;
    for (var ti = 0; ti < nRows; ti++) {
      var tk = dateKey_(timeVals[ti][0]);
      // Đọc từ (from) đến (to + 1 ngày) theo ngày lịch
      if (tk && tk >= from && tk <= toPlus1) {
        if (minR === -1) minR = ti;
        maxR = ti;
      }
    }
    // Bước 2: đọc khối dòng cần thiết
    var data = (minR === -1) ? [] : dataSh.getRange(minR + 2, 1, maxR - minR + 1, 10).getValues();
    // 0=Thời gian, 1=Loại, 2=CN, 3=Máy, 5=LSX, 8=KLCuộn, 9=Kết quả SX
    data.forEach(function (r) {
      var dk = prodDateKey_(r[0]); // Ngày SX theo phân ca
      if (!dk || dk < from || dk > to) return;
      var ca = caKey_(r[0]);       // Ca Ngày / Ca Đêm
      var loai = ('' + (r[1] || '')).trim();
      var may = ('' + (r[3] || '')).trim();
      var cn = '' + (r[2] || '');
      var rawKq = r[9];
      var kq = (rawKq instanceof Date) ? 0 : (Number(rawKq) || 0);
      if (kq > 1000000) kq = 0;
      var key = loai + '|' + dk + '|' + ca + '|' + may;
      if (!sxMap[key]) sxMap[key] = { loai: loai, ngay: dk, ca: ca, may: may, tong: 0, cuonSet: {} };
      sxMap[key].tong += kq;
      var cuonKey = ('' + r[0]) + '|' + cn + '|' + ('' + r[8]);
      sxMap[key].cuonSet[cuonKey] = 1;
      sxd.push({
        ngay: dk, ca: ca, gio: gioStr_(r[0]), loai: loai, may: may,
        lsx: '' + (r[5] || ''), cn: cn, kq: r2(kq)
      });
    });
  }
  var sxRows = [];
  Object.keys(sxMap).forEach(function (key) {
    var m = sxMap[key];
    sxRows.push({
      loai: m.loai, ngay: m.ngay, ca: m.ca, may: m.may,
      tong: r2(m.tong),
      soCuon: Object.keys(m.cuonSet).length
    });
  });
  sxRows.sort(function (a, b) {
    if (a.loai !== b.loai) return a.loai < b.loai ? -1 : 1;
    if (a.ngay !== b.ngay) return a.ngay < b.ngay ? -1 : 1;
    if (a.ca !== b.ca) return a.ca === 'Ngày' ? -1 : 1; // Ca Ngày trước
    return a.may < b.may ? -1 : 1;
  });

  // --- Phế: gom theo Loại + Ngày SX + Ca + Máy, kèm chi tiết ---
  var pheMap = {}, phed = [];
  var pheSh = ss.getSheetByName('Phe');
  if (pheSh && pheSh.getLastRow() >= 2) {
    var pn = pheSh.getLastRow() - 1;
    var pTimeVals = pheSh.getRange(2, 1, pn, 1).getValues();
    var pMinR = -1, pMaxR = -1;
    for (var pi = 0; pi < pn; pi++) {
      var pk = dateKey_(pTimeVals[pi][0]);
      if (pk && pk >= from && pk <= toPlus1) {
        if (pMinR === -1) pMinR = pi;
        pMaxR = pi;
      }
    }
    var pd = (pMinR === -1) ? [] : pheSh.getRange(pMinR + 2, 1, pMaxR - pMinR + 1, 7).getValues();
    // 0=Thời gian, 1=Loại, 2=CN, 3=Máy, 4=Phế trơn, 5=Phế màu
    pd.forEach(function (r) {
      var dk = prodDateKey_(r[0]);
      if (!dk || dk < from || dk > to) return;
      var ca = caKey_(r[0]);
      var loai = ('' + (r[1] || '')).trim();
      var may = ('' + (r[3] || '')).trim();
      var tron = Number(r[4]) || 0;
      var mau = Number(r[5]) || 0;
      var key = loai + '|' + dk + '|' + ca + '|' + may;
      if (!pheMap[key]) pheMap[key] = { loai: loai, ngay: dk, ca: ca, may: may, tron: 0, mau: 0 };
      pheMap[key].tron += tron;
      pheMap[key].mau += mau;
      phed.push({
        ngay: dk, ca: ca, gio: gioStr_(r[0]), loai: loai, may: may,
        cn: '' + (r[2] || ''), tron: r2(tron), mau: r2(mau)
      });
    });
  }
  var pheRows = [];
  Object.keys(pheMap).forEach(function (key) {
    var m = pheMap[key];
    pheRows.push({
      loai: m.loai, ngay: m.ngay, ca: m.ca, may: m.may,
      tron: r2(m.tron),
      mau: r2(m.mau)
    });
  });
  pheRows.sort(function (a, b) {
    if (a.loai !== b.loai) return a.loai < b.loai ? -1 : 1;
    if (a.ngay !== b.ngay) return a.ngay < b.ngay ? -1 : 1;
    if (a.ca !== b.ca) return a.ca === 'Ngày' ? -1 : 1;
    return a.may < b.may ? -1 : 1;
  });

  return { sx: sxRows, phe: pheRows, sxd: sxd, phed: phed };
}

// ====== Tra cứu khách hàng từ Sheet đơn hàng (cấu hình trong Config) ======
// Config cần 4 dòng: DONHANG_URL, DONHANG_TAB, DONHANG_COL_LSX, DONHANG_COL_KH
// Trả về: { khMap: {lsx: {kh, row}}, khUrl: link tab đơn hàng, khCol: cột LSX }
function getKHInfo_(masterRows) {
  var empty = { khMap: {}, khUrl: '', khCol: '' };
  var cfg = getConfigMap_();
  var url = cfg['DONHANG_URL'] || '';
  var tab = cfg['DONHANG_TAB'] || '';
  var colLSX = (cfg['DONHANG_COL_LSX'] || '').toUpperCase();
  var colKH = (cfg['DONHANG_COL_KH'] || '').toUpperCase();
  if (!url || !tab || !colLSX || !colKH) return empty;

  // Lấy ID sheet từ URL (hoặc chấp nhận dán thẳng ID)
  var idMatch = url.match(/\/d\/([a-zA-Z0-9\-_]+)/);
  var id = idMatch ? idMatch[1] : url;

  try {
    var oss = SpreadsheetApp.openById(id);
    var sh = oss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return empty;

    // Đổi chữ cột (A, B, ..., AA) -> số thứ tự cột
    var colToIdx = function (c) {
      c = ('' + c).trim().toUpperCase();
      var n = 0;
      for (var i = 0; i < c.length; i++) n = n * 26 + (c.charCodeAt(i) - 64);
      return n;
    };
    var iL = colToIdx(colLSX), iK = colToIdx(colKH);
    if (iL < 1 || iK < 1) return empty;
    // Chỉ đọc đúng 2 cột cần thiết (LSX + KH) — nhanh hơn nhiều so với đọc cả khối
    var nOrder = sh.getLastRow() - 1;
    var lsxVals = sh.getRange(2, iL, nOrder, 1).getValues();
    var khVals = sh.getRange(2, iK, nOrder, 1).getValues();

    // Chỉ tra các LSX đang có trong Master (cả mã gốc lẫn mã bù) cho nhẹ
    var wanted = {};
    (masterRows || []).forEach(function (m) {
      var g = ('' + (m.lsxGoc || '')).trim();
      var l = ('' + (m.lsx || '')).trim();
      if (g) wanted[g] = 1;
      if (l) wanted[l] = 1;
    });

    var khMap = {};
    for (var r = 0; r < nOrder; r++) {
      var lsx = ('' + (lsxVals[r][0] || '')).trim();
      if (!lsx || !wanted[lsx] || khMap[lsx]) continue;
      khMap[lsx] = {
        kh: ('' + (khVals[r][0] || '')).trim(),
        row: r + 2 // dòng thực tế trong sheet (header = dòng 1)
      };
    }

    var khUrl = 'https://docs.google.com/spreadsheets/d/' + id + '/edit#gid=' + sh.getSheetId();
    return { khMap: khMap, khUrl: khUrl, khCol: colLSX };
  } catch (err) {
    // Không có quyền / sai link -> bỏ qua, dashboard vẫn chạy bình thường
    return empty;
  }
}

// ====== Cache: tra KH có nhớ kết quả 10 phút (mở sheet đơn hàng khá chậm) ======
function getKHInfoCached_(masterRows) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('kh_info_v1');
  if (hit) {
    try { return JSON.parse(hit); } catch (err) {}
  }
  var kh = getKHInfo_(masterRows);
  try { cache.put('kh_info_v1', JSON.stringify(kh), 600); } catch (err) {}
  return kh;
}

// ====== App gọi để lấy danh sách CN/Máy + PIN (JSONP) ======
// ?action=version          -> trả phiên bản dữ liệu (siêu nhẹ, app so sánh để biết có cần tải lại)
// ?action=masterall        -> trả TOÀN BỘ Master + tên KH gắn sẵn từng dòng (app lưu về máy)
// ?action=master           -> (cũ) 300 LSX gần nhất + khMap
// ?action=daily&from=&to=  -> trả sản lượng + phế theo Tổ→Ngày→Máy + chi tiết (from/to: yyyy-MM-dd)
// ?action=chuyenstat&from=&to= -> thống kê chuyển cuộn theo khoảng ngày
// mặc định                 -> trả danh sách CN/Máy + PIN
// Kết quả master/daily được cache server-side để trả lời nhanh.
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var cache = CacheService.getScriptCache();

  if (action === 'version') {
    var ver = '0';
    try { ver = PropertiesService.getScriptProperties().getProperty('MASTER_VER') || '0'; } catch (err) {}
    return jsonOut_(e, { v: ver });
  }

  // Kiểm tra BẢN TRIỂN KHAI đang chạy code nào.
  // Mở link /exec?action=ping trên trình duyệt là biết ngay - đây là cách chắc chắn nhất
  // để phát hiện "đã dán code mới nhưng bản triển khai vẫn chạy code cũ".
  if (action === 'ping') {
    return jsonOut_(e, { ok: true, version: TKSX_VERSION, luc: new Date().toString() });
  }

  // Lịch sử SX của 1 LSX (đọc từ bảng data)
  if (action === 'history') {
    var lsxQ = ('' + (e.parameter.lsx || '')).trim();
    var hist = [];
    if (lsxQ) {
      var dSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');
      if (dSh && dSh.getLastRow() >= 2) {
        var dVals = dSh.getRange(2, 1, dSh.getLastRow() - 1, 12).getValues();
        for (var hi = 0; hi < dVals.length; hi++) {
          var hr = dVals[hi];
          if (('' + (hr[5] || '')).trim() !== lsxQ) continue;
          hist.push({
            time: fmtTime_(hr[0]),
            loai: '' + (hr[1] || ''),
            cn: '' + (hr[2] || ''),
            may: '' + (hr[3] || ''),
            klcuon: Math.round((Number(hr[8]) || 0) * 100) / 100,
            ketqua: Math.round((Number(hr[9]) || 0) * 100) / 100,
            canhbao: '' + (hr[11] || '')
          });
        }
      }
    }
    return jsonOut_(e, { lsx: lsxQ, rows: hist });
  }

  if (action === 'masterall') {
    var masterA = getMasterData_();
    var totalA = masterA.length;
    masterA.reverse(); // mới nhất lên đầu
    // Tra KH (có cache riêng 10 phút; nếu payload quá lớn không cache được thì tính lại mỗi lần)
    var khA = null;
    var khHit = cache.get('kh_full_v1');
    if (khHit) { try { khA = JSON.parse(khHit); } catch (err) {} }
    if (!khA) {
      khA = getKHInfo_(masterA);
      try { cache.put('kh_full_v1', JSON.stringify(khA), 600); } catch (err) {}
    }
    // Gắn tên KH + số dòng đơn hàng vào từng dòng Master (gọn hơn gửi map riêng)
    masterA.forEach(function (m) {
      var info = khA.khMap[('' + (m.lsxGoc || '')).trim()] || khA.khMap[('' + (m.lsx || '')).trim()];
      if (info) { m.kh = info.kh; m.khRow = info.row; }
    });
    var verA = '0';
    try { verA = PropertiesService.getScriptProperties().getProperty('MASTER_VER') || '0'; } catch (err) {}
    return jsonOut_(e, { rows: masterA, total: totalA, khUrl: khA.khUrl, khCol: khA.khCol, v: verA });
  }

  if (action === 'master') {
    // Cache 10 phút — tự xóa mỗi khi updateMaster_ chạy (có dữ liệu SX mới)
    var hit = cache.get('dash_master_v2');
    if (hit) return jsonOutRaw_(e, hit);
    var master = getMasterData_();
    var total = master.length;
    // Chỉ trả 300 LSX gần nhất (Master xếp theo thứ tự xuất hiện trong data
    // -> các dòng cuối là LSX mới nhất). Đảo lại để mới nhất nằm trên đầu.
    if (master.length > 300) master = master.slice(master.length - 300);
    master.reverse();
    var kh = getKHInfoCached_(master);
    var payload = JSON.stringify({ master: master, total: total, khMap: kh.khMap, khUrl: kh.khUrl, khCol: kh.khCol });
    try { cache.put('dash_master_v2', payload, 600); } catch (err) {}
    return jsonOutRaw_(e, payload);
  }

  if (action === 'daily') {
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var from = ('' + (e.parameter.from || today)).trim();
    var to = ('' + (e.parameter.to || today)).trim();
    if (from > to) { var tmp = from; from = to; to = tmp; }
    // Cache 2 phút cho từng khoảng ngày
    var ck = 'daily_v2_' + from + '_' + to;
    var hit2 = cache.get(ck);
    if (hit2) return jsonOutRaw_(e, hit2);
    var payload2 = JSON.stringify(getDailyData_(from, to));
    try { cache.put(ck, payload2, 120); } catch (err) {}
    return jsonOutRaw_(e, payload2);
  }

  // Thống kê chuyển cuộn theo khoảng ngày (đọc từ tab ChuyenCuon)
  if (action === 'chuyenstat') {
    var today3 = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var from3 = ('' + (e.parameter.from || today3)).trim();
    var to3 = ('' + (e.parameter.to || today3)).trim();
    if (from3 > to3) { var tmp3 = from3; from3 = to3; to3 = tmp3; }
    var ck3 = 'ccstat_v1_' + from3 + '_' + to3;
    var hit3 = cache.get(ck3);
    if (hit3) return jsonOutRaw_(e, hit3);
    var ccRows = [];
    var ccStatSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ChuyenCuon');
    if (ccStatSh && ccStatSh.getLastRow() >= 2) {
      var ccVals = ccStatSh.getRange(2, 1, ccStatSh.getLastRow() - 1, 5).getValues();
      ccVals.forEach(function (r) {
        var dk = dateKey_(r[0]);
        if (!dk || dk < from3 || dk > to3) return;
        ccRows.push({
          time: fmtTime_(r[0]),
          ngay: dk,
          lsx: '' + (r[1] || ''),
          kl: Math.round((Number(r[2]) || 0) * 100) / 100,
          ghichu: '' + (r[4] || '')
        });
      });
    }
    var payload3 = JSON.stringify({ rows: ccRows });
    try { cache.put(ck3, payload3, 120); } catch (err) {}
    return jsonOutRaw_(e, payload3);
  }

  var data = getMasters_();
  var pins = getPIN_();
  data.pin = pins.pin;
  data.pinChuyen = pins.pinChuyen;
  return jsonOut_(e, data);
}

// ====== Định dạng thời gian -> 'dd/mm/yyyy hh:mm' (nhanh, không dùng Utilities.formatDate) ======
function fmtTime_(v) {
  if (v instanceof Date) {
    var p = function (n) { return ('0' + n).slice(-2); };
    return p(v.getDate()) + '/' + p(v.getMonth() + 1) + '/' + v.getFullYear() + ' ' + p(v.getHours()) + ':' + p(v.getMinutes());
  }
  return ('' + (v || '')).trim();
}

// ====== Trả JSON đã stringify sẵn (dùng cho dữ liệu cache) ======
function jsonOutRaw_(e, jsonStr) {
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + jsonStr + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
}

// ====== Trả JSON (hỗ trợ cả JSONP nếu có ?callback=) ======
function jsonOut_(e, obj) {
  var out = JSON.stringify(obj);
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + out + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

// ====== Đọc toàn bộ dòng dữ liệu của tab data (16 cột, bỏ header) ======
// Gom về 1 chỗ để doPost chỉ phải đọc sheet MỘT LẦN rồi dùng chung cho
// reallocateGop_ / checkWarnings_ / calcKLChuyen_ / updateMaster_.
// Trước đây mỗi hàm tự đọc lại -> 1 lần gửi đọc cả sheet 3-4 lần, rất chậm khi data lớn.
function readDataVals_(ss, soCot) {
  var dataSh = ss.getSheetByName('data');
  if (!dataSh || dataSh.getLastRow() < 2) return [];
  return dataSh.getRange(2, 1, dataSh.getLastRow() - 1, soCot || 16).getValues();
}

// ====== Cập nhật tab Master sau mỗi lần ghi data ======
// Tính lại toàn bộ Master từ tab data (đơn giản, đảm bảo chính xác)
// dataVals: ảnh chụp data đã đọc sẵn (kể cả các dòng vừa ghi). Bỏ trống thì tự đọc.
function updateMaster_(dataVals) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var data = dataVals || readDataVals_(ss);
  if (!data.length) return;
  // data columns: 0=Thời gian, 1=Loại, 2=CN, 3=Máy, 4=DATA, 5=LSX, 6=KLYC, 7=PT, 8=KLCuộn, 9=KếtQuảSX, 10=GửiLúc, 11=CảnhBáo, 12=NgàySX, 13=Ca, 14=MãCN, 15=MãMáy

  // Gom theo key = Loại + "|" + LSX
  var map = {};
  data.forEach(function (r) {
    var loai = ('' + (r[1] || '')).trim();
    var lsx = ('' + (r[5] || '')).trim();
    if (!lsx) return;
    var key = loai + '|' + lsx;
    // Đọc KLYC — bỏ qua nếu là Date (tránh timestamp bị đọc nhầm thành số)
    var rawKlyc = r[6];
    var klyc = (rawKlyc instanceof Date) ? 0 : (Number(rawKlyc) || 0);
    // Giá trị > 1 triệu chắc chắn không phải KLYC (kg), có thể là timestamp lỗi
    if (klyc > 1000000) klyc = 0;

    var rawKetqua = r[9];
    var ketqua = (rawKetqua instanceof Date) ? 0 : (Number(rawKetqua) || 0);
    if (ketqua > 1000000) ketqua = 0;

    if (!map[key]) {
      map[key] = { loai: loai, lsx: lsx, klyc: klyc, tongSX: 0, soLanSX: 0, cnCodes: {}, mayCodes: {} };
    }
    // Nếu KLYC lần trước = 0 (do lỗi Date), lấy giá trị hợp lệ từ dòng tiếp theo
    if (map[key].klyc === 0 && klyc > 0) map[key].klyc = klyc;
    map[key].tongSX += ketqua;
    map[key].soLanSX += 1;
    // Gom mã CN và mã máy (cột O=14, P=15)
    var cnCode = ('' + (r[14] == null ? '' : r[14])).trim();
    var mayCode = ('' + (r[15] == null ? '' : r[15])).trim();
    if (cnCode) map[key].cnCodes[cnCode] = 1;
    if (mayCode) map[key].mayCodes[mayCode] = 1;
  });

  // Xây dựng mảng Master (số làm tròn 2 chữ số thập phân cho sạch)
  var r2 = function (n) { return Math.round((Number(n) || 0) * 100) / 100; };
  var rows = [];
  Object.keys(map).forEach(function (key) {
    var m = map[key];
    var p = parseLSX_(m.lsx);
    var cnList = Object.keys(m.cnCodes).join(', ');
    var mayList = Object.keys(m.mayCodes).join(', ');
    rows.push([
      m.loai,                    // A: Loại
      m.lsx,                     // B: LSX
      r2(m.klyc),                // C: KLYC
      r2(m.tongSX),              // D: Tổng SX
      r2(m.klyc - m.tongSX),     // E: Còn lại
      m.soLanSX,                 // F: Số lần SX
      p.goc,                     // G: LSX gốc
      p.loaiBu,                  // H: Loại bù
      '', '', '', '', '', '',    // I-N: trống
      cnList,                    // O: Mã CN
      mayList                    // P: Mã máy
    ]);
  });

  // Ghi lại tab Master
  var masterSh = ss.getSheetByName('Master') || ss.insertSheet('Master');
  if (masterSh.getLastRow() === 0) {
    masterSh.appendRow(['Loại', 'LSX', 'KLYC', 'Tổng SX', 'Còn lại', 'Số lần SX', 'LSX gốc', 'Loại bù', '', '', '', '', '', '', 'Mã CN', 'Mã máy']);
    masterSh.getRange(1, 1, 1, 16).setFontWeight('bold');
    masterSh.setFrozenRows(1);
  }
  // Ghi đè trực tiếp rồi chỉ xóa phần thừa phía dưới.
  // Trước đây xóa sạch toàn bộ Master rồi mới ghi lại -> tốn thêm 1 thao tác
  // trên cả bảng, mỗi lần công nhân gửi 1 cuộn đều phải làm.
  var oldLast = masterSh.getLastRow();
  if (rows.length > 0) {
    masterSh.getRange(2, 1, rows.length, 16).setValues(rows);
  }
  var tail = oldLast - 1 - rows.length;   // số dòng cũ còn thừa lại
  if (tail > 0) {
    masterSh.getRange(2 + rows.length, 1, tail, 16).clearContent();
  }

  // Xóa cache dashboard — lần mở tiếp theo sẽ lấy số liệu mới nhất
  try { CacheService.getScriptCache().remove('dash_master_v2'); } catch (err) {}
  // Tăng "phiên bản dữ liệu" — app so sánh để biết có cần tải lại hay không
  try { PropertiesService.getScriptProperties().setProperty('MASTER_VER', '' + Date.now()); } catch (err) {}
}

// ============================================================================
// CẬP NHẬT MASTER TỪNG PHẦN (incremental)
// ============================================================================
// Trước đây mỗi lần công nhân gửi 1 cuộn, hệ thống chép lại TOÀN BỘ bảng Master
// (29.744 dòng x 16 cột = 475.904 ô) chỉ để cộng thêm 1 con số -> mất ~25 giây.
// Nay chỉ sửa đúng những dòng LSX có trong lô vừa gửi -> còn dưới 1 giây.
//
// An toàn: nếu ai sửa tay / xóa dòng trong sheet data thì Master có thể lệch.
// Vì vậy VẪN GIỮ updateMaster_() (tính lại toàn bộ) ở 3 chỗ:
//   - menu 🏭 TKSX > Cập nhật Master (bấm tay bất cứ lúc nào)
//   - trigger onEdit/onChange khi NGƯỜI sửa sheet data
//   - lịch tự chạy lại toàn bộ mỗi đêm (installNightlyRebuild)
// ============================================================================

// Đọc cột A:B của Master -> bản đồ "Loại|LSX" -> số dòng.
// Chỉ đọc 2 cột nên nhẹ hơn nhiều so với đọc cả 16 cột.
function readMasterIndex_(ss) {
  var sh = ss.getSheetByName('Master');
  var idx = {};
  if (!sh || sh.getLastRow() < 2) return { sheet: sh, idx: idx };
  var ab = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < ab.length; i++) {
    var lsx = ('' + (ab[i][1] == null ? '' : ab[i][1])).trim();
    if (!lsx) continue;
    var key = ('' + (ab[i][0] == null ? '' : ab[i][0])).trim() + '|' + lsx;
    if (idx[key] === undefined) idx[key] = i + 2;   // số dòng thật trên sheet
  }
  return { sheet: sh, idx: idx };
}

// Tách chuỗi "CN01, CN02" thành tập hợp, giữ nguyên thứ tự xuất hiện
function splitSet_(v) {
  var out = [];
  var seen = {};
  ('' + (v == null ? '' : v)).split(',').forEach(function (x) {
    x = x.trim();
    if (x && !seen[x]) { seen[x] = 1; out.push(x); }
  });
  return { list: out, seen: seen };
}

// Gom lô vừa gửi theo key "Loại|LSX"
function aggBatch_(batchRows) {
  var agg = {};
  batchRows.forEach(function (r) {
    var loai = ('' + (r.loai || '')).trim();
    var lsx = ('' + (r.lsx || '')).trim();
    if (!lsx) return;
    var key = loai + '|' + lsx;
    if (!agg[key]) agg[key] = { loai: loai, lsx: lsx, klyc: 0, addSX: 0, addLan: 0, cn: [], may: [] };
    var klyc = Number(r.klyc) || 0;
    if (klyc > 1000000) klyc = 0;
    if (agg[key].klyc === 0 && klyc > 0) agg[key].klyc = klyc;
    var kq = Number(r.ketqua) || 0;
    if (kq > 1000000) kq = 0;
    agg[key].addSX += kq;
    agg[key].addLan += 1;
    var c = ('' + (r.cnCode || '')).trim();
    if (c && agg[key].cn.indexOf(c) === -1) agg[key].cn.push(c);
    var m = ('' + (r.mayCode || '')).trim();
    if (m && agg[key].may.indexOf(m) === -1) agg[key].may.push(m);
  });
  return agg;
}

// Đọc các dòng Master liên quan tới lô này (mỗi key 1 lệnh đọc 16 ô — rất nhẹ)
function readMasterRows_(mi, keys) {
  var cur = {};
  keys.forEach(function (k) {
    var row = mi.idx[k];
    cur[k] = row ? { row: row, vals: mi.sheet.getRange(row, 1, 1, 16).getValues()[0] } : null;
  });
  return cur;
}

// Lấy "đã SX" của từng key từ Master (thay cho việc đọc cả sheet data)
function prevMapFromMaster_(cur) {
  var m = {};
  Object.keys(cur).forEach(function (k) {
    m[k] = cur[k] ? (Number(cur[k].vals[3]) || 0) : 0;   // cột D = Tổng SX
  });
  return m;
}

// Ghi lại CHỈ những dòng Master bị ảnh hưởng
function applyMasterDelta_(ss, mi, cur, agg) {
  var r2 = function (n) { return Math.round((Number(n) || 0) * 100) / 100; };
  var themMoi = [];
  Object.keys(agg).forEach(function (key) {
    var a = agg[key];
    var p = parseLSX_(a.lsx);
    var c = cur[key];
    if (c) {
      var v = c.vals;
      var klyc = Number(v[2]) || 0;
      if (klyc === 0 && a.klyc > 0) klyc = a.klyc;
      var tong = r2((Number(v[3]) || 0) + a.addSX);
      var lan = (Number(v[5]) || 0) + a.addLan;
      var cnS = splitSet_(v[14]);
      a.cn.forEach(function (x) { if (!cnS.seen[x]) { cnS.seen[x] = 1; cnS.list.push(x); } });
      var myS = splitSet_(v[15]);
      a.may.forEach(function (x) { if (!myS.seen[x]) { myS.seen[x] = 1; myS.list.push(x); } });
      mi.sheet.getRange(c.row, 1, 1, 16).setValues([[
        a.loai, a.lsx, r2(klyc), tong, r2(klyc - tong), lan, p.goc, p.loaiBu,
        v[8], v[9], v[10], v[11], v[12], v[13],
        cnS.list.join(', '), myS.list.join(', ')
      ]]);
    } else {
      var tongM = r2(a.addSX);
      themMoi.push([
        a.loai, a.lsx, r2(a.klyc), tongM, r2(a.klyc - tongM), a.addLan, p.goc, p.loaiBu,
        '', '', '', '', '', '', a.cn.join(', '), a.may.join(', ')
      ]);
    }
  });
  if (themMoi.length) {
    mi.sheet.getRange(mi.sheet.getLastRow() + 1, 1, themMoi.length, 16).setValues(themMoi);
  }
  try { CacheService.getScriptCache().remove('dash_master_v2'); } catch (err) {}
  try { PropertiesService.getScriptProperties().setProperty('MASTER_VER', '' + Date.now()); } catch (err) {}
}

// ====== Lịch tự tính lại TOÀN BỘ Master mỗi đêm (chạy 1 lần để cài) ======
// Đây là lưới an toàn: nếu Master lệch vì bất cứ lý do gì, sáng hôm sau tự đúng lại.
// ⚠ ĐÂY LÀ HÀM CẦN CHẠY để CÀI LỊCH (đừng nhầm với nightlyRebuildMaster bên dưới)
function installNightlyRebuild() {
  var trs = ScriptApp.getProjectTriggers();
  trs.forEach(function (t) {
    if (t.getHandlerFunction() === 'nightlyRebuildMaster') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightlyRebuildMaster').timeBased().atHour(2).everyDays(1).create();

  // Đếm lại để xác nhận thật sự đã tạo được
  var dem = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'nightlyRebuildMaster') dem++;
  });
  var msg = dem > 0
    ? '✅ ĐÃ CÀI XONG lịch tự tính lại toàn bộ Master vào khoảng 2h sáng mỗi ngày.\n\n' +
      'Kiểm chứng: trong Apps Script, cột bên trái bấm biểu tượng ⏰ (Điều kiện kích hoạt)\n' +
      'sẽ thấy dòng "nightlyRebuildMaster".'
    : '❌ CHƯA tạo được lịch. Hãy thử lại, hoặc cài tay theo hướng dẫn.';
  // Chạy từ trình soạn thảo có thể không hiện được hộp thoại -> đừng để lỗi đó làm hỏng việc
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) {}
  try { console.log(msg); } catch (err) {}
  return msg;
}

// Hàm này CHỈ tính lại Master 1 lần — do lịch gọi mỗi đêm.
// Chạy tay hàm này KHÔNG tạo ra lịch. Muốn cài lịch phải chạy installNightlyRebuild.
function nightlyRebuildMaster() {
  updateMaster_();
}

// ====== Kiểm tra các lịch/trigger đang cài ======
function KIEM_TRA_LICH() {
  var msg = 'CÁC LỊCH / TRIGGER ĐANG CÀI:\n\n';
  var coLichDem = false;
  try {
    var trs = ScriptApp.getProjectTriggers();
    if (!trs.length) {
      msg += '(chưa có cái nào)\n';
    } else {
      trs.forEach(function (t) {
        var ten = t.getHandlerFunction();
        msg += '• ' + ten + '   —   loại: ' + t.getEventType() + '\n';
        if (ten === 'nightlyRebuildMaster') coLichDem = true;
      });
    }
  } catch (err) { msg += 'Không đọc được: ' + err + '\n'; }

  msg += '\n' + (coLichDem
    ? '✅ Lịch tính lại Master mỗi đêm ĐÃ ĐƯỢC CÀI.'
    : '❌ CHƯA có lịch tính lại Master mỗi đêm.\n' +
      '   Hãy chạy hàm  installNightlyRebuild  (KHÔNG phải nightlyRebuildMaster).');
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) {}
  try { console.log(msg); } catch (err) {}
  return msg;
}

// ====== Phân bổ lại GỘP dựa trên "Còn lại" thực tế từ Master ======
// Khi 1 LSX đã SX trước đó, phần phân bổ cho nó chỉ lấy tối đa = phần còn thiếu,
// phần dư dồn cho LSX khác trong cùng lô.
// Phương thức ĐƠN: không cắt, giữ nguyên.
// prevMapIn: bảng "đã SX" lấy sẵn từ Master (nhanh). Bỏ trống -> tự đọc sheet data như cũ.
function reallocateGop_(rows, ss, dataVals, prevMapIn) {
  if (!rows.length) return rows;
  // Chỉ áp dụng cho GỘP
  if (rows[0].phuongthuc !== 'GỘP') return rows;

  if (prevMapIn) return reallocateGopCore_(rows, prevMapIn);

  // dataVals = ảnh chụp sheet data đã đọc sẵn (tối ưu tốc độ).
  // Không có thì tự đọc như cũ — để onEdit/onChange gọi vẫn chạy bình thường.
  var prevMap = {}; // key: loai|lsx -> tổng KQ SX đã có
  var data = dataVals || readDataVals_(ss);
  if (data.length) {
    data.forEach(function (r) {
      var loai = ('' + (r[1] || '')).trim();
      var lsx = ('' + (r[5] || '')).trim();
      if (!lsx) return;
      var key = loai + '|' + lsx;
      if (!prevMap[key]) prevMap[key] = 0;
      prevMap[key] += (Number(r[9]) || 0);
    });
  }
  return reallocateGopCore_(rows, prevMap);
}

// Phần tính toán thuần tuý — tách riêng để dùng chung cho cả 2 nguồn số liệu
// (đọc từ sheet data như cũ, hoặc lấy sẵn từ Master cho nhanh). Logic KHÔNG đổi.
function reallocateGopCore_(rows, prevMap) {
  // Tính "còn lại" cho từng LSX trong batch
  var conLaiArr = [];
  var allFull = true;
  for (var i = 0; i < rows.length; i++) {
    var key = rows[i].loai + '|' + rows[i].lsx;
    var prevSX = prevMap[key] || 0;
    var klyc = Number(rows[i].klyc) || 0;
    var cl = klyc - prevSX;
    conLaiArr.push(cl);
    if (cl > 0) allFull = false;
  }

  // Nếu TẤT CẢ LSX đều đã đủ/vượt -> giữ nguyên phân bổ gốc (cảnh báo sẽ do checkWarnings_ xử lý)
  if (allFull) return rows;

  // Phân bổ lại: tổng KL cuộn cần chia
  var totalWeight = Number(rows[0].klcuon) || 0;
  var newKetqua = [];
  for (var j = 0; j < rows.length; j++) newKetqua.push(0);
  var remaining = totalWeight;

  // Vòng 1: mỗi LSX nhận tối đa = phần còn thiếu (conLai), không quá remaining
  for (var k = 0; k < rows.length && remaining > 0; k++) {
    var cap = Math.max(0, conLaiArr[k]); // không cho âm
    var take = Math.min(cap, remaining);
    newKetqua[k] = take;
    remaining -= take;
  }

  // Nếu vẫn còn dư (tất cả LSX đã lấp đầy nhưng cuộn nặng hơn) -> dồn vào LSX cuối
  if (remaining > 0) {
    newKetqua[newKetqua.length - 1] += remaining;
  }

  // Cập nhật ketqua trong rows
  for (var m = 0; m < rows.length; m++) {
    rows[m].ketqua = newKetqua[m];
  }

  return rows;
}

// ====== Tính KL chuyển cuộn cho từng LSX ======
// Mỗi lần chuyển = 1 cuộn SX cụ thể trong sheet data (cột J = Kết quả SX).
// Lần chuyển thứ N của LSX → lấy giá trị cột J dòng SX thứ N của LSX đó.
// Trả về mảng object [{kl, ghiChu}], mỗi phần tử tương ứng 1 LSX trong danh sách scan.
function calcKLChuyen_(ss, lsxList, dataVals) {
  if (!lsxList.length) return [];

  // Đọc tất cả dòng SX từ data, gom theo LSX → mảng giá trị cột J theo thứ tự + tổng SX + KLYC
  var sxMap = {};   // lsx -> [kq1, kq2, kq3, ...]
  var tongMap = {};  // lsx -> tổng Kết quả SX
  var klycMap = {};  // lsx -> KLYC (lấy giá trị hợp lệ đầu tiên)
  var data = dataVals || readDataVals_(ss);
  if (data.length) {
    data.forEach(function (r) {
      var lsx = ('' + (r[5] || '')).trim();
      if (!lsx) return;
      if (!sxMap[lsx]) sxMap[lsx] = [];
      var kq = (r[9] instanceof Date) ? 0 : (Number(r[9]) || 0);
      if (kq > 1000000) kq = 0;
      sxMap[lsx].push(kq); // cột J = Kết quả SX
      // Tổng SX
      if (!tongMap[lsx]) tongMap[lsx] = 0;
      tongMap[lsx] += kq;
      // KLYC — lấy giá trị hợp lệ (bỏ qua Date / số quá lớn)
      if (!klycMap[lsx]) {
        var rawKlyc = r[6];
        var klyc = (rawKlyc instanceof Date) ? 0 : (Number(rawKlyc) || 0);
        if (klyc > 0 && klyc <= 1000000) klycMap[lsx] = klyc;
      }
    });
  }

  // Đếm số lần đã chuyển THÀNH CÔNG của mỗi LSX từ ChuyenCuon
  // Chỉ đếm record không có ghi chú lỗi (cột E trống = chuyển thành công).
  // Record KL=0 + có ghi chú "Đã chuyển hết cuộn" = lần chuyển thất bại, không chiếm slot index.
  var countChuyenMap = {}; // lsx -> số lần đã chuyển thành công
  var ccSh = ss.getSheetByName('ChuyenCuon');
  if (ccSh && ccSh.getLastRow() >= 2) {
    var ccData = ccSh.getRange(2, 1, ccSh.getLastRow() - 1, 5).getValues();
    ccData.forEach(function (r) {
      var lsx = ('' + (r[1] || '')).trim(); // cột B = LSX
      if (!lsx) return;
      var ghiChu = ('' + (r[4] || '')).trim(); // cột E = Ghi chú
      // Bỏ qua record thất bại (có ghi chú lỗi)
      if (ghiChu) return;
      if (!countChuyenMap[lsx]) countChuyenMap[lsx] = 0;
      countChuyenMap[lsx]++;
    });
  }

  // Với mỗi LSX được scan: lấy cuộn SX tiếp theo chưa chuyển
  var result = [];
  var batchOffset = {};
  for (var i = 0; i < lsxList.length; i++) {
    var lsx = ('' + (lsxList[i] || '')).trim();
    var arr = sxMap[lsx] || [];
    var alreadyTransferred = countChuyenMap[lsx] || 0;
    if (!batchOffset[lsx]) batchOffset[lsx] = 0;
    var idx = alreadyTransferred + batchOffset[lsx];

    if (arr.length === 0) {
      // LSX không có trong sheet data
      result.push({ kl: 0, ghiChu: 'Không có dữ liệu thổi - in' });
    } else if (idx >= arr.length) {
      // Đã chuyển hết cuộn ĐÃ SX — kiểm tra KLYC để phân biệt
      var klyc = klycMap[lsx] || 0;
      var tongSX = tongMap[lsx] || 0;
      if (klyc > 0 && tongSX < klyc) {
        var conLai = Math.round((klyc - tongSX) * 100) / 100;
        result.push({ kl: 0, ghiChu: 'Đã chuyển hết cuộn đã SX (còn thiếu ' + conLai + ' kg chưa SX)' });
      } else {
        result.push({ kl: 0, ghiChu: 'Đã chuyển hết cuộn' });
      }
    } else {
      result.push({ kl: Math.round(arr[idx] * 100) / 100, ghiChu: '' });
    }
    batchOffset[lsx]++;
  }
  return result;
}

// ====== Kiểm tra cảnh báo: LSX nào bị vượt KLYC sau khi ghi nhận ======
// Áp dụng cho cả ĐƠN lẫn GỘP. Trả về mảng warning text cho từng row.
// prevVals: ảnh chụp sheet data TRƯỚC khi ghi batch này. Bỏ trống thì tự đọc như cũ.
// prevMapIn: bảng "đã SX" lấy sẵn từ Master (nhanh nhất) — ưu tiên dùng nếu có.
function checkWarnings_(ss, batchRows, prevVals, prevMapIn) {
  var warnings = [];
  if (!batchRows.length) return warnings;

  if (prevMapIn) return checkWarningsCore_(batchRows, prevMapIn);

  // Tính Tổng SX cũ (chỉ tính các dòng có TRƯỚC batch này)
  var prevMap = {}; // key: loai|lsx -> tổng SX trước đó
  var prev = null;
  if (prevVals) {
    prev = prevVals;
  } else {
    var dataSh = ss.getSheetByName('data');
    if (dataSh && dataSh.getLastRow() >= 2) {
      var lastRow = dataSh.getLastRow();
      var startOfBatch = lastRow - batchRows.length + 1;
      prev = (startOfBatch > 2) ? dataSh.getRange(2, 1, startOfBatch - 2, 10).getValues() : [];
    } else {
      prev = [];
    }
  }
  prev.forEach(function (r) {
    var key = ('' + (r[1] || '')).trim() + '|' + ('' + (r[5] || '')).trim();
    if (!prevMap[key]) prevMap[key] = 0;
    prevMap[key] += (Number(r[9]) || 0);
  });
  return checkWarningsCore_(batchRows, prevMap);
}

// Phần tính toán thuần tuý — dùng chung cho cả 2 nguồn số liệu. Logic KHÔNG đổi.
function checkWarningsCore_(batchRows, prevMap) {
  var warnings = [];

  // Kiểm tra từng row: nếu (đã SX trước + KQ lần này) > KLYC → cảnh báo
  batchRows.forEach(function (r) {
    var key = r.loai + '|' + r.lsx;
    var prevSX = prevMap[key] || 0;
    var klyc = Number(r.klyc) || 0;
    var ketqua = Number(r.ketqua) || 0;
    var totalAfter = prevSX + ketqua;
    var vuot = totalAfter - klyc;
    if (vuot > 0) {
      warnings.push('Vượt ' + Math.round(vuot * 100) / 100 + ' kg (KLYC=' + klyc + ', Tổng SX=' + Math.round(totalAfter * 100) / 100 + ')');
    } else {
      warnings.push('');
    }
  });

  return warnings;
}

// ====== App gửi kết quả sản xuất -> ghi vào tab data ======
// Body JSON: { batchId: "...", rows: [ {...}, ... ] }
// Trả về: { ok:true, dup:false, n:số_dòng }  hoặc  { ok:true, dup:true } nếu đã ghi trước đó
// ====== Đồng hồ đo từng bước trong doPost ======
// Kết quả in ra Nhật ký thực thi (Executions) của Apps Script, giúp biết CHÍNH XÁC
// bước nào đang ngốn thời gian thay vì phải đoán. Không ảnh hưởng dữ liệu.
// ĐỔI SỐ NÀY mỗi lần dán code mới -> nhìn sheet _TOCDO_LOG là biết bản nào đang chạy thật.
var TKSX_VERSION = 'v2026-08-20-c  (Master cộng dồn)';

var _tPrev = 0, _tAll = 0, _tSteps = [];
function _tStart_() { _tPrev = new Date().getTime(); _tAll = _tPrev; _tSteps = []; }
function _tMark_(label) {
  var now = new Date().getTime();
  _tSteps.push(label + ' ' + (now - _tPrev) + 'ms');
  _tPrev = now;
}

// Ghi kết quả đo VÀO SHEET "_TOCDO_LOG" (không cần GCP, không cần Nhật ký thực thi).
// Mỗi lần công nhân gửi 1 cuộn -> 1 dòng. Mở sheet đó ra là thấy ngay bước nào chậm.
function _tLog_(tag) {
  var tong = new Date().getTime() - _tAll;
  try { console.log('[TKSX ' + tag + '] TỔNG ' + tong + 'ms  <<  ' + _tSteps.join('  |  ')); } catch (err) {}
  try {
    if (PropertiesService.getScriptProperties().getProperty('TOCDO_OFF') === '1') return;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('_TOCDO_LOG');
    if (!sh) {
      sh = ss.insertSheet('_TOCDO_LOG');
      sh.appendRow(['Thời điểm', 'Loại', 'TỔNG (giây)', 'TỔNG (ms)', 'Chi tiết từng bước', 'Phiên bản code']);
      sh.getRange(1, 1, 1, 6).setFontWeight('bold');
      sh.setFrozenRows(1);
      sh.setColumnWidth(5, 700);
    }
    sh.appendRow([new Date(), tag, Math.round(tong / 100) / 10, tong,
                  _tSteps.join('  |  '), TKSX_VERSION]);
    // Chỉ giữ 300 dòng gần nhất cho gọn
    var last = sh.getLastRow();
    if (last > 301) sh.deleteRows(2, last - 301);
  } catch (err) {}
}

// Bật/tắt ghi nhật ký tốc độ (menu). Tắt đi khi đã xử lý xong cho nhẹ.
function toggleTocDoLog() {
  var p = PropertiesService.getScriptProperties();
  var dangTat = p.getProperty('TOCDO_OFF') === '1';
  p.setProperty('TOCDO_OFF', dangTat ? '0' : '1');
  SpreadsheetApp.getUi().alert(dangTat
    ? '✅ ĐÃ BẬT ghi nhật ký tốc độ vào sheet _TOCDO_LOG'
    : '⏹️ ĐÃ TẮT ghi nhật ký tốc độ');
}

function doPost(e) {
  _tStart_();
  var lock = LockService.getScriptLock();
  // QUAN TRỌNG: trước đây lỗi waitLock bị nuốt và code vẫn chạy tiếp KHÔNG có khóa.
  // Khi 2-3 máy gửi cùng lúc: phân bổ GỘP đọc số liệu cũ -> sai, updateMaster_ ghi đè nhau,
  // và kiểm tra trùng batchId cũng mất tác dụng.
  // Nay báo BUSY để app giữ lại và tự gửi lại sau (app không mất dữ liệu, công nhân không phải chờ).
  var gotLock = false;
  try { lock.waitLock(30000); gotLock = true; } catch (errLock) { gotLock = false; }
  if (!gotLock) {
    _tMark_('chờ-khóa-THẤT-BẠI');
    _tLog_('BUSY');
    return jsonOut_(e, { ok: false, error: 'BUSY - server dang ban, thu lai sau' });
  }
  _tMark_('chờ-khóa');
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('data') || ss.insertSheet('data');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Thời gian', 'Loại', 'Công nhân', 'Máy', 'DATA', 'LSX', 'Khối lượng yêu cầu', 'Phương thức', 'Khối lượng cuộn', 'Kết quả sản xuất', 'Gửi lúc', 'Cảnh báo', 'Ngày SX', 'Ca', 'Mã CN', 'Mã máy']);
    }
    var d = JSON.parse(e.postData.contents);
    var rows = d.rows || [];
    var batchId = ('' + (d.batchId || '')).trim();
    var dataType = ('' + (d.type || '')).trim(); // 'phe' hoặc rỗng (sản xuất)

    // Chống ghi trùng khi app gửi lại cùng 1 cuộn (mất mạng giữa chừng, retry...)
    if (batchId) {
      var log = ss.getSheetByName('_log') || ss.insertSheet('_log');
      if (log.getLastRow() === 0) { log.appendRow(['batchId', 'Gửi lúc']); }
      var found = log.getRange(1, 1, log.getLastRow(), 1)
        .createTextFinder(batchId).matchEntireCell(true).findNext();
      _tMark_('tra-_log(' + log.getLastRow() + ' dòng)');
      if (found) {
        _tLog_('DUP');
        return jsonOut_(e, { ok: true, dup: true, n: 0 });
      }
    }

    var sentAt = new Date();

    // ===== GHI PHẾ =====
    if (dataType === 'phe') {
      var pheSh = ss.getSheetByName('Phe') || ss.insertSheet('Phe');
      if (pheSh.getLastRow() === 0) {
        pheSh.appendRow(['Thời gian', 'Loại', 'Công nhân', 'Máy', 'Phế trơn', 'Phế màu', 'Gửi lúc', 'Ngày SX', 'Ca', 'Mã CN', 'Mã máy']);
      }
      var pheMatrix = rows.map(function (r) {
        var timeVal = r.time || '';
        return [
          timeVal,
          r.loai || '',
          r.cn || '',
          r.may || '',
          r.pheTron === undefined || r.pheTron === null ? 0 : r.pheTron,
          r.pheMau === undefined || r.pheMau === null ? 0 : r.pheMau,
          sentAt,
          prodDateKey_(timeVal),  // Ngày SX (cột H)
          caKey_(timeVal),        // Ca (cột I)
          r.cnCode || '',         // Mã CN (cột J)
          r.mayCode || ''         // Mã máy (cột K)
        ];
      });
      if (pheMatrix.length) {
        // Ghi 1 lệnh cho cả lô (thay vì appendRow từng dòng): nhanh hơn và
        // gần như không còn khoảng "đã ghi dữ liệu nhưng chưa kịp ghi _log".
        pheSh.getRange(pheSh.getLastRow() + 1, 1, pheMatrix.length, 11).setValues(pheMatrix);
      }
      if (batchId) {
        ss.getSheetByName('_log').appendRow([batchId, sentAt]);   // khóa chống trùng NGAY sau khi ghi
      }
      _tMark_('ghi-Phe+log');
      var outPhe = jsonOut_(e, { ok: true, dup: false, n: rows.length });
      _tLog_('PHE');
      return outPhe;
    }

    // ===== GHI CHUYỂN CUỘN =====
    if (dataType === 'chuyencuon') {
      var ccSh = ss.getSheetByName('ChuyenCuon') || ss.insertSheet('ChuyenCuon');
      if (ccSh.getLastRow() === 0) {
        ccSh.appendRow(['Thời gian', 'LSX', 'KL chuyển (kg)', 'Gửi lúc', 'Ghi chú']);
        ccSh.getRange(1, 1, 1, 5).setFontWeight('bold');
        ccSh.setFrozenRows(1);
      }

      // Lấy danh sách LSX từ client, tự tính KL chuyển từ data SX
      var lsxList = rows.map(function (r) { return ('' + (r.lsx || '')).trim(); });
      // Chuyển cuộn cần biết TỪNG cuộn của LSX (không gộp được từ Master),
      // nên vẫn phải đọc sheet data — nhưng chỉ 10 cột đầu thay vì 16.
      var ccPre = readDataVals_(ss, 10);
      _tMark_('đọc-sheet-data(' + ccPre.length + ' dòng x 10 cột)');
      var klArr = calcKLChuyen_(ss, lsxList, ccPre);
      _tMark_('tính-KL-chuyển');

      var ccMatrix = [];
      for (var ci = 0; ci < lsxList.length; ci++) {
        ccMatrix.push([
          rows[ci].time || '',
          lsxList[ci],
          klArr[ci].kl,
          sentAt,
          klArr[ci].ghiChu || ''
        ]);
      }
      if (ccMatrix.length) {
        // Ghi 1 lệnh cho cả lô
        ccSh.getRange(ccSh.getLastRow() + 1, 1, ccMatrix.length, 5).setValues(ccMatrix);
      }
      if (batchId) {
        ss.getSheetByName('_log').appendRow([batchId, sentAt]);   // khóa chống trùng NGAY sau khi ghi
      }
      _tMark_('ghi-ChuyenCuon+log');
      var outCC = jsonOut_(e, { ok: true, dup: false, n: lsxList.length });
      _tLog_('CHUYENCUON');
      return outCC;
    }

    // ===== GHI SẢN XUẤT (mặc định) =====
    // KHÔNG đọc sheet data nữa (39.000 dòng, mất ~10 giây).
    // Số "đã SX" của mỗi LSX vốn đã có sẵn ở cột D bảng Master -> lấy từ đó.
    var mi = readMasterIndex_(ss);                 // chỉ đọc 2 cột A:B
    _tMark_('đọc-Master-A:B(' + Object.keys(mi.idx).length + ' LSX)');

    // LƯỚI AN TOÀN: cách cộng dồn chỉ đúng khi Master đang khớp với sheet data.
    // Nếu Master trống (hoặc chưa có) mà data đã có dữ liệu -> cộng dồn sẽ ra số SAI.
    // Trường hợp đó phải dựng lại Master đầy đủ trước, dù chậm.
    if ((!mi.sheet || mi.sheet.getLastRow() < 2) && sh.getLastRow() >= 2) {
      updateMaster_();
      mi = readMasterIndex_(ss);
      _tMark_('DỰNG-LẠI-MASTER(Master trống)');
    }

    var agg = aggBatch_(rows);                     // gom lô theo Loại|LSX
    var keys = Object.keys(agg);
    var curRows = mi.sheet ? readMasterRows_(mi, keys) : {};
    var prevMap = prevMapFromMaster_(curRows);     // "đã SX" trước lô này
    _tMark_('đọc-' + keys.length + '-dòng-Master');

    // Phân bổ lại nếu GỘP (cắt theo "còn lại" thực tế, dồn dư cho LSX khác)
    rows = reallocateGop_(rows, ss, null, prevMap);
    _tMark_('phân-bổ-GỘP');

    // Ghi data (với ketqua đã được phân bổ lại nếu GỘP) + Ngày SX + Ca
    var matrix = rows.map(function (r) {
      var timeVal = r.time || '';
      return [
        timeVal,
        r.loai || '',
        r.cn || '',
        r.may || '',
        r.data || '',
        r.lsx || '',
        r.klyc === undefined || r.klyc === null ? '' : r.klyc,
        r.phuongthuc || '',
        r.klcuon === undefined || r.klcuon === null ? '' : r.klcuon,
        r.ketqua === undefined || r.ketqua === null ? '' : r.ketqua,
        sentAt,
        '',                       // Cảnh báo - cập nhật sau
        prodDateKey_(timeVal),    // Ngày SX (cột M)
        caKey_(timeVal),          // Ca (cột N)
        r.cnCode || '',           // Mã CN (cột O)
        r.mayCode || ''           // Mã máy (cột P)
      ];
    });

    var startRow = sh.getLastRow() + 1;
    if (matrix.length) {
      // Ghi 1 lệnh duy nhất cho cả lô thay vì appendRow từng dòng.
      // Vừa nhanh hơn nhiều, vừa thu hẹp gần như bằng 0 khoảng thời gian
      // "đã ghi data nhưng chưa ghi _log" — đây chính là khe hở gây GHI TRÙNG
      // khi script bị timeout giữa chừng và app gửi lại.
      sh.getRange(startRow, 1, matrix.length, 16).setValues(matrix);
    }
    _tMark_('GHI-DATA(' + matrix.length + ' dòng)');   // <- tới đây bạn đã thấy dữ liệu trên sheet

    // Ghi log chống trùng NGAY sau khi dữ liệu đã nằm trong sheet.
    // Các bước phía dưới (cảnh báo, updateMaster_) nếu lỗi cũng không gây ghi trùng nữa:
    // app gửi lại -> server thấy batchId trong _log -> trả dup:true.
    if (batchId) {
      ss.getSheetByName('_log').appendRow([batchId, sentAt]);
    }
    _tMark_('ghi-_log');

    // Kiểm tra cảnh báo vượt KLYC (dùng số "đã SX" lấy từ Master)
    var warnings = checkWarnings_(ss, rows, null, prevMap);
    var nWarn = 0;
    for (var i = 0; i < warnings.length; i++) {
      if (warnings[i]) {
        nWarn++;
        sh.getRange(startRow + i, 12).setValue(warnings[i]);
        sh.getRange(startRow + i, 1, 1, 16).setBackground('#fff3cd');
      }
    }
    _tMark_('cảnh-báo(' + nWarn + ')');

    // Cập nhật Master: CHỈ sửa những dòng LSX có trong lô này
    if (!mi.sheet) {
      updateMaster_();                             // Master chưa tồn tại -> dựng đầy đủ
    } else {
      // rows đã được reallocateGop_ cập nhật ketqua -> gom lại cho đúng
      applyMasterDelta_(ss, mi, curRows, aggBatch_(rows));
    }
    _tMark_('cập-nhật-MASTER');

    // Đánh dấu "app vừa ghi" để trigger onChange bỏ qua, không tính lại toàn bộ Master
    try { PropertiesService.getScriptProperties().setProperty('APP_WRITE_AT', '' + new Date().getTime()); } catch (err) {}

    var outSX = jsonOut_(e, { ok: true, dup: false, n: rows.length });
    _tLog_('SX');
    return outSX;
  } catch (err) {
    _tMark_('LỖI');
    _tLog_('ERROR ' + err);
    return jsonOut_(e, { ok: false, error: '' + err });
  } finally {
    try { lock.releaseLock(); } catch (err2) {}
  }
}
