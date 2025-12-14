/**
 * Google Sheet Auto Line Break Tool
 * Author: [Tên của bạn]
 * License: MIT
 */

/**
 * --- PHẦN 1: KHỞI TẠO MENU ---
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⚡ Công Cụ Xử Lý')
      .addItem('Mở bảng điều khiển...', 'showDialog')
      .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
      .setWidth(420)
      .setHeight(500)
      .setTitle('Công Cụ Xuống Dòng Tự Động');
  SpreadsheetApp.getUi().showModalDialog(html, 'Công Cụ Xuống Dòng Tự Động');
}

/**
 * --- PHẦN 2: XỬ LÝ DỮ LIỆU TỐI ƯU (CORE) ---
 */
function processFormat(settings) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();

  if (!range) {
    return { success: false, message: "⚠️ Lỗi: Bạn chưa chọn vùng dữ liệu!" };
  }

  // 1. Xây dựng Regex (Biểu thức chính quy)
  var regexParts = [];
  
  if (settings.useCircleNum) regexParts.push('[\\u2460-\\u2473]'); // ①...⑳
  if (settings.useBrackets) regexParts.push('[\\u3010\\u3011]');   // 【 】

  // Xử lý ký tự tùy chỉnh
  if (settings.customChars && settings.customChars.trim()) {
    var tokens = settings.customChars.split(',');
    tokens.forEach(function(token) {
      token = token.trim();
      if (token) {
        // Escape ký tự đặc biệt an toàn tuyệt đối
        regexParts.push(token.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
      }
    });
  }

  if (regexParts.length === 0) {
     return { success: false, message: "⚠️ Bạn chưa chọn điều kiện nào!" };
  }

  // Tạo Regex gộp: (A|B|C)
  var regex = new RegExp('(' + regexParts.join('|') + ')', 'g');

  // 2. Xử lý dữ liệu trong bộ nhớ (In-memory processing)
  var values = range.getValues();
  var isModified = false; // Cờ kiểm tra xem có thay đổi gì không
  var count = 0;

  // Dùng hàm .map để duyệt mảng nhanh hơn vòng lặp for truyền thống
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      // Chỉ xử lý nếu ô là chuỗi (tránh lỗi ngày tháng, số)
      if (typeof cell !== 'string' || !cell) return cell;

      // Kiểm tra nhanh trước khi xử lý sâu
      if (!regex.test(cell)) return cell;

      regex.lastIndex = 0; // Reset regex

      // THAY THẾ & LÀM SẠCH (One-line logic)
      // 1. Thay thế ký tự khớp bằng \n + ký tự đó
      // 2. .replace(/^\n+/, '') -> Xóa ngay lập tức các dấu xuống dòng dư thừa ở đầu chuỗi
      var newTxt = cell.replace(regex, '\n$1').replace(/^\n+/, '');

      if (newTxt !== cell) {
        isModified = true;
        count++;
        return newTxt;
      }
      return cell;
    });
  });

  // 3. Chỉ ghi lại vào Sheet nếu thực sự có thay đổi (Tiết kiệm API)
  if (isModified) {
    range.setValues(newValues);
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); // Bật Wrap Text
    return { success: true, count: count };
  } else {
    return { success: true, count: 0, message: "Không có ô nào cần thay đổi." };
  }
}
