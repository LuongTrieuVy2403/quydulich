// ID của Google Sheet
const SPREADSHEET_ID = "11PVVu-vBtoQ-Rhip3wviDxtq7Vp-xdNytw1VwwmzWog";

// Mật khẩu Admin của thủ quỹ
const ADMIN_PASSWORD = "24032004"; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Quản Lý Quỹ Du Lịch')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getMemberList() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('ThanhVien');
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return names.map(function(r) { return r[0]; }).filter(String);
}

function getMemberBalance(memberName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('ThanhVien');
  if (!sheet) return { found: false };
  
  var data = sheet.getDataRange().getValues();
  var keyword = memberName.toLowerCase().trim();
  
  if (!keyword) return { found: false };

  for (var i = 1; i < data.length; i++) {
    var fullName = data[i][0] ? data[i][0].toString().toLowerCase() : '';
    if (fullName.includes(keyword)) {
      return {
        found: true,
        name: data[i][0],
        daDong: data[i][1],
        daDung: data[i][2],
        conDu: data[i][3]
      };
    }
  }
  return { found: false };
}

function getSummary() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('TongQuan');
  if (!sheet) return { tongQuy: 0, daChi: 0, conLai: 0 };
  
  return {
    tongQuy: sheet.getRange('B1').getValue(),
    daChi: sheet.getRange('B2').getValue(),
    conLai: sheet.getRange('B3').getValue()
  };
}

// 1. CHỨC NĂNG TRỪ TIỀN (CHI TIÊU)
function addExpense(password, description, targetList, amount, note) {
  if (password !== ADMIN_PASSWORD) {
    return { success: false, message: "Mật khẩu Admin không chính xác!" };
  }
  
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('LichSuChiTieu');
  if (!sheet) {
    return { success: false, message: "Không tìm thấy sheet 'LichSuChiTieu'!" };
  }
  
  var targetText = "";
  if (Array.isArray(targetList)) {
    targetText = targetList.join(', ');
  } else {
    targetText = targetList;
  }
  
  sheet.appendRow([new Date(), description, targetText, Number(amount), note || '']);
  
  return { success: true, message: "Đã trừ tiền thành công!" };
}

// 2. CHỨC NĂNG MỚI: CỘNG TIỀN ĐÓNG CỦA THÀNH VIÊN
function addMemberContribution(password, memberName, amount) {
  if (password !== ADMIN_PASSWORD) {
    return { success: false, message: "Mật khẩu Admin không chính xác!" };
  }

  if (!memberName || !amount || Number(amount) <= 0) {
    return { success: false, message: "Vui lòng chọn thành viên và nhập số tiền hợp lệ!" };
  }

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('ThanhVien');
  if (!sheet) {
    return { success: false, message: "Không tìm thấy sheet 'ThanhVien'!" };
  }

  var data = sheet.getDataRange().getValues();
  var foundRow = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === memberName.trim()) {
      foundRow = i + 1; // Hàng thực tế trong Google Sheet (1-indexed)
      break;
    }
  }

  if (foundRow === -1) {
    return { success: false, message: "Không tìm thấy thành viên trong danh sách!" };
  }

  // Lấy tiền cũ và cộng dồn tiền mới
  var currentAmount = Number(sheet.getRange(foundRow, 2).getValue()) || 0;
  var newTotal = currentAmount + Number(amount);
  sheet.getRange(foundRow, 2).setValue(newTotal);

  return { 
    success: true, 
    message: "Đã nạp " + Number(amount).toLocaleString('vi-VN') + " đ cho " + memberName + " thành công!" 
  };
}

function getRecentExpenses() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('LichSuChiTieu');
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  
  var startRow = Math.max(2, lastRow - 9);
  var numRows = lastRow - startRow + 1;
  var values = sheet.getRange(startRow, 1, numRows, 5).getValues();
  
  return values.reverse().map(function(row) {
    var dateVal = row[0] instanceof Date ? row[0] : new Date(row[0]);
    return {
      time: Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "HH:mm dd/MM"),
      description: row[1],
      target: row[2],
      amount: row[3]
    };
  });
}