/**
 * DuoPay Backend - Google Apps Script (Code.gs)
 * จัดการฐานข้อมูลรายจ่าย 2 คนและคำนวณการเคลียร์หนี้ลง Google Sheets
 */

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", message: "DuoPay API is running" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // เตรียม Sheet: Transactions
    let txSheet = ss.getSheetByName("Transactions");
    if (!txSheet) {
      txSheet = ss.insertSheet("Transactions");
      txSheet.appendRow(["ID", "วันที่ (พ.ศ.)", "เวลา", "ประเภท", "ผู้จ่าย/ผู้โอน", "ผู้รับ", "หมวดหมู่", "โน้ต", "วิธีหาร", "จำนวนเงิน (บาท)", "Timestamp"]);
      txSheet.getRange("A1:K1").setFontWeight("bold").setBackground("#FFE5EC");
    }

    if (action === "syncData") {
      const items = postData.transactions || [];
      // บันทึกรายการใหม่หรืออัปเดต
      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: items.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "unknown_action" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
