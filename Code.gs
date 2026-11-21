/**
 * NB DuoPay Backend - Google Apps Script (Code.gs)
 * เชื่อมต่อ Google Sheets สำหรับบันทึกรายจ่าย 2 คนและเคลียร์หนี้
 */

const SHEET_TX_NAME = "Transactions";
const SHEET_CFG_NAME = "Config";

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getAll";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    initSheets(ss);

    if (action === "getAll") {
      const txData = getTransactionsData(ss);
      const cfgData = getConfigData(ss);
      return createJsonResponse({
        status: "success",
        transactions: txData,
        config: cfgData
      });
    }

    return createJsonResponse({ status: "ok", message: "NB Backend is running" });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    initSheets(ss);

    if (action === "addBatch") {
      const items = postData.items || [];
      const txSheet = ss.getSheetByName(SHEET_TX_NAME);
      
      const rowsToAdd = items.map(it => [
        it.id,
        it.dateStr || "",
        it.timeStr || "",
        it.type || "expense",
        it.payerKey || "",
        it.receiverKey || "",
        it.catIcon || "",
        it.catName || "",
        it.amount || 0,
        it.splitMode || "half",
        it.note || "",
        it.timestamp || Date.now()
      ]);

      if (rowsToAdd.length > 0) {
        const lastRow = txSheet.getLastRow();
        txSheet.getRange(lastRow + 1, 1, rowsToAdd.length, 12).setValues(rowsToAdd);
      }

      return createJsonResponse({ status: "success", added: rowsToAdd.length });
    }

    if (action === "deleteItem") {
      const idToDelete = postData.id;
      const txSheet = ss.getSheetByName(SHEET_TX_NAME);
      const data = txSheet.getDataRange().getValues();
      let deleted = false;

      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(idToDelete)) {
          txSheet.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }

      return createJsonResponse({ status: deleted ? "success" : "not_found", id: idToDelete });
    }

    if (action === "settleDebt") {
      const settleTx = postData.tx;
      const txSheet = ss.getSheetByName(SHEET_TX_NAME);
      txSheet.appendRow([
        settleTx.id,
        settleTx.dateStr,
        settleTx.timeStr,
        "settle",
        settleTx.settlerKey,
        settleTx.receiverKey,
        "🤝",
        "เคลียร์หนี้",
        settleTx.amount,
        "settle",
        settleTx.note || "เคลียร์หนี้เรียบร้อย",
        settleTx.timestamp
      ]);
      return createJsonResponse({ status: "success", message: "Settled recorded" });
    }

    if (action === "saveConfig") {
      const cfg = postData.config;
      saveConfigData(ss, cfg);
      return createJsonResponse({ status: "success" });
    }

    return createJsonResponse({ status: "unknown_action" });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

function initSheets(ss) {
  let txSheet = ss.getSheetByName(SHEET_TX_NAME);
  if (!txSheet) {
    txSheet = ss.insertSheet(SHEET_TX_NAME);
    const headers = [
      "ID", "วันที่ (พ.ศ.)", "เวลา", "ประเภท", "ผู้จ่าย/ผู้โอน",
      "ผู้รับ", "ไอคอน", "หมวดหมู่", "จำนวนเงิน (บาท)", "วิธีหาร", "หมายเหตุ", "Timestamp"
    ];
    txSheet.appendRow(headers);
    txSheet.getRange("A1:L1")
      .setFontWeight("bold")
      .setBackground("#FFCCD5")
      .setFontColor("#C92A2A")
      .setHorizontalAlignment("center");
    txSheet.setFrozenRows(1);
  }

  let cfgSheet = ss.getSheetByName(SHEET_CFG_NAME);
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet(SHEET_CFG_NAME);
    cfgSheet.appendRow(["Key", "Value"]);
    cfgSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#FFE3E8");
  }
}

function getTransactionsData(ss) {
  const txSheet = ss.getSheetByName(SHEET_TX_NAME);
  if (!txSheet) return [];
  const rows = txSheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const list = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    list.push({
      id: r[0],
      dateStr: r[1],
      timeStr: r[2],
      type: r[3],
      payerKey: r[4],
      receiverKey: r[5],
      catIcon: r[6],
      catName: r[7],
      amount: parseFloat(r[8]) || 0,
      splitMode: r[9],
      note: r[10],
      timestamp: r[11]
    });
  }
  return list;
}

function getConfigData(ss) {
  const cfgSheet = ss.getSheetByName(SHEET_CFG_NAME);
  if (!cfgSheet) return null;
  const rows = cfgSheet.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) cfg[rows[i][0]] = rows[i][1];
  }
  return cfg;
}

function saveConfigData(ss, cfg) {
  const cfgSheet = ss.getSheetByName(SHEET_CFG_NAME);
  if (!cfgSheet) return;
  cfgSheet.clear();
  cfgSheet.appendRow(["Key", "Value"]);
  for (let k in cfg) {
    cfgSheet.appendRow([k, typeof cfg[k] === "object" ? JSON.stringify(cfg[k]) : cfg[k]]);
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
