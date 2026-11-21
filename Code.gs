/**
 * Lunar (NB) Backend - Google Apps Script (Code.gs)
 * เชื่อมต่อ Google Sheets สำหรับบันทึกรายจ่าย 2 คนและเคลียร์หนี้
 * รองรับการแจ้งเตือนผ่าน LINE Messaging API
 */

const SPREADSHEET_ID = "1C2zEgZLgSyGdD_Ef9gEezVYredxy1-SgYfHpOPcEV7o";
const SHEET_TX_NAME = "Transactions";
const SHEET_CFG_NAME = "Config";

// ตั้งค่า LINE Notification (สามารถใส่ Token และ User/Group ID ได้ที่นี่ หรือระบุใน Sheet Config)
const LINE_CHANNEL_ACCESS_TOKEN = ""; // ใส่ Channel Access Token ของ LINE Messaging API
const LINE_TARGET_ID = "";            // ใส่ User ID หรือ Group ID ที่ต้องการให้แจ้งเตือน

/**
 * ดึง Spreadsheet อัตโนมัติ (รองรับทั้งเปิดตรงและเปิดผ่าน ID)
 */
function getSpreadsheet() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/**
 * ฟังก์ชันสำหรับทดสอบรันใน Apps Script Editor เพื่อขอสิทธิ์ (Authorize)
 */
function initSheets(ss) {
  if (!ss) ss = getSpreadsheet();
  
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
      .setBackground("#FFD166")
      .setFontColor("#0B0C10")
      .setHorizontalAlignment("center");
    txSheet.setFrozenRows(1);
  }

  let cfgSheet = ss.getSheetByName(SHEET_CFG_NAME);
  if (!cfgSheet) {
    cfgSheet = ss.insertSheet(SHEET_CFG_NAME);
    cfgSheet.appendRow(["Key", "Value"]);
    cfgSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#FFEAA7");
  }

  Logger.log("✅ สร้าง/ตรวจสอบ Sheet เรียบร้อยแล้ว: " + ss.getName());
  return "OK";
}

function testConnection() {
  return initSheets();
}

/**
 * จัดการคำขอ HTTP GET (ดึงข้อมูล)
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getAll";
    const ss = getSpreadsheet();
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

    return createJsonResponse({ status: "ok", message: "Lunar Backend is running" });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * จัดการคำขอ HTTP POST (บันทึก / ลบ / เคลียร์หนี้ / ตั้งค่าผู้ใช้)
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const ss = getSpreadsheet();
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

        // ส่งการแจ้งเตือนผ่าน LINE
        notifyLineBatch(items, ss);
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

      // ส่งการแจ้งเตือนผ่าน LINE
      notifyLineSettle(settleTx, ss);

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
  if (rows.length <= 1) return null;
  const cfg = {};
  for (let i = 1; i < rows.length; i++) {
    const k = rows[i][0];
    const val = rows[i][1];
    if (k) {
      try {
        cfg[k] = JSON.parse(val);
      } catch (e) {
        cfg[k] = val;
      }
    }
  }
  return cfg;
}

function saveConfigData(ss, cfg) {
  const cfgSheet = ss.getSheetByName(SHEET_CFG_NAME);
  if (!cfgSheet) return;
  cfgSheet.clear();
  cfgSheet.appendRow(["Key", "Value"]);
  for (let k in cfg) {
    if (cfg.hasOwnProperty(k)) {
      cfgSheet.appendRow([k, typeof cfg[k] === "object" ? JSON.stringify(cfg[k]) : cfg[k]]);
    }
  }
}

/**
 * ระบบส่งการแจ้งเตือนผ่าน LINE Messaging API
 */
function sendLinePush(messageText, ss) {
  let token = LINE_CHANNEL_ACCESS_TOKEN;
  let targetId = LINE_TARGET_ID;

  // หากไม่ได้ระบุในโค้ด ให้ค้นหาจาก Sheet Config
  if (!token || !targetId) {
    const cfg = getConfigData(ss);
    if (cfg) {
      token = token || cfg["line_token"];
      targetId = targetId || cfg["line_target_id"];
    }
  }

  if (!token || !targetId) {
    Logger.log("ไม่ได้กำหนด LINE Token หรือ Target ID");
    return;
  }

  const url = "https://api.line.me/v2/bot/message/push";
  const payload = {
    to: targetId,
    messages: [
      {
        type: "text",
        text: messageText
      }
    ]
  };

  try {
    UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log("LINE Notification Error: " + err.toString());
  }
}

function notifyLineBatch(items, ss) {
  if (!items || items.length === 0) return;
  const cfg = getConfigData(ss) || {};
  const payerKey = items[0].payerKey;
  let payerName = "ผู้ใช้";
  if (cfg[payerKey] && cfg[payerKey].name) {
    payerName = cfg[payerKey].name;
  } else if (payerKey === "user1") {
    payerName = "บอส";
  } else if (payerKey === "user2") {
    payerName = "คู่หู";
  }

  let msg = "🌙 Lunar — บันทึกรายจ่ายใหม่!\n";
  msg += "👤 ผู้บันทึก: " + payerName + "\n";
  msg += "📅 วันที่: " + (items[0].dateStr || "") + " " + (items[0].timeStr || "") + "\n\n";
  
  let total = 0;
  items.forEach(it => {
    const splitLabel = it.splitMode === "half" ? "จ่ายคนละครึ่ง" : "จ่ายเต็ม";
    msg += "• " + it.catIcon + " " + it.note + " : ฿" + it.amount.toLocaleString() + " (" + splitLabel + ")\n";
    total += it.amount;
  });

  msg += "\n💰 ยอดรวมรอบนี้: ฿" + total.toLocaleString() + " บาท";

  sendLinePush(msg, ss);
}

function notifyLineSettle(settleTx, ss) {
  const cfg = getConfigData(ss) || {};
  let settlerName = (cfg[settleTx.settlerKey] && cfg[settleTx.settlerKey].name) || "ผู้โอน";
  let receiverName = (cfg[settleTx.receiverKey] && cfg[settleTx.receiverKey].name) || "ผู้รับ";

  let msg = "🤝 Lunar — แจ้งเตือนการเคลียร์หนี้!\n\n";
  msg += "💸 ยอดเงิน: ฿" + settleTx.amount.toLocaleString() + " บาท\n";
  msg += "🔄 " + settlerName + " ➜ " + receiverName + "\n";
  msg += "📅 วันที่: " + settleTx.dateStr + " " + settleTx.timeStr + "\n";
  msg += "✨ สถานะ: เคลียร์ยอดเรียบร้อยแล้ว";

  sendLinePush(msg, ss);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
