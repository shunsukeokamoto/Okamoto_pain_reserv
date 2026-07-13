const SITE_NAME = "岡本パン";
const SHEET_NAME = "予約一覧";
const SPREADSHEET_PROPERTY_KEY = "RESERVATION_SPREADSHEET_ID";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("岡本パン 予約サイト")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function submitReservation(data) {
  validateReservation_(data);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const reservationId = createReservationId_();
    const createdAt = new Date();
    const total = calculateTotal_(data.items);
    const spreadsheet = getOrCreateSpreadsheet_();
    const sheet = getOrCreateReservationSheet_(spreadsheet);
    const itemSummary = data.items.map(function(item) {
      return item.name + " × " + formatQuantity_(item.quantity);
    }).join("\n");
    sheet.appendRow([createdAt,reservationId,data.customerName,data.customerEmail,data.memo || "",itemSummary,total]);
    sendConfirmationMail_({reservationId:reservationId,customerName:data.customerName,customerEmail:data.customerEmail,memo:data.memo || "",items:data.items,total:total});
    return {success:true,reservationId:reservationId};
  } finally {
    lock.releaseLock();
  }
}

function validateReservation_(data) {
  if (!data || typeof data !== "object") throw new Error("予約データを受信できませんでした。");
  const name = String(data.customerName || "").trim();
  const email = String(data.customerEmail || "").trim();
  if (!name) throw new Error("お名前を入力してください。");
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error("正しいメールアドレスを入力してください。");
  if (!Array.isArray(data.items) || data.items.length === 0) throw new Error("パンを1つ以上選択してください。");
  data.items.forEach(function(item) {
    const quantity = Number(item.quantity), price = Number(item.price);
    if (!item.name || !Number.isFinite(quantity) || quantity <= 0 || Math.round(quantity * 2) !== quantity * 2 || !Number.isFinite(price) || price < 0) {
      throw new Error("予約内容に不正な値があります。");
    }
  });
}

function calculateTotal_(items) {
  return items.reduce(function(total,item){return total + Number(item.price) * Number(item.quantity);},0);
}

function createReservationId_() {
  const timeZone = Session.getScriptTimeZone() || "Asia/Tokyo";
  const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyyMMddHHmmss");
  const random = Math.floor(100 + Math.random() * 900);
  return "OP-" + timestamp + "-" + random;
}

function getOrCreateSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const storedId = properties.getProperty(SPREADSHEET_PROPERTY_KEY);
  if (storedId) {
    try { return SpreadsheetApp.openById(storedId); }
    catch (error) { properties.deleteProperty(SPREADSHEET_PROPERTY_KEY); }
  }
  const spreadsheet = SpreadsheetApp.create("岡本パン 予約一覧");
  properties.setProperty(SPREADSHEET_PROPERTY_KEY, spreadsheet.getId());
  console.log("予約一覧スプレッドシート: " + spreadsheet.getUrl());
  return spreadsheet;
}

function getOrCreateReservationSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = spreadsheet.getSheets()[0]; sheet.setName(SHEET_NAME); }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["予約日時","予約番号","お名前","メールアドレス","備考","予約内容","合計金額"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,7).setFontWeight("bold");
    sheet.getRange("A:A").setNumberFormat("yyyy/mm/dd hh:mm:ss");
    sheet.getRange("G:G").setNumberFormat("¥#,##0");
    sheet.autoResizeColumns(1,7);
  }
  return sheet;
}

function sendConfirmationMail_(reservation) {
  const itemLines = reservation.items.map(function(item) {
    const subtotal = Number(item.price) * Number(item.quantity);
    return [item.name,"数量：" + formatQuantity_(item.quantity),"小計：" + formatYen_(subtotal)].join("\n");
  });
  const lines = [
    reservation.customerName + " 様","",
    "岡本パンをご予約いただき、ありがとうございます。",
    "以下の内容で予約を受け付けました。","",
    "【予約番号】",reservation.reservationId,"",
    "【予約内容】",itemLines.join("\n\n"),"",
    "【合計】",formatYen_(reservation.total),""
  ];
  if (reservation.memo) lines.push("【備考】",reservation.memo,"");
  lines.push("このメールは予約確認のため自動送信されています。","",SITE_NAME);
  MailApp.sendEmail({to:reservation.customerEmail,subject:"【岡本パン】ご予約ありがとうございます",body:lines.join("\n"),name:SITE_NAME});
}

function formatQuantity_(value) {
  const number = Number(value);
  return Number.isInteger(number) ? String(number) : number.toFixed(1);
}

function formatYen_(value) {
  return "¥" + Number(value).toLocaleString("ja-JP");
}

function getReservationSpreadsheetUrl() {
  const id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_PROPERTY_KEY);
  return id ? SpreadsheetApp.openById(id).getUrl() : "";
}
