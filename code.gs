const SITE_NAME = "岡本パン";
const RESERVATION_SHEET_NAME = "予約一覧";
const PRODUCT_SHEET_NAME = "商品管理";
const SPREADSHEET_PROPERTY_KEY = "RESERVATION_SPREADSHEET_ID";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("岡本パン 予約サイト")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function getProducts() {
  const spreadsheet = getOrCreateSpreadsheet_();
  const productSheet = getOrCreateProductSheet_(spreadsheet);
  return readProducts_(productSheet).filter(function(product) {
    return product.isActive;
  });
}

function submitReservation(data) {
  validateBasicReservation_(data);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getOrCreateSpreadsheet_();
    const productSheet = getOrCreateProductSheet_(spreadsheet);
    const reservationSheet = getOrCreateReservationSheet_(spreadsheet);
    const products = readProducts_(productSheet);
    const productMap = {};

    products.forEach(function(product) {
      productMap[String(product.id)] = product;
    });

    const normalizedItems = data.items.map(function(item) {
      const product = productMap[String(item.id)];

      if (!product || !product.isActive) {
        throw new Error("販売を終了した商品が含まれています。ページを再読み込みしてください。");
      }

      const quantity = Number(item.quantity);
      validateQuantity_(quantity);

      if (quantity > product.remaining) {
        throw new Error(
          product.name + "は残り" + formatQuantity_(product.remaining) +
          "個のため、選択された数量を予約できません。"
        );
      }

      return {
        id: product.id,
        name: product.name,
        price: product.price,
        quantity: quantity,
        rowNumber: product.rowNumber
      };
    });

    const reservationId = createReservationId_();
    const createdAt = new Date();
    const total = calculateTotal_(normalizedItems);
    const itemSummary = normalizedItems.map(function(item) {
      return item.name + " × " + formatQuantity_(item.quantity);
    }).join("\n");

    reservationSheet.appendRow([
      createdAt,
      reservationId,
      String(data.customerName || "").trim(),
      String(data.customerEmail || "").trim(),
      String(data.memo || "").trim(),
      itemSummary,
      total
    ]);

    normalizedItems.forEach(function(item) {
      const reservedCell = productSheet.getRange(item.rowNumber, 5);
      const currentReserved = Number(reservedCell.getValue()) || 0;
      reservedCell.setValue(currentReserved + item.quantity);
    });

    SpreadsheetApp.flush();

    sendConfirmationMail_({
      reservationId: reservationId,
      customerName: String(data.customerName || "").trim(),
      customerEmail: String(data.customerEmail || "").trim(),
      memo: String(data.memo || "").trim(),
      items: normalizedItems,
      total: total
    });

    return {
      success: true,
      reservationId: reservationId,
      products: readProducts_(productSheet).filter(function(product) {
        return product.isActive;
      })
    };
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const storedId = properties.getProperty(SPREADSHEET_PROPERTY_KEY);

  if (storedId) {
    try {
      return SpreadsheetApp.openById(storedId);
    } catch (error) {
      properties.deleteProperty(SPREADSHEET_PROPERTY_KEY);
    }
  }

  const spreadsheet = SpreadsheetApp.create("岡本パン 予約管理");
  properties.setProperty(SPREADSHEET_PROPERTY_KEY, spreadsheet.getId());
  console.log("予約管理スプレッドシート: " + spreadsheet.getUrl());
  return spreadsheet;
}

function getOrCreateReservationSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(RESERVATION_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(RESERVATION_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "予約日時",
      "予約番号",
      "お名前",
      "メールアドレス",
      "備考",
      "予約内容",
      "合計金額"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    sheet.getRange("A:A").setNumberFormat("yyyy/mm/dd hh:mm:ss");
    sheet.getRange("G:G").setNumberFormat("¥#,##0");
    sheet.autoResizeColumns(1, 7);
  }

  return sheet;
}

function getOrCreateProductSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(PRODUCT_SHEET_NAME);
  if (!sheet) {
    sheet = createProductSheet_(spreadsheet);
  }
  return sheet;
}

/**
 * 商品管理シートだけを新しい構成で作り直します。
 * 予約一覧シートは残ります。
 */
function rebuildProductSheet() {
  const spreadsheet = getOrCreateSpreadsheet_();
  const oldSheet = spreadsheet.getSheetByName(PRODUCT_SHEET_NAME);

  if (oldSheet) {
    spreadsheet.deleteSheet(oldSheet);
  }

  const newSheet = createProductSheet_(spreadsheet);
  spreadsheet.setActiveSheet(newSheet);

  console.log("商品管理シートを作り直しました: " + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

function createProductSheet_(spreadsheet) {
  const sheet = spreadsheet.insertSheet(PRODUCT_SHEET_NAME);

  sheet.appendRow([
    "商品ID",
    "商品名",
    "価格",
    "製造数",
    "予約数",
    "特徴",
    "原材料",
    "内容量",
    "保存期間",
    "画像URL1",
    "画像URL2",
    "画像URL3",
    "画像URL4",
    "画像URL5",
    "販売中"
  ]);

  const initialProducts = [
    [
      "campagne",
      "カンパーニュ",
      900,
      27,
      0,
      "静岡県産小麦と自家製酵母で長時間発酵。小麦の香りを楽しめる、岡本パンの定番です。",
      "小麦、塩、自家製酵母",
      "700g",
      "常温で3日程度。食べきれない場合はスライスして冷凍してください。",
      "https://picsum.photos/seed/campagne1/1200/800",
      "https://picsum.photos/seed/campagne2/1200/800",
      "https://picsum.photos/seed/campagne3/1200/800",
      "",
      "",
      true
    ],
    [
      "pain-de-mie",
      "パンドミ",
      650,
      32,
      0,
      "毎日食べたい、やさしい味わいの食事パンです。",
      "小麦、牛乳、バター、塩、酵母",
      "600g",
      "常温で2日程度。食べきれない場合は冷凍保存してください。",
      "https://picsum.photos/seed/paindemie1/1200/800",
      "https://picsum.photos/seed/paindemie2/1200/800",
      "",
      "",
      "",
      true
    ]
  ];

  sheet.getRange(2, 1, initialProducts.length, initialProducts[0].length)
    .setValues(initialProducts);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 15).setFontWeight("bold");
  sheet.getRange("C:C").setNumberFormat("¥#,##0");
  sheet.getRange("D:E").setNumberFormat("0.0");
  sheet.autoResizeColumns(1, 15);

  return sheet;
}

function readProducts_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

  return values.map(function(row, index) {
    const production = Number(row[3]) || 0;
    const reserved = Number(row[4]) || 0;

    const images = [
      String(row[9] || "").trim(),
      String(row[10] || "").trim(),
      String(row[11] || "").trim(),
      String(row[12] || "").trim(),
      String(row[13] || "").trim()
    ].filter(function(url) {
      return url !== "";
    });

    return {
      id: String(row[0]),
      name: String(row[1]),
      price: Number(row[2]) || 0,
      production: production,
      reserved: reserved,
      remaining: Math.max(0, Math.round((production - reserved) * 2) / 2),
      feature: String(row[5] || ""),
      ingredients: String(row[6] || ""),
      amount: String(row[7] || ""),
      shelfLife: String(row[8] || ""),
      images: images,
      isActive: row[14] === true || String(row[14]).toUpperCase() === "TRUE",
      rowNumber: index + 2
    };
  });
}

function validateBasicReservation_(data) {
  if (!data || typeof data !== "object") {
    throw new Error("予約データを受信できませんでした。");
  }

  const name = String(data.customerName || "").trim();
  const email = String(data.customerEmail || "").trim();

  if (!name) {
    throw new Error("お名前を入力してください。");
  }

  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error("正しいメールアドレスを入力してください。");
  }

  if (!Array.isArray(data.items) || data.items.length === 0) {
    throw new Error("パンを1つ以上選択してください。");
  }
}

function validateQuantity_(quantity) {
  if (
    !Number.isFinite(quantity) ||
    quantity <= 0 ||
    Math.round(quantity * 2) !== quantity * 2
  ) {
    throw new Error("予約数量に不正な値があります。");
  }
}

function calculateTotal_(items) {
  return items.reduce(function(total, item) {
    return total + Number(item.price) * Number(item.quantity);
  }, 0);
}

function createReservationId_() {
  const timeZone = Session.getScriptTimeZone() || "Asia/Tokyo";
  const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyyMMddHHmmss");
  const random = Math.floor(100 + Math.random() * 900);
  return "OP-" + timestamp + "-" + random;
}

function sendConfirmationMail_(reservation) {
  const itemLines = reservation.items.map(function(item) {
    const subtotal = Number(item.price) * Number(item.quantity);
    return [
      item.name,
      "数量：" + formatQuantity_(item.quantity),
      "小計：" + formatYen_(subtotal)
    ].join("\n");
  });

  const lines = [
    reservation.customerName + " 様",
    "",
    "岡本パンをご予約いただき、ありがとうございます。",
    "以下の内容で予約を受け付けました。",
    "",
    "【予約番号】",
    reservation.reservationId,
    "",
    "【予約内容】",
    itemLines.join("\n\n"),
    "",
    "【合計】",
    formatYen_(reservation.total),
    ""
  ];

  if (reservation.memo) {
    lines.push("【備考】", reservation.memo, "");
  }

  lines.push(
    "このメールは予約確認のため自動送信されています。",
    "",
    SITE_NAME
  );

  MailApp.sendEmail({
    to: reservation.customerEmail,
    subject: "【岡本パン】ご予約ありがとうございます",
    body: lines.join("\n"),
    name: SITE_NAME
  });
}

function formatQuantity_(value) {
  const number = Number(value);
  return Number.isInteger(number) ? String(number) : number.toFixed(1);
}

function formatYen_(value) {
  return "¥" + Number(value).toLocaleString("ja-JP");
}

function getReservationSpreadsheetUrl() {
  const id = PropertiesService.getScriptProperties()
    .getProperty(SPREADSHEET_PROPERTY_KEY);

  return id ? SpreadsheetApp.openById(id).getUrl() : "";
}
