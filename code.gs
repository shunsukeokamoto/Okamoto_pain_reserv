const SITE_NAME = "岡本パン";
const RESERVATION_SHEET_NAME = "予約一覧";
const PRODUCT_SHEET_NAME = "商品管理";
const SPREADSHEET_PROPERTY_KEY = "RESERVATION_SPREADSHEET_ID";

const RESERVATION_STATUS_ACTIVE = "予約済";
const RESERVATION_STATUS_CANCELED = "キャンセル";

const PRODUCT_STATUS_ACTIVE = "販売中";
const PRODUCT_STATUS_SOLD_OUT = "完売";
const PRODUCT_STATUS_HIDDEN = "非表示";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("岡本パン 予約サイト")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function getProducts() {
  const spreadsheet = getOrCreateSpreadsheet_();
  const reservationSheet = getOrCreateReservationSheet_(spreadsheet);
  const productSheet = getOrCreateProductSheet_(spreadsheet);

  ensureProductFormulas_(productSheet);
  SpreadsheetApp.flush();

  return readProducts_(productSheet).filter(function(product) {
    return product.status !== PRODUCT_STATUS_HIDDEN;
  });
}

function submitReservation(data) {
  validateBasicReservation_(data);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getOrCreateSpreadsheet_();
    const reservationSheet = getOrCreateReservationSheet_(spreadsheet);
    const productSheet = getOrCreateProductSheet_(spreadsheet);

    ensureProductFormulas_(productSheet);
    SpreadsheetApp.flush();

    const products = readProducts_(productSheet);
    const productMap = {};

    products.forEach(function(product) {
      productMap[String(product.id)] = product;
    });

    const normalizedItems = data.items.map(function(item) {
      const product = productMap[String(item.id)];

      if (!product || product.status === PRODUCT_STATUS_HIDDEN) {
        throw new Error("販売を終了した商品が含まれています。ページを再読み込みしてください。");
      }

      if (product.status === PRODUCT_STATUS_SOLD_OUT || product.remaining <= 0) {
        throw new Error(product.name + "は完売しています。");
      }

      const quantity = Number(item.quantity);
      validateQuantity_(quantity, product.step, product.name);

      if (quantity > product.remaining + 0.000001) {
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
        step: product.step
      };
    });

    const reservationId = createReservationId_();
    const createdAt = new Date();
    const customerName = String(data.customerName || "").trim();
    const customerEmail = String(data.customerEmail || "").trim();
    const memo = String(data.memo || "").trim();

    const rows = normalizedItems.map(function(item) {
      const subtotal = Number(item.price) * Number(item.quantity);

      return [
        createdAt,
        reservationId,
        customerName,
        customerEmail,
        memo,
        item.id,
        item.name,
        item.quantity,
        item.price,
        subtotal,
        RESERVATION_STATUS_ACTIVE
      ];
    });

    const startRow = reservationSheet.getLastRow() + 1;
    reservationSheet
      .getRange(startRow, 1, rows.length, rows[0].length)
      .setValues(rows);

    SpreadsheetApp.flush();
    ensureProductFormulas_(productSheet);
    SpreadsheetApp.flush();

    const total = normalizedItems.reduce(function(sum, item) {
      return sum + item.price * item.quantity;
    }, 0);

    sendConfirmationMail_({
      reservationId: reservationId,
      customerName: customerName,
      customerEmail: customerEmail,
      memo: memo,
      items: normalizedItems,
      total: total
    });

    return {
      success: true,
      reservationId: reservationId,
      products: readProducts_(productSheet).filter(function(product) {
        return product.status !== PRODUCT_STATUS_HIDDEN;
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
    sheet = createReservationSheet_(spreadsheet);
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
 * V2用のシートを作成します。
 *
 * 既存の「商品管理」「予約一覧」は、日時付きのバックアップ名へ変更します。
 * その後、新しい構成の2シートを作ります。
 *
 * Apps Script画面から一度だけ手動実行してください。
 */
function setupV2Sheets() {
  const spreadsheet = getOrCreateSpreadsheet_();
  const timeZone = Session.getScriptTimeZone() || "Asia/Tokyo";
  const suffix = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd_HHmmss");

  backupSheetIfExists_(spreadsheet, PRODUCT_SHEET_NAME, suffix);
  backupSheetIfExists_(spreadsheet, RESERVATION_SHEET_NAME, suffix);

  const reservationSheet = createReservationSheet_(spreadsheet);
  const productSheet = createProductSheet_(spreadsheet);

  spreadsheet.setActiveSheet(productSheet);
  console.log("V2用シートを作成しました: " + spreadsheet.getUrl());

  return spreadsheet.getUrl();
}

function backupSheetIfExists_(spreadsheet, sheetName, suffix) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    return;
  }

  let backupName = sheetName + "_旧_" + suffix;
  let number = 2;

  while (spreadsheet.getSheetByName(backupName)) {
    backupName = sheetName + "_旧_" + suffix + "_" + number;
    number += 1;
  }

  sheet.setName(backupName);
}

function createReservationSheet_(spreadsheet) {
  const sheet = spreadsheet.insertSheet(RESERVATION_SHEET_NAME);

  sheet.appendRow([
    "予約日時",
    "予約番号",
    "お名前",
    "メールアドレス",
    "備考",
    "商品ID",
    "商品名",
    "数量",
    "単価",
    "小計",
    "状態"
  ]);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 11).setFontWeight("bold");
  sheet.getRange("A:A").setNumberFormat("yyyy/mm/dd hh:mm:ss");
  sheet.getRange("H:H").setNumberFormat("0.0");
  sheet.getRange("I:J").setNumberFormat("¥#,##0");

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [RESERVATION_STATUS_ACTIVE, RESERVATION_STATUS_CANCELED],
      true
    )
    .setAllowInvalid(false)
    .build();

  sheet.getRange("K2:K").setDataValidation(statusRule);
  sheet.autoResizeColumns(1, 11);

  return sheet;
}

function createProductSheet_(spreadsheet) {
  const sheet = spreadsheet.insertSheet(PRODUCT_SHEET_NAME);

  sheet.appendRow([
    "商品ID",
    "商品名",
    "価格",
    "製造数",
    "数量単位",
    "予約数合計",
    "残数",
    "特徴",
    "原材料",
    "内容量",
    "保存期間",
    "画像URL1",
    "画像URL2",
    "画像URL3",
    "画像URL4",
    "画像URL5",
    "販売状態",
    "並び順"
  ]);

  const initialProducts = [
    [
      "campagne",
      "カンパーニュ",
      900,
      27,
      0.5,
      "",
      "",
      "静岡県産小麦と自家製酵母で長時間発酵。小麦の香りを楽しめる、岡本パンの定番です。",
      "小麦、塩、自家製酵母",
      "700g",
      "常温で3日程度。食べきれない場合はスライスして冷凍してください。",
      "",
      "",
      "",
      "",
      "",
      PRODUCT_STATUS_ACTIVE,
      1
    ],
    [
      "donut",
      "ドーナッツ",
      350,
      30,
      1,
      "",
      "",
      "ふんわりとした生地を香ばしく揚げたドーナッツです。",
      "小麦、砂糖、卵、乳製品、油、酵母、塩",
      "1個",
      "当日中にお召し上がりください。",
      "",
      "",
      "",
      "",
      "",
      PRODUCT_STATUS_ACTIVE,
      2
    ]
  ];

  sheet
    .getRange(2, 1, initialProducts.length, initialProducts[0].length)
    .setValues(initialProducts);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 18).setFontWeight("bold");
  sheet.getRange("C:C").setNumberFormat("¥#,##0");
  sheet.getRange("D:G").setNumberFormat("0.0");

  const productStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [
        PRODUCT_STATUS_ACTIVE,
        PRODUCT_STATUS_SOLD_OUT,
        PRODUCT_STATUS_HIDDEN
      ],
      true
    )
    .setAllowInvalid(false)
    .build();

  sheet.getRange("Q2:Q").setDataValidation(productStatusRule);

  const stepRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["0.5", "1"], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange("E2:E").setDataValidation(stepRule);

  ensureProductFormulas_(sheet);
  sheet.autoResizeColumns(1, 18);

  return sheet;
}

function ensureProductFormulas_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return;
  }

  for (let row = 2; row <= lastRow; row += 1) {
    const productId = String(sheet.getRange(row, 1).getValue() || "").trim();

    if (!productId) {
      sheet.getRange(row, 6, 1, 2).clearContent();
      continue;
    }

    const reservedFormula =
      '=SUMIFS(\'' + RESERVATION_SHEET_NAME + '\'!$H:$H,' +
      '\'' + RESERVATION_SHEET_NAME + '\'!$F:$F,$A' + row + ',' +
      '\'' + RESERVATION_SHEET_NAME + '\'!$K:$K,"' +
      RESERVATION_STATUS_ACTIVE + '")';

    const remainingFormula = '=MAX(0,$D' + row + '-$F' + row + ')';

    sheet.getRange(row, 6).setFormula(reservedFormula);
    sheet.getRange(row, 7).setFormula(remainingFormula);
  }
}

function readProducts_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 18).getValues();

  return values
    .map(function(row, index) {
      const images = [
        String(row[11] || "").trim(),
        String(row[12] || "").trim(),
        String(row[13] || "").trim(),
        String(row[14] || "").trim(),
        String(row[15] || "").trim()
      ].filter(function(url) {
        return url !== "";
      });

      const step = Number(row[4]) === 1 ? 1 : 0.5;
      const status = normalizeProductStatus_(row[16]);

      return {
        id: String(row[0] || "").trim(),
        name: String(row[1] || "").trim(),
        price: Number(row[2]) || 0,
        production: Number(row[3]) || 0,
        step: step,
        reserved: Number(row[5]) || 0,
        remaining: Math.max(0, Number(row[6]) || 0),
        feature: String(row[7] || ""),
        ingredients: String(row[8] || ""),
        amount: String(row[9] || ""),
        shelfLife: String(row[10] || ""),
        images: images,
        status: status,
        sortOrder: Number(row[17]) || 9999,
        rowNumber: index + 2
      };
    })
    .filter(function(product) {
      return product.id && product.name;
    })
    .sort(function(a, b) {
      return a.sortOrder - b.sortOrder;
    });
}

function normalizeProductStatus_(value) {
  const status = String(value || "").trim();

  if (status === PRODUCT_STATUS_SOLD_OUT) {
    return PRODUCT_STATUS_SOLD_OUT;
  }

  if (status === PRODUCT_STATUS_HIDDEN) {
    return PRODUCT_STATUS_HIDDEN;
  }

  return PRODUCT_STATUS_ACTIVE;
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

function validateQuantity_(quantity, step, productName) {
  if (!Number.isFinite(quantity) || quantity <= 0) {
    throw new Error(productName + "の予約数量が正しくありません。");
  }

  const ratio = quantity / step;

  if (Math.abs(ratio - Math.round(ratio)) > 0.000001) {
    throw new Error(
      productName + "は" + formatQuantity_(step) + "個刻みで選択してください。"
    );
  }
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
