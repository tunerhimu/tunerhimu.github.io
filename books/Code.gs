/**
 * ===================================================================
 *  My Kid BD Shop — অর্ডার ম্যানেজমেন্ট ব্যাকএন্ড (Google Apps Script)
 * ===================================================================
 *  এই কোডটি আপনার Google Sheet-এর সাথে যুক্ত Apps Script এডিটরে
 *  হুবহু কপি-পেস্ট করুন। বিস্তারিত গাইডলাইন সাথে দেওয়া
 *  "গুগল-শীট-কানেকশন-গাইড.md" ফাইলে দেওয়া আছে।
 * ===================================================================
 */

const SHEET_NAME = 'Orders';

/** ওয়েবসাইট থেকে নতুন অর্ডার এলে এখানে POST রিকোয়েস্ট আসে */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'newOrder') {
      return saveOrder(body.order);
    }

    return jsonResponse({ success: false, message: 'Unknown action: ' + body.action });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

/** ট্র্যাকিং মোডাল থেকে GET রিকোয়েস্ট আসে (?action=track&query=...) */
function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'track') {
      return trackOrder(e.parameter.query);
    }

    // ব্রাউজারে সরাসরি URL খুললে একটা ছোট্ট স্ট্যাটাস মেসেজ দেখাবে
    return jsonResponse({ success: true, message: 'My Kid BD Shop Order API is running ✅' });
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

/** Orders শীট খুঁজে বের করে, না থাকলে হেডারসহ নতুন করে তৈরি করে */
function getOrdersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      'Timestamp', 'OrderID', 'Name', 'Phone', 'Address', 'Product',
      'Quantity', 'UnitPrice', 'DeliveryLocation', 'DeliveryFee',
      'Subtotal', 'Total', 'Date', 'Status'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4f46e5')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);

    // Status কলামে ড্রপডাউন যোগ করা হলো, যাতে ম্যানুয়ালি আপডেট করা সহজ হয়
    const statusRange = sheet.getRange(2, 14, 500, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['পেন্ডিং', 'প্রসেসিং', 'শিপড', 'ডেলিভারড'], true)
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(rule);
  }

  return sheet;
}

/** নতুন অর্ডার শীটে যোগ করে */
function saveOrder(order) {
  const sheet = getOrdersSheet();

  sheet.appendRow([
    new Date(),
    order.orderId,
    order.name,
    order.phone,
    order.address,
    order.product,
    order.quantity,
    order.unitPrice,
    order.deliveryLocation,
    order.deliveryFee,
    order.subtotal,
    order.total,
    order.date,
    order.status || 'পেন্ডিং'
  ]);

  return jsonResponse({ success: true, orderId: order.orderId });
}

/** অর্ডার আইডি বা ফোন নম্বর দিয়ে সর্বশেষ ম্যাচিং অর্ডার খুঁজে বের করে */
function trackOrder(query) {
  if (!query) {
    return jsonResponse({ found: false, message: 'query missing' });
  }

  const sheet = getOrdersSheet();
  const data = sheet.getDataRange().getValues();
  const q = String(query).trim().toLowerCase();

  // সবচেয়ে নতুন অর্ডার আগে খুঁজবে (নিচ থেকে উপরে)
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const orderId = String(row[1]).toLowerCase();
    const phone = String(row[3]).trim();

    if (orderId === q || phone === query.trim()) {
      const order = {
        orderId: row[1],
        name: row[2],
        phone: row[3],
        address: row[4],
        product: row[5],
        quantity: row[6],
        unitPrice: row[7],
        deliveryLocation: row[8],
        deliveryFee: row[9],
        subtotal: row[10],
        total: row[11],
        date: row[12],
        status: row[13]
      };
      return jsonResponse({ found: true, order: order });
    }
  }

  return jsonResponse({ found: false });
}

/** JSON রেসপন্স তৈরির হেল্পার ফাংশন */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
