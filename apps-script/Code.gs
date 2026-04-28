// ============================================================
// DC&S Headstone Order Tracker — Google Apps Script
// Full rebuild — includes Stripe webhook + payment-success URL
// ============================================================

const SHEET_NAME = "Orders";
const SHEET_ID = "1_pbyhKL1IElgneBZHFIWG76Cv6hNKdUw0fgSHQjMN6M";
const DRIVE_FOLDER_ID = "1nAxdUKug-s3pEQnX9RCps86crK--Vd4k";
const PRICE_BOOK_ID = SHEET_ID;

// ── STRIPE KEY stored in Apps Script Script Properties ──
// Project Settings → Script Properties → STRIPE_SECRET_KEY
const STRIPE_SECRET_KEY = PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') || '';

// ── STONE MASON CONTACT ──
// Where the auto-notification email is sent when a proof is approved.
// Override at runtime by setting Script Property MASON_EMAIL / MASON_NAME.
const MASON_EMAIL = PropertiesService.getScriptProperties().getProperty('MASON_EMAIL') || 'mason@example.com';
const MASON_NAME  = PropertiesService.getScriptProperties().getProperty('MASON_NAME')  || 'Stone Mason';

const HEADERS = [
  "Order ID", "Created", "Last Updated", "Status", "Payment Status",
  "Customer Name", "Phone", "Email", "Address",
  "Deceased Name", "Date of Birth", "Date of Passing", "Order Date",
  "Headstone Type", "Headstone Size", "Headstone Colour", "Headstone Colour Adj", "Headstone Finish",
  "Headstone Sell Price", "Headstone Cost Price",
  "Surround Type", "Granite Upgrade", "Surround Sell Price", "Surround Cost Price",
  "Stone / Chippings", "Stone Sell Price", "Stone Cost Price",
  "Accessories", "Accessories Sell Price", "Accessories Cost Price",
  "Inscription Type", "Inscription Text", "Inscription Lines", "Letter Style", "Inscription Colour",
  "Inscription Sell Price", "Inscription Cost Price",
  "Cemetery / Location", "Cemetery Fee",
  "Additional Services", "Services Sell Price", "Services Cost Price",
  "Total Sell Price", "Total Cost Price", "Profit Margin", "Margin Percentage",
  "Deposit Paid", "Balance Due",
  "Proof Date", "Proof Approved", "Production Start", "Install Date",
  "Artwork Notes", "General Notes", "Mason Notes",
  "Folder Link", "Files", "Extra Charges", "Mason Charges",
  "Log Entries",
  "Note Entries", "Mason Note Entries",
  "Stripe Link ID", "Stripe Payment Date", "Stripe Payment Amount",
  "Mason Notified At", "Mason Notified By"
];

// ============================================================
// SHEET HELPERS
// ============================================================
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setBackground("#1e2530");
    headerRange.setFontColor("#b89a5e");
    headerRange.setFontWeight("bold");
    headerRange.setFontFamily("Arial");
    headerRange.setFontSize(9);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaders(sheet);
  }
  return sheet;
}

// Append any HEADERS that are missing from row 1 so new fields persist
// instead of being silently dropped by upsertOrder's header-driven mapping.
function ensureHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  const existing = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
    : [];
  const missing = HEADERS.filter(h => existing.indexOf(h) === -1);
  if (missing.length === 0) return;
  const startCol = existing.length + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
  const newRange = sheet.getRange(1, startCol, 1, missing.length);
  newRange.setBackground("#1e2530");
  newRange.setFontColor("#b89a5e");
  newRange.setFontWeight("bold");
  newRange.setFontFamily("Arial");
  newRange.setFontSize(9);
}

function respond(success, message, data) {
  const result = { success, message };
  if (data !== undefined) result.data = data;
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function styleDataRow(sheet, rowNum, status) {
  const statusColors = {
    enquiry: "#f1f5f9", quoted: "#fef9e7", confirmed: "#dbeafe",
    design: "#ede9fe", production: "#fff7ed", ready: "#d1fae5", installed: "#dcfce7"
  };
  const bg = statusColors[(status || '').toLowerCase()] || "#ffffff";
  const range = sheet.getRange(rowNum, 1, 1, HEADERS.length);
  range.setBackground(bg);
  range.setFontFamily("Arial");
  range.setFontSize(9);
  range.setVerticalAlignment("middle");
}

// ============================================================
// doGet — load all orders + price book (JSONP supported)
// ============================================================
function doGet(e) {
  // ── Customer proof page: public, limited data, no login ──
  if (e && e.parameter && e.parameter.action === 'getProofData') {
    return getProofData(e.parameter.id, e.parameter.callback);
  }

  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const priceBook = loadPriceBook();
    const orders = [];

    if (data.length > 1) {
      const headers = data[0];
      data.slice(1)
        .filter(row => row[0] && row[0] !== 'SETUP_ROW' && !String(row[0]).startsWith('SETUP'))
        .forEach(row => {
          const sheetOrder = {};
          headers.forEach((h, i) => {
            sheetOrder[h] = row[i] !== undefined ? String(row[i]) : '';
          });
          orders.push(mapSheetOrderToTracker(sheetOrder));
        });
    }

    const result = JSON.stringify({ success: true, orders, priceBook });
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + result + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const result = JSON.stringify({ success: false, message: err.toString() });
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + result + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// doPost — route actions and Stripe webhooks
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    // Stripe webhooks have a 'type' field; app actions have an 'action' field
    if (body.type) return handleStripeWebhook(body);
    const action = body.action || "upsert";
    if (action === "upsert")              return upsertOrder(body.order);
    if (action === "delete")              return deleteOrder(body.orderId);
    if (action === "uploadFile")          return uploadFileToDrive(body);
    if (action === "deleteFile")          return deleteFileFromDrive(body.fileId);
    if (action === "createPaymentLink")   return handleCreatePaymentLink(body);
    if (action === "submitProofResponse") return submitProofResponse(body.orderId, body.approved, body.message);
    if (action === "sendEstimateEmail")   return sendEstimateEmail(body.email, body.customerName, body.ref, body.pdfBase64, body.proofUrl);
    if (action === "storeEstimatePdf")    return storeEstimatePdf(body.orderId, body.ref, body.pdfBase64);
    if (action === "notifyMason")         return notifyMason(body.orderId, body.triggeredBy, !!body.force);
    return respond(false, "Unknown action");
  } catch (err) {
    return respond(false, err.toString());
  }
}

// ============================================================
// STRIPE WEBHOOK HANDLER
// Stripe sends checkout.session.completed → update order in sheet
// ============================================================
function handleStripeWebhook(event) {
  try {
    Logger.log('Stripe webhook received: ' + event.type);
    if (event.type === 'checkout.session.completed') {
      const session = event.data && event.data.object;
      if (session) {
        const orderId      = session.metadata && session.metadata.order_id;
        const amountPounds = (session.amount_total || 0) / 100;
        const paymentType  = (session.metadata && session.metadata.payment_type) || 'payment';
        if (orderId) {
          markPaymentReceived(orderId, amountPounds, paymentType);
        } else {
          Logger.log('Webhook: no order_id in metadata — cannot match order');
        }
      }
    }
    // Always return 200 to acknowledge receipt
    return ContentService
      .createTextOutput(JSON.stringify({ received: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('Webhook error: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ received: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function markPaymentReceived(orderId, amountPounds, paymentType) {
  try {
    const sheet   = getOrCreateSheet();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(orderId)) continue;
      const rowNum = i + 1;

      // Add to deposit paid
      const depCol = col('Deposit Paid');
      if (depCol >= 0) {
        const current = parseFloat(data[i][depCol]) || 0;
        const newDeposit = current + amountPounds;
        sheet.getRange(rowNum, depCol + 1).setValue(newDeposit);

        // Update balance due
        const totalCol = col('Total Sell Price');
        const balCol   = col('Balance Due');
        if (totalCol >= 0 && balCol >= 0) {
          const total  = parseFloat(data[i][totalCol]) || 0;
          const newBal = Math.max(0, total - newDeposit);
          sheet.getRange(rowNum, balCol + 1).setValue(newBal);

          // Update payment status
          const psCol = col('Payment Status');
          if (psCol >= 0) {
            sheet.getRange(rowNum, psCol + 1).setValue(newBal <= 0 ? 'Paid' : 'Part Paid');
          }
        }
      }

      // Record Stripe payment details
      const sdCol = col('Stripe Payment Date');
      if (sdCol >= 0) sheet.getRange(rowNum, sdCol + 1).setValue(new Date().toLocaleString('en-GB'));
      const saCol = col('Stripe Payment Amount');
      if (saCol >= 0) sheet.getRange(rowNum, saCol + 1).setValue(amountPounds);

      // Append to log
      const logCol = col('Log Entries');
      if (logCol >= 0) {
        const existing = String(data[i][logCol] || '');
        const ts = new Date().toLocaleDateString('en-GB') + ' ' +
                   new Date().toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
        const entry = '[' + ts + '] System: \uD83D\uDCB3 Stripe payment received \u2014 \u00A3' +
                      amountPounds.toFixed(2) + ' (' + paymentType + ')';
        sheet.getRange(rowNum, logCol + 1).setValue(existing ? existing + ' | ' + entry : entry);
      }

      Logger.log('Payment marked for order: ' + orderId + ' \u2014 \u00A3' + amountPounds);
      return;
    }
    Logger.log('Order not found for webhook orderId: ' + orderId);
  } catch (err) {
    Logger.log('markPaymentReceived error: ' + err);
  }
}

// ============================================================
// STRIPE PAYMENT LINK CREATOR
// Creates Stripe product + price + payment link with success URL
// that shows balance details on payment-success.html
// ============================================================
function handleCreatePaymentLink(data) {
  try {
    const { orderId, customerName, deceasedName, orderRef, amountPence, description } = data;

    if (!STRIPE_SECRET_KEY || STRIPE_SECRET_KEY.startsWith("PASTE_") || STRIPE_SECRET_KEY === '') {
      return respond(false, "Stripe secret key not configured in Script Properties");
    }
    if (!amountPence || amountPence < 100) {
      return respond(false, "Amount must be at least \u00A31.00");
    }

    const auth = "Basic " + Utilities.base64Encode(STRIPE_SECRET_KEY + ":");

    // 1 — Create a one-time Product for this order
    const productPayload = {
      "name": "DC&S Memorial \u2014 " + (deceasedName || "Order") + " #" + (orderRef || orderId).slice(-8).toUpperCase(),
      "description": description || ("Memorial headstone order for " + (customerName || "Customer"))
    };
    const productResp = UrlFetchApp.fetch("https://api.stripe.com/v1/products", {
      method: "post",
      headers: { "Authorization": auth, "Content-Type": "application/x-www-form-urlencoded" },
      payload: encodeStripePayload(productPayload),
      muteHttpExceptions: true
    });
    const product = JSON.parse(productResp.getContentText());
    if (!product.id) return respond(false, "Stripe product error: " + productResp.getContentText());

    // 2 — Create a Price for that product
    const pricePayload = {
      "unit_amount": String(Math.round(amountPence)),
      "currency": "gbp",
      "product": product.id
    };
    const priceResp = UrlFetchApp.fetch("https://api.stripe.com/v1/prices", {
      method: "post",
      headers: { "Authorization": auth, "Content-Type": "application/x-www-form-urlencoded" },
      payload: encodeStripePayload(pricePayload),
      muteHttpExceptions: true
    });
    const price = JSON.parse(priceResp.getContentText());
    if (!price.id) return respond(false, "Stripe price error: " + priceResp.getContentText());

    // 3 — Build customer-facing success URL with balance details
    const amountPounds = amountPence / 100;
    const totalOrder   = parseFloat(data.totalSellPrice)   || 0;
    const prevDeposit  = parseFloat(data.previousDeposit)  || 0;
    const newBalance   = Math.max(0, totalOrder - prevDeposit - amountPounds);
    const shortRef     = (orderRef || orderId || "").slice(-8).toUpperCase();
    const successUrl   = "https://andrewcrymble.github.io/dcfs-memorial-tracker/payment-success.html"
      + "?ref="  + encodeURIComponent(shortRef)
      + "&amt="  + amountPounds.toFixed(2)
      + "&bal="  + newBalance.toFixed(2)
      + "&name=" + encodeURIComponent(deceasedName || "")
      + "&cust=" + encodeURIComponent(customerName || "");

    Logger.log("Success URL: " + successUrl);

    // 4 — Create the Payment Link
    const linkPayload = {
      "line_items[0][price]": price.id,
      "line_items[0][quantity]": "1",
      "after_completion[type]": "redirect",
      "after_completion[redirect][url]": successUrl,
      "metadata[order_id]": orderId || "",
      "metadata[order_ref]": shortRef,
      "metadata[payment_type]": data.paymentType || "payment",
      "metadata[customer]": customerName || ""
    };
    const linkResp = UrlFetchApp.fetch("https://api.stripe.com/v1/payment_links", {
      method: "post",
      headers: { "Authorization": auth, "Content-Type": "application/x-www-form-urlencoded" },
      payload: encodeStripePayload(linkPayload),
      muteHttpExceptions: true
    });
    const link = JSON.parse(linkResp.getContentText());
    if (!link.url) return respond(false, "Stripe link error: " + linkResp.getContentText());

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, url: link.url, linkId: link.id }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return respond(false, "Stripe error: " + err.toString());
  }
}

function encodeStripePayload(obj) {
  return Object.keys(obj).map(k =>
    encodeURIComponent(k) + "=" + encodeURIComponent(obj[k])
  ).join("&");
}

// ============================================================
// UPSERT ORDER (create or update row in sheet)
// ============================================================
function upsertOrder(order) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  let existingRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(order.id || order.orderId)) {
      existingRow = i + 1;
      break;
    }
  }

  const balance = (parseFloat(order.totalSellPrice) || 0) - (parseFloat(order.depositPaid) || 0);
  const logText = (order.log || []).map(l =>
    '[' + new Date(l.ts).toLocaleDateString("en-GB") + ' ' +
    new Date(l.ts).toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" }) +
    '] ' + (l.author || "Staff") + ': ' + l.text
  ).join(" | ");

  const fmtDate = d => {
    if (!d || d === "Invalid Date") return "";
    try {
      const dt = new Date(d);
      if (isNaN(dt.getTime())) return "";
      return dt.toLocaleDateString("en-GB");
    } catch (e) { return ""; }
  };

  const fmtDateTime = () =>
    new Date().toLocaleDateString("en-GB") + " " +
    new Date().toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });

  const valueMap = {
    "Order ID":            order.id || order.orderId || "",
    "Created": (function () {
      const v = order.createdAt || order.created;
      if (!v || v === "Invalid Date") return fmtDateTime();
      try {
        const d = new Date(v);
        if (isNaN(d.getTime())) return fmtDateTime();
        return d.toLocaleDateString("en-GB") + " " + d.toLocaleTimeString("en-GB", { hour: "2-digit", minute: "2-digit" });
      } catch (e) { return fmtDateTime(); }
    })(),
    "Last Updated":        fmtDateTime(),
    "Status":              order.status || "enquiry",
    "Payment Status":      order.paymentStatus || "Unpaid",
    "Customer Name":       order.customerName || "",
    "Phone":               order.phone || "",
    "Email":               order.email || "",
    "Address":             order.address || "",
    "Deceased Name":       order.deceasedName || "",
    "Date of Birth":       fmtDate(order.deceasedDob),
    "Date of Passing":     fmtDate(order.deceasedDod),
    "Order Date":          fmtDate(order.orderDate),
    "Headstone Type":      order.hsType || "",
    "Headstone Size":      order.hsSize || "",
    "Headstone Colour":    order.hsColour || "",
    "Headstone Colour Adj": order.hsColourAdj || 0,
    "Headstone Finish":    order.hsFinish || "",
    "Headstone Sell Price": order.hsSellPrice || 0,
    "Headstone Cost Price": order.hsCostPrice || 0,
    "Sell Price":          order.hsSellPrice || 0,
    "Surround Type":       order.surroundType || "",
    "Granite Upgrade":     order.surroundGranite ? "Yes" : "No",
    "Surround Sell Price": order.surroundSellPrice || 0,
    "Surround Cost Price": order.surroundCostPrice || 0,
    "Stone / Chippings":   order.stoneType || "",
    "Stone Sell Price":    order.stoneSellPrice || 0,
    "Stone Cost Price":    order.stoneCostPrice || 0,
    "Accessories":         (order.accessories || []).join(", "),
    "Accessories Sell Price": order.accessoriesSellPrice || 0,
    "Accessories Cost Price": order.accessoriesCostPrice || 0,
    "Inscription Type":    order.inscriptionType === "additional" ? "Additional on Existing" : "New Inscription",
    "Inscription Text":    order.inscriptionText || "",
    "Inscription Lines":   order.inscriptionLines || 0,
    "Letter Style":        order.inscriptionStyle || "",
    "Inscription Colour":  order.inscriptionColour || "Silver",
    "Inscription Sell Price": order.inscriptionSellPrice || 0,
    "Inscription Cost Price": order.inscriptionCostPrice || 0,
    "Inscription Charge":  order.inscriptionSellPrice || 0,
    "Price Per Line":      order.inscriptionPpl || 35,
    "Cemetery / Location": order.cemetery || "",
    "Cemetery Fee":        order.cemeteryFee || 0,
    "Additional Services": order.additionalServices || "",
    "Services Sell Price": order.servicesSellPrice || 0,
    "Services Cost Price": order.servicesCostPrice || 0,
    "Total Sell Price":    order.totalSellPrice || 0,
    "Total Cost Price":    order.totalCostPrice || 0,
    "Total Price":         order.totalSellPrice || 0,
    "Profit Margin":       order.profitMargin || 0,
    "Margin Percentage":   order.marginPercentage || 0,
    "Deposit Paid":        order.depositPaid || 0,
    "Balance Due":         Math.max(0, balance),
    "Proof Date":          fmtDate(order.proofDate),
    "Proof Approved":      order.artworkApproved ? "Yes" : "No",
    "Production Start":    fmtDate(order.productionDate),
    "Install Date":        fmtDate(order.installDate),
    "Artwork Notes":       order.artworkNotes || "",
    "General Notes":       order.notes || "",
    "Mason Notes":         order.masonNotes || "",
    "Folder Link":         order.folderLink || "",
    "Files":               order.files ? JSON.stringify(order.files) : "[]",
    "Extra Charges":       order.extraCharges ? JSON.stringify(order.extraCharges) : "[]",
    "Mason Charges":       order.masonCharges ? JSON.stringify(order.masonCharges) : "[]",
    "Log Entries":         logText,
    "Note Entries":        order.noteEntries ? JSON.stringify(order.noteEntries) : "[]",
    "Mason Note Entries":  order.masonNoteEntries ? JSON.stringify(order.masonNoteEntries) : "[]",
    "Stripe Link ID":      order.stripePaymentUrl || order.stripeLinkId || "",
    "Stripe Payment Date": order.stripePaymentDate || "",
    "Stripe Payment Amount": order.stripePaymentAmount || 0,
    "Mason Notified At":   order.masonNotifiedAt || "",
    "Mason Notified By":   order.masonNotifiedBy || "",
  };

  const row = headers.map(h => valueMap.hasOwnProperty(h) ? valueMap[h] : "");

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    styleDataRow(sheet, existingRow, order.status);
    return respond(true, "updated");
  } else {
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, row.length).setValues([row]);
    styleDataRow(sheet, newRow, order.status);
    return respond(true, "created");
  }
}

function deleteOrder(orderId) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === orderId) {
      sheet.deleteRow(i + 1);
      return respond(true, "deleted");
    }
  }
  return respond(false, "Order not found");
}

// ============================================================
// FILE UPLOAD / DELETE (Google Drive)
// ============================================================
function getOrCreateOrderFolder(orderId, customerName, deceasedName) {
  const rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const folderName = (customerName || "Unknown") + " \u2014 " + (deceasedName || "Memorial") +
                     " \u2014 #" + String(orderId).slice(-6).toUpperCase();
  const existing = rootFolder.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return rootFolder.createFolder(folderName);
}

function uploadFileToDrive(data) {
  try {
    const { orderId, customerName, deceasedName, fileName, fileType, fileData, mimeType } = data;
    const folder = getOrCreateOrderFolder(orderId, customerName, deceasedName);
    const decoded = Utilities.base64Decode(fileData);
    const blob = Utilities.newBlob(decoded, mimeType || "application/octet-stream", fileName);
    const file = folder.createFile(blob);
    file.setDescription(fileType || "");
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        fileId: file.getId(),
        fileName: file.getName(),
        viewUrl: "https://drive.google.com/file/d/" + file.getId() + "/view",
        folderUrl: "https://drive.google.com/drive/folders/" + folder.getId(),
        folderId: folder.getId(),
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return respond(false, "Upload failed: " + err.toString());
  }
}

function deleteFileFromDrive(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return respond(true, "deleted");
  } catch (err) {
    return respond(false, err.toString());
  }
}

// ============================================================
// MAP SHEET ROW → TRACKER OBJECT
// ============================================================
function mapSheetOrderToTracker(sheetOrder) {
  return {
    orderId:             sheetOrder["Order ID"],
    orderRef:            sheetOrder["Order ID"],
    created:             sheetOrder["Created"],
    lastUpdated:         sheetOrder["Last Updated"],
    orderDate:           sheetOrder["Order Date"],
    status:              capitalizeStatus(sheetOrder["Status"]),
    paymentStatus:       sheetOrder["Payment Status"],
    customerName:        sheetOrder["Customer Name"],
    phone:               sheetOrder["Phone"],
    email:               sheetOrder["Email"],
    address:             sheetOrder["Address"],
    deceasedName:        sheetOrder["Deceased Name"],
    deceasedDob:         sheetOrder["Date of Birth"],
    deceasedDod:         sheetOrder["Date of Passing"],
    hsType:              sheetOrder["Headstone Type"],
    hsSize:              sheetOrder["Headstone Size"],
    hsColour:            sheetOrder["Headstone Colour"] || "",
    hsColourAdj:         parseFloat(sheetOrder["Headstone Colour Adj"]) || 0,
    hsFinish:            sheetOrder["Headstone Finish"] || "",
    hsSellPrice:         parseFloat(sheetOrder["Headstone Sell Price"] || sheetOrder["Sell Price"]) || 0,
    hsCostPrice:         parseFloat(sheetOrder["Headstone Cost Price"]) || 0,
    surroundType:        sheetOrder["Surround Type"],
    surroundGranite:     sheetOrder["Granite Upgrade"] === "Yes",
    surroundSellPrice:   parseFloat(sheetOrder["Surround Sell Price"]) || 0,
    surroundCostPrice:   parseFloat(sheetOrder["Surround Cost Price"]) || 0,
    stoneType:           sheetOrder["Stone / Chippings"],
    stoneSellPrice:      parseFloat(sheetOrder["Stone Sell Price"]) || 0,
    stoneCostPrice:      parseFloat(sheetOrder["Stone Cost Price"]) || 0,
    accessories:         sheetOrder["Accessories"] && sheetOrder["Accessories"] !== "No"
                           ? sheetOrder["Accessories"].split(",").map(a => a.trim())
                           : [],
    accessoriesSellPrice: parseFloat(sheetOrder["Accessories Sell Price"]) || 0,
    accessoriesCostPrice: parseFloat(sheetOrder["Accessories Cost Price"]) || 0,
    inscriptionType:     sheetOrder["Inscription Type"] === "Additional on Existing" ? "additional" : "new",
    inscriptionText:     sheetOrder["Inscription Text"],
    inscriptionLines:    parseInt(sheetOrder["Inscription Lines"]) || 0,
    inscriptionStyle:    sheetOrder["Letter Style"],
    inscriptionColour:   sheetOrder["Inscription Colour"] || "",
    inscriptionPpl:      parseFloat(sheetOrder["Price Per Line"]) || 35,
    inscriptionSellPrice: parseFloat(sheetOrder["Inscription Sell Price"] || sheetOrder["Inscription Charge"]) || 0,
    inscriptionCostPrice: parseFloat(sheetOrder["Inscription Cost Price"]) || 0,
    cemetery:            sheetOrder["Cemetery / Location"],
    cemeteryFee:         parseFloat(sheetOrder["Cemetery Fee"]) || 0,
    additionalServices:  sheetOrder["Additional Services"] || "",
    servicesSellPrice:   parseFloat(sheetOrder["Services Sell Price"]) || 0,
    servicesCostPrice:   parseFloat(sheetOrder["Services Cost Price"]) || 0,
    totalSellPrice:      parseFloat(sheetOrder["Total Sell Price"] || sheetOrder["Total Price"]) || 0,
    totalCostPrice:      parseFloat(sheetOrder["Total Cost Price"]) || 0,
    profitMargin:        parseFloat(sheetOrder["Profit Margin"]) || 0,
    marginPercentage:    parseFloat(sheetOrder["Margin Percentage"]) || 0,
    depositPaid:         parseFloat(sheetOrder["Deposit Paid"]) || 0,
    balanceDue:          parseFloat(sheetOrder["Balance Due"]) || 0,
    proofDate:           sheetOrder["Proof Date"],
    artworkApproved:     sheetOrder["Proof Approved"] === "Yes",
    productionDate:      sheetOrder["Production Start"],
    installDate:         sheetOrder["Install Date"],
    artworkNotes:        sheetOrder["Artwork Notes"],
    notes:               sheetOrder["General Notes"],
    masonNotes:          sheetOrder["Mason Notes"],
    folderLink:          sheetOrder["Folder Link"],
    files:               parseJSON(sheetOrder["Files"]),
    extraCharges:        parseJSON(sheetOrder["Extra Charges"]),
    masonCharges:        parseJSON(sheetOrder["Mason Charges"]),
    noteEntries:         parseJSON(sheetOrder["Note Entries"]),
    masonNoteEntries:    parseJSON(sheetOrder["Mason Note Entries"]),
    stripeLinkId:        sheetOrder["Stripe Link ID"] || "",
    stripePaymentUrl:    (sheetOrder["Stripe Link ID"] || "").startsWith("http") ? sheetOrder["Stripe Link ID"] : "",
    stripePaymentDate:   sheetOrder["Stripe Payment Date"] || "",
    stripePaymentAmount: parseFloat(sheetOrder["Stripe Payment Amount"]) || 0,
    stripePaymentReceived: !!(sheetOrder["Stripe Payment Date"] && sheetOrder["Stripe Payment Date"] !== ""),
    masonNotifiedAt:     sheetOrder["Mason Notified At"] || "",
    masonNotifiedBy:     sheetOrder["Mason Notified By"] || "",
    log:                 parseLogEntries(sheetOrder["Log Entries"])
  };
}

function capitalizeStatus(status) {
  if (!status) return "Enquiry";
  const statusMap = {
    "enquiry": "Enquiry", "quoted": "Quoted", "confirmed": "Confirmed",
    "design": "In Design", "in design": "In Design", "production": "Production",
    "ready": "Ready", "installed": "Installed"
  };
  return statusMap[status.toLowerCase()] || "Enquiry";
}

function parseJSON(str) {
  if (!str || str === "") return [];
  try { return JSON.parse(str); } catch (e) { return []; }
}

function parseLogEntries(logString) {
  if (!logString || logString === "") return [];
  try {
    return logString.split(" | ").map((entry, index) => {
      const match = entry.match(/\[(.*?)\]\s*(.*?):\s*(.*)/);
      if (match) {
        return { id: Date.now() + index, ts: match[1], author: match[2], text: match[3] };
      }
      return null;
    }).filter(Boolean);
  } catch (e) { return []; }
}

// ============================================================
// PRICE BOOK LOADER
// ============================================================
function loadPriceBook() {
  const ss = SpreadsheetApp.openById(PRICE_BOOK_ID);
  const priceBook = {};

  try {
    const sh = ss.getSheetByName("Headstones");
    if (sh) {
      const d = sh.getDataRange().getValues();
      priceBook.Headstones = d.slice(1).filter(r => r[0]).map(r => ({
        type: r[0], size: r[1], cost: parseFloat(r[2]) || 0,
        sell: parseFloat(r[3]) || 0, margin: parseFloat(r[4]) || 0, marginPct: parseFloat(r[5]) || 0
      }));
    }
  } catch (e) { Logger.log('Headstones tab error: ' + e); }

  try {
    const sh = ss.getSheetByName("Headstone_Colours");
    if (sh) {
      const d = sh.getDataRange().getValues();
      priceBook.Headstone_Colours = d.slice(1).filter(r => r[0]).map(r => ({
        name: r[0], costAdj: parseFloat(r[1]) || 0, sellAdj: parseFloat(r[2]) || 0, margin: parseFloat(r[3]) || 0
      }));
      // Also expose as 'Colours' for backward compat
      priceBook.Colours = priceBook.Headstone_Colours;
    }
  } catch (e) { Logger.log('Colours tab error: ' + e); }

  try {
    const sh = ss.getSheetByName("Surrounds");
    if (sh) {
      const d = sh.getDataRange().getValues();
      priceBook.Surrounds = d.slice(1).filter(r => r[0]).map(r => ({
        type: r[0], baseCost: parseFloat(r[1]) || 0, baseSell: parseFloat(r[2]) || 0,
        graniteCostAdd: parseFloat(r[3]) || 0, graniteSellAdd: parseFloat(r[4]) || 0,
        baseMargin: parseFloat(r[5]) || 0, graniteMargin: parseFloat(r[6]) || 0
      }));
    }
  } catch (e) { Logger.log('Surrounds tab error: ' + e); }

  try {
    const sh = ss.getSheetByName("Stones");
    if (sh) {
      const d = sh.getDataRange().getValues();
      priceBook.Stones = d.slice(1).filter(r => r[0]).map(r => ({
        type: r[0], standaloneCost: parseFloat(r[1]) || 0,
        withSurroundCost: parseFloat(r[2]) || 0, sell: parseFloat(r[3]) || 0
      }));
    }
  } catch (e) { Logger.log('Stones tab error: ' + e); }

  try {
    const sh = ss.getSheetByName("Accessories");
    if (sh) {
      const d = sh.getDataRange().getValues();
      priceBook.Accessories = d.slice(1).filter(r => r[0]).map(r => ({
        name: r[0], size: r[1], cost: parseFloat(r[2]) || 0, sell: parseFloat(r[3]) || 0
      }));
    }
  } catch (e) { Logger.log('Accessories tab error: ' + e); }

  try {
    const sh = ss.getSheetByName("Cemetery_Fees");
    if (sh) {
      const d = sh.getDataRange().getValues();
      priceBook.Cemetery_Fees = d.slice(1).filter(r => r[0]).map(r => ({
        location: r[0], fee: parseFloat(r[1]) || 0
      }));
    }
  } catch (e) { Logger.log('Cemetery tab error: ' + e); }

  try {
    const sh = ss.getSheetByName("Additional_Services");
    if (sh) {
      const d = sh.getDataRange().getValues();
      priceBook.Services = d.slice(1).filter(r => r[0]).map(r => ({
        name: r[0], cost: parseFloat(r[1]) || 0, sell: parseFloat(r[2]) || 0
      }));
    }
  } catch (e) { Logger.log('Services tab error: ' + e); }

  priceBook.NewInscription = {
    freeLetter: 100,
    costPerLetterAfter100: 2.00,
    sellPerLetterAfter100: 3.00
  };

  priceBook.AdditionalInscription = {
    costPerLetterAfter50: 3.00,
    sellPerLetterAfter50: 4.50
  };

  return priceBook;
}

// ============================================================
// SETUP FUNCTIONS (run once from Apps Script editor)
// ============================================================
function setupPriceBook() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  function getOrCreate(name) {
    return ss.getSheetByName(name) || ss.insertSheet(name);
  }

  function writeTab(sheet, headers, rows) {
    sheet.clearContents();
    const allRows = [headers, ...rows];
    sheet.getRange(1, 1, allRows.length, headers.length).setValues(allRows);
    const hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setBackground('#1e3a5f');
    hdr.setFontColor('#ffffff');
    hdr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  writeTab(getOrCreate('Headstones'),
    ['Type', 'Size', 'Cost (£)', 'Sell (£)', 'Margin (£)', 'Margin (%)'],
    [
      ['Ogee', '1.9ft (Base 2ft)',              750,  900,  150, 16.7],
      ['Ogee', '2ft (Base 2.6ft)',             1000, 1150,  150, 13.0],
      ['Ogee', '2.6ft (Base 3ft)',             1200, 1400,  200, 14.3],
      ['Ogee', '3ft (Base 3.6ft)',             1400, 1600,  200, 12.5],
      ['Ogee', '3.6ft (Base 4ft)',             1600, 1800,  200, 11.1],
      ['G3',   '1.9ft (Base 2ft)',              750,  900,  150, 16.7],
      ['G3',   '2ft (Base 2.6ft)',             1000, 1150,  150, 13.0],
      ['G3',   '2.6ft (Base 3ft)',             1200, 1400,  200, 14.3],
      ['G3',   '3ft (Base 3.6ft)',             1400, 1600,  200, 12.5],
      ['G3',   '3.6ft (Base 4ft)',             1600, 1800,  200, 11.1],
      ['Denmore',      '3ft (Base 3.6ft)',     1400, 1600,  200, 12.5],
      ['Denmore',      '3.6ft (Base 4ft)',     1600, 1800,  200, 11.1],
      ['Half Denmore', '2ft (Base 2.6ft)',     1000, 1150,  150, 13.0],
      ['Half Denmore', '2.6ft (Base 3ft)',     1200, 1400,  200, 14.3],
      ['Half Denmore', '3ft (Base 3.6ft)',     1400, 1600,  200, 12.5],
      ['Murphy', '36"x30" / Base 42"x12"x5"', 1400, 2600, 1200, 46.2],
    ]
  );

  writeTab(getOrCreate('Headstone_Colours'),
    ['Colour Name', 'Cost Adjustment (£)', 'Sell Adjustment (£)', 'Margin (£)', 'Notes'],
    [
      ['Black',                     0,    0,    0, 'Standard - no adjustment'],
      ['G603 Light Grey',        -100,    0,  100, 'Mason discount, customer pays standard'],
      ['Bahamas Blue (Visac Blue)', 0,  100,  100, 'Same cost as black, customer premium'],
      ['SA Impala',               50,  150,  100, 'Premium granite'],
    ]
  );

  writeTab(getOrCreate('Surrounds'),
    ['Type', 'Base Cost (£)', 'Base Sell (£)', 'Granite Cost Add (£)', 'Granite Sell Add (£)', 'Base Margin (£)', 'With Granite Margin (£)'],
    [
      ['Full Surround', 1400, 1600, 300, 400, 200, 300],
      ['Half Surround',  900, 1200, 300, 275, 300, 275],
      ['Tree Surround', 1050, 1400, 300, 275, 350, 325],
    ]
  );

  writeTab(getOrCreate('Stones'),
    ['Type', 'Standalone Cost (£)', 'With Surround Cost (£)', 'Sell Price (£)', 'Standalone Margin (£)', 'With Surround Margin (£)'],
    [
      ['Grey',                   60,   0,  100,  40, 100],
      ['White Quartz',          140,  40,  200,  60, 160],
      ['Black Pebbles',         195,  95,  300, 105, 205],
      ['White Pebbles',         195,  95,  300, 105, 205],
      ['Green Pebbles',         210, 110,  300,  90, 190],
      ['Blue Pebbles',          210, 110,  300,  90, 190],
      ['Blue Glass Chippings',  210, 110,  300,  90, 190],
      ['Green Glass Chippings', 210, 110,  300,  90, 190],
      ['Black Glass Chippings', 210, 110,  300,  90, 190],
    ]
  );

  writeTab(getOrCreate('Accessories'),
    ['Item Name', 'Size', 'Cost (£)', 'Sell (£)', 'Margin (£)', 'Margin (%)'],
    [
      ['Martin Vase',             'Standard',      160, 210, 50, 23.8],
      ['Chamfered Top Vase',      'Standard',      150, 210, 60, 28.6],
      ['Round Vase 4',            'Standard',      180, 210, 30, 14.3],
      ['12" x 12" Splayed Vase',  '12" x 12"',    160, 230, 70, 30.4],
      ['18" x 12" Splayed Vase',  '18" x 12"',    180, 250, 70, 28.0],
      ['6" x 6" x 12" Rose Vase', '6" x 6" x 12"',180,240, 60, 25.0],
      ['10" x 10" Heart Vase',    '10" x 10"',    200, 250, 50, 20.0],
      ['16" x 12" Book',          '16" x 12"',    180, 250, 70, 28.0],
      ['15" x 15" Heart Plaque',  '15" x 15"',    160, 210, 50, 23.8],
    ]
  );

  writeTab(getOrCreate('Cemetery_Fees'),
    ['Cemetery / Location', 'Fee (£)', 'Notes'],
    [
      ['None',        0,   'No cemetery fee'],
      ['Roselawn',  300,   ''],
      ['Blaris',    200,   ''],
      ['Church Yard',300,  'Varies - confirm with church before quoting'],
    ]
  );

  writeTab(getOrCreate('Additional_Services'),
    ['Service Name', 'Cost (£)', 'Sell (£)', 'Margin (£)', 'Margin (%)', 'Notes'],
    [['Reconcrete Full Grave', 120, 200, 80, 40.0, 'Full grave foundation']]
  );

  SpreadsheetApp.getUi().alert(
    '\u2705 Price book tabs created successfully!\n\n' +
    'Tabs: Headstones, Headstone_Colours, Surrounds, Stones, Accessories, Cemetery_Fees, Additional_Services\n\n' +
    'You can edit prices directly in each tab.\n' +
    'Redeploy as a new version, then refresh the tracker.'
  );
}

// ============================================================
// CUSTOMER PROOF PAGE — public data endpoint
// ============================================================
function getProofData(orderId, callback) {
  function send(obj) {
    const json = JSON.stringify(obj);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  try {
    if (!orderId) return send({ success: false, message: 'No order ID provided' });
    const sheet   = getOrCreateSheet();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    const get     = (row, col) => String(row[headers.indexOf(col)] || '');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(orderId)) continue;
      const row = data[i];

      // Parse proof status from log entries
      const logText = get(row, 'Log Entries');
      let proofStatus = 'pending';
      if (logText.includes('PROOF APPROVED'))       proofStatus = 'approved';
      else if (logText.includes('CHANGES REQUESTED')) proofStatus = 'changes_requested';
      else if (logText.includes('Proof link sent'))  proofStatus = 'sent';

      const totalSell    = parseFloat(get(row, 'Total Sell Price') || 0);
      const depositPaid  = parseFloat(get(row, 'Deposit Paid') || 0);
      const depositAmt   = totalSell > 0 ? Math.round((totalSell / 2) * 100) / 100 : 0;
      const stripeLinkRaw = get(row, 'Stripe Link ID') || '';
      const stripePayUrl  = stripeLinkRaw.startsWith('http') ? stripeLinkRaw : '';

      // Get estimate PDF URL from Files field
      let estimatePdfUrl = '';
      try {
        const files = JSON.parse(get(row, 'Files') || '[]');
        const estFile = files.find(f => f.type === 'Estimate PDF');
        if (estFile) estimatePdfUrl = estFile.viewUrl || estFile.downloadUrl || '';
      } catch(e) {}

      const proof = {
        orderId:           get(row, 'Order ID'),
        orderRef:          get(row, 'Order Ref') || String(orderId).slice(-8).toUpperCase(),
        customerFirstName: (get(row, 'Customer Name') || '').split(' ')[0],
        deceasedName:      get(row, 'Deceased Name'),
        hsType:            get(row, 'Headstone Type'),
        hsSize:            get(row, 'Headstone Size'),
        hsColour:          get(row, 'Headstone Colour'),
        inscriptionText:   get(row, 'Inscription Text'),
        inscriptionColour: get(row, 'Inscription Colour') || 'Gold',
        proofStatus:       proofStatus,
        totalSellPrice:    totalSell,
        depositPaid:       depositPaid,
        depositAmount:     depositAmt,
        stripePaymentUrl:  stripePayUrl,
        balanceDue:        Math.max(0, totalSell - depositPaid),
        estimatePdfUrl:    estimatePdfUrl
      };
      return send({ success: true, proof });
    }
    return send({ success: false, message: 'Order not found' });
  } catch (err) {
    return send({ success: false, message: err.toString() });
  }
}

// ============================================================
// CUSTOMER PROOF RESPONSE — approve or request changes
// ============================================================
function submitProofResponse(orderId, approved, message) {
  try {
    const sheet   = getOrCreateSheet();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(orderId)) continue;
      const rowNum = i + 1;

      // Append to log entries
      const logCol = headers.indexOf('Log Entries');
      if (logCol >= 0) {
        const existing = String(data[i][logCol] || '');
        const ts = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yyyy HH:mm');
        const entry = approved
          ? '[' + ts + '] Customer: \u2705 PROOF APPROVED \u2014 Customer confirmed all details correct'
          : '[' + ts + '] Customer: \u270f\ufe0f CHANGES REQUESTED \u2014 ' + (message || 'No details given');
        sheet.getRange(rowNum, logCol + 1).setValue(existing ? existing + ' | ' + entry : entry);
      }

      // Update Last Updated
      const luCol = headers.indexOf('Last Updated');
      if (luCol >= 0) sheet.getRange(rowNum, luCol + 1)
        .setValue(Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yyyy HH:mm'));

      // Auto-notify mason when the customer approves (skip on changes-requested)
      if (approved) {
        try {
          notifyMason(orderId, 'Customer approval', false);
        } catch (e) {
          Logger.log('Auto-notifyMason failed: ' + e);
        }
      }

      return respond(true, approved ? 'Proof approved' : 'Changes requested recorded');
    }
    return respond(false, 'Order not found');
  } catch (err) {
    return respond(false, err.toString());
  }
}

// ============================================================
// STORE ESTIMATE PDF IN GOOGLE DRIVE
// ============================================================
function storeEstimatePdf(orderId, ref, pdfBase64) {
  try {
    const base64Data = pdfBase64.replace(/^data:application\/pdf;base64,/, '');
    const pdfBytes   = Utilities.base64Decode(base64Data);
    const blob       = Utilities.newBlob(pdfBytes, 'application/pdf', 'Estimate_' + ref + '.pdf');

    // Save into the DC&S Drive folder
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const viewUrl     = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + file.getId();

    // Save the URL back to the order (Files JSON field)
    const sheet   = getOrCreateSheet();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== String(orderId)) continue;
      const rowNum  = i + 1;
      const filesCol = headers.indexOf('Files');
      if (filesCol >= 0) {
        let files = [];
        try { files = JSON.parse(data[i][filesCol] || '[]'); } catch(e) { files = []; }
        // Remove any old estimate entries
        files = files.filter(f => f.type !== 'Estimate PDF');
        files.push({
          type: 'Estimate PDF',
          name: 'Estimate_' + ref + '.pdf',
          viewUrl: viewUrl,
          downloadUrl: downloadUrl,
          uploadedAt: new Date().toISOString()
        });
        sheet.getRange(rowNum, filesCol + 1).setValue(JSON.stringify(files));
      }
      break;
    }

    return respond(true, 'Estimate stored in Drive', { viewUrl, downloadUrl });
  } catch (err) {
    return respond(false, 'Drive store error: ' + err.toString());
  }
}

// ============================================================
// NOTIFY STONE MASON — sends approved proof + full order details
// ============================================================
// Called automatically from submitProofResponse() when the customer
// approves a proof, and on demand from the tracker's "Notify Mason" button.
// Looks up the order, finds the latest "Proof" file in Drive, attaches it,
// and emails MASON_EMAIL with all the build-relevant fields.
function notifyMason(orderId, triggeredBy, force) {
  try {
    if (!MASON_EMAIL || MASON_EMAIL === 'mason@example.com') {
      return respond(false, 'MASON_EMAIL not configured. Set Script Property MASON_EMAIL in Apps Script project settings.');
    }
    const sheet   = getOrCreateSheet();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0];

    let rowNum = -1, rowVals = null;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(orderId)) {
        rowNum = i + 1;
        rowVals = data[i];
        break;
      }
    }
    if (rowNum < 0) return respond(false, 'Order not found: ' + orderId);

    // Build a sheetOrder map and parse it the same way getAllOrders does
    const sheetOrder = {};
    headers.forEach((h, idx) => sheetOrder[h] = rowVals[idx] !== undefined ? String(rowVals[idx]) : '');
    const order = mapSheetOrderToTracker(sheetOrder);

    // Skip if already notified — unless force=true (manual resend)
    if (!force && order.masonNotifiedAt) {
      return respond(true, 'Mason already notified at ' + order.masonNotifiedAt + ' (use force=true to resend)');
    }

    // Locate the most recent Proof file
    const proofFile = (order.files || []).filter(f => f.type === 'Proof')
      .sort((a, b) => new Date(b.uploadedAt || 0) - new Date(a.uploadedAt || 0))[0];

    const attachments = [];
    if (proofFile && proofFile.fileId) {
      try {
        attachments.push(DriveApp.getFileById(proofFile.fileId).getBlob());
      } catch (e) {
        Logger.log('Could not attach proof file ' + proofFile.fileId + ': ' + e);
      }
    }

    const ref = (order.orderId || '').slice(-6).toUpperCase();
    const subject = 'New Approved Job — ' + (order.deceasedName || 'Memorial') + ' — Ref ' + ref;

    const fmt = v => v == null || v === '' ? '<em style="color:#888;">—</em>' : String(v);
    const fmtList = arr => (arr && arr.length) ? arr.join(', ') : '<em style="color:#888;">none</em>';
    const noteEntries = (order.masonNoteEntries || [])
      .map(n => '<li>' + (n.text || '').replace(/</g, '&lt;') + ' <span style="color:#888;font-size:11px;">— ' + (n.author || '') + ', ' + (n.ts ? new Date(n.ts).toLocaleString('en-GB') : '') + '</span></li>')
      .join('');
    const masonNotesBlock = (order.masonNotes || noteEntries)
      ? '<h3 style="color:#15803d;margin:18px 0 6px;">Mason Notes</h3>'
        + (order.masonNotes ? '<p style="white-space:pre-wrap;background:#f0fdf4;border-left:3px solid #16a34a;padding:8px 12px;margin:0 0 8px;">' + String(order.masonNotes).replace(/</g, '&lt;') + '</p>' : '')
        + (noteEntries ? '<ul style="margin:0;padding-left:18px;">' + noteEntries + '</ul>' : '')
      : '';

    const proofLink = proofFile && proofFile.viewUrl
      ? '<p>Proof file: <a href="' + proofFile.viewUrl + '">' + (proofFile.name || 'View in Drive') + '</a> (also attached)</p>'
      : '<p style="color:#b45309;"><strong>No proof file attached — please request from office.</strong></p>';

    const htmlBody =
      '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;color:#1f2937;">'
      + '<div style="background:#15803d;padding:18px 22px;">'
      + '<h1 style="color:white;margin:0;font-size:20px;">David Crymble &amp; Sons — Approved Job</h1>'
      + '<p style="color:rgba(255,255,255,0.85);margin:4px 0 0;font-size:13px;">Ref ' + ref + ' — Proof approved</p>'
      + '</div>'
      + '<div style="padding:22px;">'
      + '<p>Hello ' + MASON_NAME + ',</p>'
      + '<p>The customer has approved the proof for the following memorial. Please proceed with manufacture.</p>'
      + proofLink
      + '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:14px;">'
      + tr('Customer', fmt(order.customerName))
      + tr('Deceased', fmt(order.deceasedName))
      + tr('DOB / DOD', fmt(order.deceasedDob) + ' — ' + fmt(order.deceasedDod))
      + tr('Cemetery', fmt(order.cemetery))
      + tr('Install date', fmt(order.installDate))
      + '</table>'
      + '<h3 style="margin:20px 0 6px;">Headstone</h3>'
      + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
      + tr('Type', fmt(order.hsType))
      + tr('Size', fmt(order.hsSize))
      + tr('Colour', fmt(order.hsColour))
      + tr('Finish', fmt(order.hsFinish))
      + '</table>'
      + '<h3 style="margin:20px 0 6px;">Surround / Stone</h3>'
      + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
      + tr('Surround', fmt(order.surroundType) + (order.surroundGranite ? ' (Granite upgrade)' : ''))
      + tr('Stone / chippings', fmt(order.stoneType))
      + tr('Accessories', fmtList(order.accessories))
      + '</table>'
      + '<h3 style="margin:20px 0 6px;">Inscription</h3>'
      + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
      + tr('Type', order.inscriptionType === 'additional' ? 'Additional on existing' : 'New inscription')
      + tr('Letter style', fmt(order.inscriptionStyle))
      + tr('Colour', fmt(order.inscriptionColour))
      + tr('Lines', fmt(order.inscriptionLines))
      + '</table>'
      + (order.inscriptionText
          ? '<pre style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:4px;padding:10px 12px;margin-top:8px;font-family:Georgia,serif;font-size:13px;white-space:pre-wrap;">'
            + String(order.inscriptionText).replace(/</g, '&lt;') + '</pre>'
          : '')
      + masonNotesBlock
      + '<hr style="border:none;border-top:1px solid #e5e7eb;margin:22px 0;">'
      + '<p style="font-size:12px;color:#6b7280;">Sent automatically by the DC&amp;S Memorial Tracker on customer proof approval'
      + (triggeredBy ? ' (manually re-sent by ' + triggeredBy + ')' : '')
      + '.</p>'
      + '</div></div>';

    MailApp.sendEmail({
      to: MASON_EMAIL,
      subject: subject,
      htmlBody: htmlBody,
      attachments: attachments,
      name: 'David Crymble & Sons'
    });

    // Stamp masonNotifiedAt + log entry on the order row
    const ts = new Date().toISOString();
    const notifiedAtCol = headers.indexOf('Mason Notified At');
    const notifiedByCol = headers.indexOf('Mason Notified By');
    if (notifiedAtCol >= 0) sheet.getRange(rowNum, notifiedAtCol + 1).setValue(ts);
    if (notifiedByCol >= 0) sheet.getRange(rowNum, notifiedByCol + 1).setValue(triggeredBy || 'auto');

    const logCol = headers.indexOf('Log Entries');
    if (logCol >= 0) {
      const existing = String(rowVals[logCol] || '');
      const stamp = Utilities.formatDate(new Date(), 'Europe/London', 'dd/MM/yyyy HH:mm');
      const author = triggeredBy || 'System';
      const entry = '[' + stamp + '] ' + author + ': 📧 Mason notified — ' + MASON_EMAIL
        + (proofFile ? ' (proof attached)' : ' (no proof file found)');
      sheet.getRange(rowNum, logCol + 1).setValue(existing ? existing + ' | ' + entry : entry);
    }

    return respond(true, 'Mason notified', { masonNotifiedAt: ts, attached: !!proofFile });
  } catch (err) {
    return respond(false, 'notifyMason error: ' + err.toString());
  }
}

// Tiny helper used by the mason-email HTML builder above
function tr(label, value) {
  return '<tr><td style="padding:4px 10px 4px 0;color:#6b7280;width:140px;vertical-align:top;">' + label
    + '</td><td style="padding:4px 0;font-weight:600;">' + value + '</td></tr>';
}

// ============================================================
// EMAIL ESTIMATE TO CUSTOMER
// ============================================================
function sendEstimateEmail(emailTo, customerName, ref, pdfBase64, proofUrl) {
  try {
    if (!emailTo) return respond(false, 'No email address provided');

    // Decode PDF from base64 data URI
    const base64Data = pdfBase64.replace(/^data:application\/pdf;base64,/, '');
    const pdfBytes   = Utilities.base64Decode(base64Data);
    const pdfBlob    = Utilities.newBlob(pdfBytes, 'application/pdf', 'Estimate_' + ref + '.pdf');

    const proofSection = proofUrl
      ? '<p>You can also <a href="' + proofUrl + '" style="color:#4a7c2f;font-weight:bold;">view and approve your proof online</a> — simply click the link on your phone or computer.</p>'
      : '';

    MailApp.sendEmail({
      to: emailTo,
      subject: 'Your Memorial Estimate \u2014 DC&S Memorials \u2014 Ref: ' + ref,
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">' +
        '<div style="background:#4a7c2f;padding:20px;text-align:center;">' +
        '<h1 style="color:white;margin:0;font-size:22px;">David Crymble &amp; Sons</h1>' +
        '<p style="color:rgba(255,255,255,0.8);margin:4px 0 0;font-size:13px;">Funeral Directors &amp; Memorial Masons</p>' +
        '</div>' +
        '<div style="padding:28px 24px;">' +
        '<p style="font-size:16px;">Dear ' + (customerName || 'Customer') + ',</p>' +
        '<p>Thank you for choosing David Crymble &amp; Sons. Please find attached your memorial estimate (Ref: <strong>' + ref + '</strong>).</p>' +
        '<p>The estimate includes a full price breakdown, a visual proof of the headstone with the inscription, and a 50% deposit payment link.</p>' +
        proofSection +
        '<p>Please review everything carefully. If you have any questions or would like any changes, please contact us.</p>' +
        '<div style="background:#f9f9f9;border-left:4px solid #4a7c2f;padding:14px 18px;margin:20px 0;border-radius:4px;">' +
        '<strong>T.</strong> 028 9066 7784<br>' +
        '<strong>E.</strong> info@Crymbleandsons.com<br>' +
        '<strong>W.</strong> Crymbleandsons.com' +
        '</div>' +
        '<p style="font-size:12px;color:#888;font-style:italic;">\u2018Our God will take care of everything you need.\u2019 Phil 4:19</p>' +
        '</div>' +
        '</div>',
      attachments: [pdfBlob],
      name: 'David Crymble & Sons'
    });

    return respond(true, 'Estimate emailed to ' + emailTo);
  } catch (err) {
    return respond(false, 'Email error: ' + err.toString());
  }
}
