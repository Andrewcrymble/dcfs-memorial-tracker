// ============================================================
// DC&S Headstone Order Tracker — Google Apps Script
// WITH COST/RETAIL PRICING SYSTEM
// ============================================================

const SHEET_NAME = "Orders";
const SHEET_ID = "1zLOSUsLgwglGkQkKa9x11EEIB04s2EOy1cBLBllkbRE";
const DRIVE_FOLDER_ID = "1nAxdUKug-s3pEQnX9RCps86crK--Vd4k";

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
  "Log Entries"
];

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
  }
  return sheet;
}

// ── PRICE BOOK LOADING ──
function loadPriceBook() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const priceBook = {};
  
  // Load Headstones
  try {
    const hsSheet = ss.getSheetByName("Headstones");
    if (hsSheet) {
      const hsData = hsSheet.getDataRange().getValues();
      priceBook.Headstones = hsData.slice(1).filter(row => row[0] && row[0] !== '').map(row => ({
        type: row[0],
        size: row[1],
        cost: parseFloat(row[2]) || 0,
        sell: parseFloat(row[3]) || 0,
        margin: parseFloat(row[4]) || 0,
        marginPct: parseFloat(row[5]) || 0
      }));
    }
  } catch(e) { Logger.log('Headstones tab error: ' + e); }
  
  // Load Headstone Colours
  try {
    const colSheet = ss.getSheetByName("Headstone_Colours");
    if (colSheet) {
      const colData = colSheet.getDataRange().getValues();
      priceBook.Colours = colData.slice(1).filter(row => row[0] && row[0] !== '').map(row => ({
        name: row[0],
        costAdj: parseFloat(row[1]) || 0,
        sellAdj: parseFloat(row[2]) || 0,
        margin: parseFloat(row[3]) || 0
      }));
    }
  } catch(e) { Logger.log('Colours tab error: ' + e); }
  
  // Load Surrounds
  try {
    const surSheet = ss.getSheetByName("Surrounds");
    if (surSheet) {
      const surData = surSheet.getDataRange().getValues();
      priceBook.Surrounds = surData.slice(1).filter(row => row[0] && row[0] !== '').map(row => ({
        type: row[0],
        baseCost: parseFloat(row[1]) || 0,
        baseSell: parseFloat(row[2]) || 0,
        graniteCostAdd: parseFloat(row[3]) || 0,
        graniteSellAdd: parseFloat(row[4]) || 0,
        baseMargin: parseFloat(row[5]) || 0,
        graniteMargin: parseFloat(row[6]) || 0
      }));
    }
  } catch(e) { Logger.log('Surrounds tab error: ' + e); }
  
  // Load Stones
  try {
    const stoneSheet = ss.getSheetByName("Stones");
    if (stoneSheet) {
      const stoneData = stoneSheet.getDataRange().getValues();
      priceBook.Stones = stoneData.slice(1).filter(row => row[0] && row[0] !== '').map(row => ({
        type: row[0],
        standaloneCost: parseFloat(row[1]) || 0,
        withSurroundCost: parseFloat(row[2]) || 0,
        sell: parseFloat(row[3]) || 0
      }));
    }
  } catch(e) { Logger.log('Stones tab error: ' + e); }
  
  // Load Accessories
  try {
    const accSheet = ss.getSheetByName("Accessories");
    if (accSheet) {
      const accData = accSheet.getDataRange().getValues();
      priceBook.Accessories = accData.slice(1).filter(row => row[0] && row[0] !== '').map(row => ({
        name: row[0],
        size: row[1],
        cost: parseFloat(row[2]) || 0,
        sell: parseFloat(row[3]) || 0
      }));
    }
  } catch(e) { Logger.log('Accessories tab error: ' + e); }
  
  // Load Cemetery Fees
  try {
    const cemSheet = ss.getSheetByName("Cemetery_Fees");
    if (cemSheet) {
      const cemData = cemSheet.getDataRange().getValues();
      priceBook.Cemetery_Fees = cemData.slice(1).filter(row => row[0] && row[0] !== '').map(row => ({
        location: row[0],
        fee: parseFloat(row[1]) || 0
      }));
    }
  } catch(e) { Logger.log('Cemetery tab error: ' + e); }
  
  // Load Additional Services
  try {
    const servSheet = ss.getSheetByName("Additional_Services");
    if (servSheet) {
      const servData = servSheet.getDataRange().getValues();
      priceBook.Services = servData.slice(1).filter(row => row[0] && row[0] !== '').map(row => ({
        name: row[0],
        cost: parseFloat(row[1]) || 0,
        sell: parseFloat(row[2]) || 0
      }));
    }
  } catch(e) { Logger.log('Services tab error: ' + e); }
  
  // Inscription pricing (hardcoded as per rules)
  priceBook.NewInscription = {
    freeLetter: 100,
    costPerLetterAfter100: 2.00,
    sellPerLetterAfter100: 3.00
  };
  
  priceBook.AdditionalInscription = {
    baseLetters: 50,
    baseCost: 150,
    baseSell: 250,
    costPerLetterAfter50: 3.00,
    sellPerLetterAfter50: 4.50
  };
  
  return priceBook;
}

function getOrCreateOrderFolder(orderId, customerName, deceasedName) {
  const rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const folderName = (customerName||"Unknown") + " — " + (deceasedName||"Memorial") + " — #" + String(orderId).slice(-6).toUpperCase();
  const existing = rootFolder.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return rootFolder.createFolder(folderName);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || "upsert";
    if (action === "upsert")      return upsertOrder(data.order);
    if (action === "delete")      return deleteOrder(data.orderId);
    if (action === "uploadFile")  return uploadFileToDrive(data);
    if (action === "deleteFile")  return deleteFileFromDrive(data.fileId);
    return respond(false, "Unknown action");
  } catch (err) {
    return respond(false, err.toString());
  }
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

// ── UPDATED doGet WITH COST/RETAIL MAPPING ──
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const priceBook = loadPriceBook(); // Load price book for frontend
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
          
          // Map to tracker field names (camelCase)
          const mappedOrder = mapSheetOrderToTracker(sheetOrder);
          orders.push(mappedOrder);
        });
    }
    
    const result = JSON.stringify({ 
      success: true, 
      orders,
      priceBook // Include price book in response
    });
    
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + result + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const result = JSON.stringify({ success: false, message: err.toString() });
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + result + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function mapSheetOrderToTracker(sheetOrder) {
  return {
    // IDs and timestamps
    orderId: sheetOrder["Order ID"],
    orderRef: sheetOrder["Order ID"],
    created: sheetOrder["Created"],
    lastUpdated: sheetOrder["Last Updated"],
    orderDate: sheetOrder["Order Date"],
    
    // Status
    status: capitalizeStatus(sheetOrder["Status"]),
    paymentStatus: sheetOrder["Payment Status"],
    
    // Customer details
    customerName: sheetOrder["Customer Name"],
    phone: sheetOrder["Phone"],
    email: sheetOrder["Email"],
    address: sheetOrder["Address"],
    
    // Deceased details
    deceasedName: sheetOrder["Deceased Name"],
    deceasedDob: sheetOrder["Date of Birth"],
    deceasedDod: sheetOrder["Date of Passing"],
    
    // Headstone — supports both old ("Sell Price") and new ("Headstone Sell Price") column names
    hsType: sheetOrder["Headstone Type"],
    hsSize: sheetOrder["Headstone Size"],
    hsColour: sheetOrder["Headstone Colour"] || "",
    hsColourAdj: parseFloat(sheetOrder["Headstone Colour Adj"]) || 0,
    hsFinish: sheetOrder["Headstone Finish"] || "",
    hsSellPrice: parseFloat(sheetOrder["Headstone Sell Price"] || sheetOrder["Sell Price"]) || 0,
    hsCostPrice: parseFloat(sheetOrder["Headstone Cost Price"]) || 0,

    // Surround
    surroundType: sheetOrder["Surround Type"],
    surroundGranite: sheetOrder["Granite Upgrade"] === "Yes",
    surroundSellPrice: parseFloat(sheetOrder["Surround Sell Price"]) || 0,
    surroundCostPrice: parseFloat(sheetOrder["Surround Cost Price"]) || 0,

    // Stones
    stoneType: sheetOrder["Stone / Chippings"],
    stoneSellPrice: parseFloat(sheetOrder["Stone Sell Price"]) || 0,
    stoneCostPrice: parseFloat(sheetOrder["Stone Cost Price"]) || 0,

    // Accessories
    accessories: sheetOrder["Accessories"] && sheetOrder["Accessories"] !== "No"
      ? sheetOrder["Accessories"].split(",").map(a => a.trim())
      : [],
    accessoriesSellPrice: parseFloat(sheetOrder["Accessories Sell Price"]) || 0,
    accessoriesCostPrice: parseFloat(sheetOrder["Accessories Cost Price"]) || 0,

    // Inscription — supports both old ("Inscription Charge") and new ("Inscription Sell Price")
    inscriptionType: sheetOrder["Inscription Type"] === "New Inscription" ? "new" : "additional",
    inscriptionText: sheetOrder["Inscription Text"],
    inscriptionLines: parseInt(sheetOrder["Inscription Lines"]) || 0,
    inscriptionStyle: sheetOrder["Letter Style"],
    inscriptionColour: sheetOrder["Inscription Colour"] || "",
    inscriptionPpl: parseFloat(sheetOrder["Price Per Line"]) || 35,
    inscriptionSellPrice: parseFloat(sheetOrder["Inscription Sell Price"] || sheetOrder["Inscription Charge"]) || 0,
    inscriptionCostPrice: parseFloat(sheetOrder["Inscription Cost Price"]) || 0,

    // Cemetery
    cemetery: sheetOrder["Cemetery / Location"],
    cemeteryFee: parseFloat(sheetOrder["Cemetery Fee"]) || 0,

    // Additional Services
    additionalServices: sheetOrder["Additional Services"] || "",
    servicesSellPrice: parseFloat(sheetOrder["Services Sell Price"]) || 0,
    servicesCostPrice: parseFloat(sheetOrder["Services Cost Price"]) || 0,

    // Totals — supports both old ("Total Price") and new ("Total Sell Price")
    totalSellPrice: parseFloat(sheetOrder["Total Sell Price"] || sheetOrder["Total Price"]) || 0,
    totalCostPrice: parseFloat(sheetOrder["Total Cost Price"]) || 0,
    profitMargin: parseFloat(sheetOrder["Profit Margin"]) || 0,
    marginPercentage: parseFloat(sheetOrder["Margin Percentage"]) || 0,
    
    // Payments
    depositPaid: parseFloat(sheetOrder["Deposit Paid"]) || 0,
    balanceDue: parseFloat(sheetOrder["Balance Due"]) || 0,
    
    // Dates
    proofDate: sheetOrder["Proof Date"],
    artworkApproved: sheetOrder["Proof Approved"] === "Yes",
    productionDate: sheetOrder["Production Start"],
    installDate: sheetOrder["Install Date"],
    
    // Notes
    artworkNotes: sheetOrder["Artwork Notes"],
    notes: sheetOrder["General Notes"],
    masonNotes: sheetOrder["Mason Notes"],
    
    // Files and charges
    folderLink: sheetOrder["Folder Link"],
    files: parseJSON(sheetOrder["Files"]),
    extraCharges: parseJSON(sheetOrder["Extra Charges"]),
    masonCharges: parseJSON(sheetOrder["Mason Charges"]),
    
    // Activity log
    log: parseLogEntries(sheetOrder["Log Entries"])
  };
}

function capitalizeStatus(status) {
  if (!status) return "Enquiry";
  const statusMap = {
    "enquiry": "Enquiry",
    "quoted": "Quoted",
    "confirmed": "Confirmed",
    "design": "In Design",
    "in design": "In Design",
    "production": "Production",
    "ready": "Ready",
    "installed": "Installed"
  };
  return statusMap[status.toLowerCase()] || "Enquiry";
}

function parseJSON(str) {
  if (!str || str === "") return [];
  try {
    return JSON.parse(str);
  } catch (e) {
    return [];
  }
}

function parseLogEntries(logString) {
  if (!logString || logString === "") return [];
  try {
    return logString.split(" | ").map((entry, index) => {
      const match = entry.match(/\[(.*?)\]\s*(.*?):\s*(.*)/);
      if (match) {
        return {
          id: Date.now() + index,
          ts: match[1],
          author: match[2],
          text: match[3]
        };
      }
      return null;
    }).filter(Boolean);
  } catch (e) {
    return [];
  }
}

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

  const balance = (parseFloat(order.totalSellPrice)||0) - (parseFloat(order.depositPaid)||0);
  const logText = (order.log||[]).map(l =>
    `[${new Date(l.ts).toLocaleDateString("en-GB")} ${new Date(l.ts).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}] ${l.author||"Staff"}: ${l.text}`
  ).join(" | ");

  const fmtDate = d => {
    if(!d || d === "Invalid Date") return "";
    try {
      const dt = new Date(d);
      if(isNaN(dt.getTime())) return "";
      return dt.toLocaleDateString("en-GB");
    } catch(e) { return ""; }
  };
  
  const fmtDateTime = () => new Date().toLocaleDateString("en-GB") + " " + new Date().toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"});

  const valueMap = {
    "Order ID": order.id || order.orderId || "",
    "Created": (()=>{
      const v = order.createdAt || order.created;
      if(!v || v === "Invalid Date") return fmtDateTime();
      try {
        const d = new Date(v);
        if(isNaN(d.getTime())) return fmtDateTime();
        return d.toLocaleDateString("en-GB") + " " + d.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"});
      } catch(e) { return fmtDateTime(); }
    })(),
    "Last Updated": fmtDateTime(),
    "Status": order.status || "enquiry",
    "Payment Status": order.paymentStatus || "Unpaid",
    "Customer Name": order.customerName || "",
    "Phone": order.phone || "",
    "Email": order.email || "",
    "Address": order.address || "",
    "Deceased Name": order.deceasedName || "",
    "Date of Birth": fmtDate(order.deceasedDob),
    "Date of Passing": fmtDate(order.deceasedDod),
    "Order Date": fmtDate(order.orderDate),
    "Headstone Type": order.hsType || "",
    "Headstone Size": order.hsSize || "",
    "Headstone Colour": order.hsColour || "",
    "Headstone Colour Adj": order.hsColourAdj || 0,
    "Headstone Finish": order.hsFinish || "",
    // New column names
    "Headstone Sell Price": order.hsSellPrice || 0,
    "Headstone Cost Price": order.hsCostPrice || 0,
    // Old column names (for existing sheets)
    "Sell Price": order.hsSellPrice || 0,
    "Surround Type": order.surroundType || "",
    "Granite Upgrade": order.surroundGranite ? "Yes" : "No",
    "Surround Sell Price": order.surroundSellPrice || 0,
    "Surround Cost Price": order.surroundCostPrice || 0,
    "Stone / Chippings": order.stoneType || "",
    "Stone Sell Price": order.stoneSellPrice || 0,
    "Stone Cost Price": order.stoneCostPrice || 0,
    "Accessories": (order.accessories||[]).join(", "),
    "Accessories Sell Price": order.accessoriesSellPrice || 0,
    "Accessories Cost Price": order.accessoriesCostPrice || 0,
    "Inscription Type": order.inscriptionType === "additional" ? "Additional on Existing" : "New Inscription",
    "Inscription Text": order.inscriptionText || "",
    "Inscription Lines": order.inscriptionLines || 0,
    "Letter Style": order.inscriptionStyle || "",
    "Inscription Colour": order.inscriptionColour || "Silver",
    // New column name
    "Inscription Sell Price": order.inscriptionSellPrice || 0,
    "Inscription Cost Price": order.inscriptionCostPrice || 0,
    // Old column name
    "Inscription Charge": order.inscriptionSellPrice || 0,
    "Price Per Line": order.inscriptionPpl || 35,
    "Cemetery / Location": order.cemetery || "",
    "Cemetery Fee": order.cemeteryFee || 0,
    "Additional Services": order.additionalServices || "",
    "Services Sell Price": order.servicesSellPrice || 0,
    "Services Cost Price": order.servicesCostPrice || 0,
    // New column name
    "Total Sell Price": order.totalSellPrice || 0,
    "Total Cost Price": order.totalCostPrice || 0,
    // Old column name
    "Total Price": order.totalSellPrice || 0,
    "Profit Margin": order.profitMargin || 0,
    "Margin Percentage": order.marginPercentage || 0,
    "Deposit Paid": order.depositPaid || 0,
    "Balance Due": Math.max(0, balance),
    "Proof Date": fmtDate(order.proofDate),
    "Proof Approved": order.artworkApproved ? "Yes" : "No",
    "Production Start": fmtDate(order.productionDate),
    "Install Date": fmtDate(order.installDate),
    "Artwork Notes": order.artworkNotes || "",
    "General Notes": order.notes || "",
    "Mason Notes": order.masonNotes || "",
    "Folder Link": order.folderLink || "",
    "Files": order.files ? JSON.stringify(order.files) : "[]",
    "Extra Charges": order.extraCharges ? JSON.stringify(order.extraCharges) : "[]",
    "Mason Charges": order.masonCharges ? JSON.stringify(order.masonCharges) : "[]",
    "Log Entries": logText,
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

function styleDataRow(sheet, rowNum, status) {
  const statusColors = {
    enquiry: "#f1f5f9", quoted: "#fef9e7", confirmed: "#dbeafe",
    design: "#ede9fe", production: "#fff7ed", ready: "#d1fae5", installed: "#dcfce7"
  };
  const bg = statusColors[status] || "#ffffff";
  const range = sheet.getRange(rowNum, 1, 1, HEADERS.length);
  range.setBackground(bg);
  range.setFontFamily("Arial");
  range.setFontSize(9);
  range.setVerticalAlignment("middle");
}

function respond(success, message, data) {
  const result = { success, message };
  if (data !== undefined) result.data = data;
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
