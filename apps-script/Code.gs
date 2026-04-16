// ============================================================
// DC&S Headstone Order Tracker — Google Apps Script
// WITH COST/RETAIL PRICING SYSTEM
// ============================================================

const SHEET_NAME = "Orders";
const SHEET_ID = "1_pbyhKL1IElgneBZHFIWG76Cv6hNKdUw0fgSHQjMN6M";
const DRIVE_FOLDER_ID = "1nAxdUKug-s3pEQnX9RCps86crK--Vd4k";
const PRICE_BOOK_ID = SHEET_ID; // Same sheet as orders — pricing tabs live here too

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

// ============================================================
// RUN THIS ONCE — creates a brand new master sheet with:
//   • Orders tab  (your 4 existing customers pre-loaded)
//   • Headstones, Surrounds, Stones, Accessories, Cemetery_Fees tabs
// After running, copy the new Sheet ID from the popup and update
// SHEET_ID at the top of this file, then redeploy.
// ============================================================
function createMasterSheet() {
  // --- Create new spreadsheet ---
  const ss = SpreadsheetApp.create('DC&S Memorial Tracker');
  const id = ss.getId();

  function styleHeader(sheet, cols) {
    const r = sheet.getRange(1, 1, 1, cols);
    r.setBackground('#1e3a5f');
    r.setFontColor('#ffffff');
    r.setFontWeight('bold');
    r.setFontSize(9);
    sheet.setFrozenRows(1);
  }

  // ── ORDERS TAB ──────────────────────────────────────────────
  const ordersSheet = ss.getActiveSheet();
  ordersSheet.setName('Orders');

  const orderHeaders = [
    'Order ID','Created','Last Updated','Status','Payment Status',
    'Customer Name','Phone','Email','Address',
    'Deceased Name','Date of Birth','Date of Passing','Order Date',
    'Headstone Type','Headstone Size','Headstone Colour','Headstone Finish',
    'Headstone Sell Price','Headstone Cost Price',
    'Surround Type','Granite Upgrade','Surround Sell Price','Surround Cost Price',
    'Stone / Chippings','Stone Sell Price','Stone Cost Price',
    'Accessories','Accessories Sell Price','Accessories Cost Price',
    'Inscription Type','Inscription Text','Inscription Lines','Letter Style','Inscription Colour',
    'Inscription Sell Price','Inscription Cost Price',
    'Cemetery / Location','Cemetery Fee',
    'Total Sell Price','Total Cost Price','Profit Margin','Margin Percentage',
    'Deposit Paid','Balance Due',
    'Proof Date','Proof Approved','Production Start','Install Date',
    'Artwork Notes','General Notes','Mason Notes',
    'Folder Link','Files','Extra Charges','Log Entries'
  ];

  const insc1 = 'AICKEN\nIn loving memory of\nGloria\nDied 18th February 2026\nA loving Wife, daughter, sister, aunt and friend\n\n(Gone from our lives, but forever in our hearts)';
  const insc2 = 'ENGLAND\nIn loving memory of\nOwen\nMuch loved husband of the late Ruth\nand devoted father to Wilson\nDied 23rd October 2025\n\n(Resting where no shadows fall)';
  const insc3 = 'And their son Richard, died 9th February 2026 much loved brother and Dad';
  const log1  = '[08/04/2026 10:49] ANDREW CRYMBLE: left voice message for Brian to call me back | [08/04/2026 14:40] ANDREW CRYMBLE: spoke with brian sending proof request over to Orchard | [08/04/2026 14:46] ANDREW CRYMBLE: Job sent to Orchard for proof';
  const log2  = '[08/04/2026 14:55] ANDREW CRYMBLE: job sent to Gerard | [08/04/2026 14:57] Andrew Crymble: estimate sent to Heather';
  const log3  = '[08/04/2026 14:47] ANDREW CRYMBLE: confirmed with Orchard and approved';

  // Columns match orderHeaders above (56 columns)
  // Sell Price stored in "Headstone Sell Price" col; Total in "Total Sell Price"
  const orderRows = [
    // Order 1 — Brian Aicken / Gloria
    ['mn6bfu5ld4cmk','08/04/2026 14:37','08/04/2026 14:46','confirmed','Unpaid',
     'Brian Aicken','7719593390','baicken07@gmail.cm','',
     'Gloria Aicken','','18/02/2026','25/03/2026',
     'Denmore','3ft (Base 3.6ft)','Black','Polished',
     3392,0,
     'Half Surround','No',0,0,
     '','','',
     '','','',
     'new',insc1,0,'Standard','',0,0,
     'Roselawn',300,
     3392,0,0,0,
     0,3392,
     '','No','','',
     '','Gone from our lives, but forever in our hearts on base','',
     '','','',log1],

    // Order 2 — Heather Kirker / Owen England
    ['mn7ej3cmh6eco','08/04/2026 14:37','08/04/2026 14:58','quoted','Unpaid',
     'Heather Kirker','','','',
     'Owen England','','23/10/2025','26/03/2026',
     'Ogee','2ft (Base 2.6ft)','Black','Polished',
     1246,0,
     '','No',0,0,
     '','','',
     '','','',
     'new',insc2,0,'Standard','',0,0,
     '',0,
     1246,0,0,0,
     0,1246,
     '','No','','',
     '','','',
     '','','',log2],

    // Order 3 — Drew Anderson / Richard Anderson
    ['mn7g3e0z0ocv1','08/04/2026 14:37','08/04/2026 14:47','production','Unpaid',
     'Drew Anderson','','','148 Erinvale Drive, BT10 0GF',
     'Richard Anderson','','','26/03/2026',
     '','','','',
     1600,0,
     'Half Surround','No',0,0,
     '','','',
     '','','',
     'additional',insc3,0,'Standard','',0,0,
     '',0,
     1600,0,0,0,
     0,1600,
     '','No','','',
     '','','',
     '','','',log3],

    // Order 4 — Lillian Kirkpatrick / Mum
    ['mn7gm54f82k8l','01/04/2026','01/04/2026 20:21','enquiry','Unpaid',
     'Lillian Kirkpatrick','','','',
     'Mum','','','26/03/2026',
     'Half Denmore','2.6ft (Base 3ft)','Black','Polished',
     1400,0,
     '','No',0,0,
     '','','',
     '','','',
     'new','',0,'Standard','',0,0,
     '',0,
     1400,0,0,0,
     0,1400,
     '','No','','',
     '','','',
     '','','',''],
  ];

  ordersSheet.getRange(1, 1, 1, orderHeaders.length).setValues([orderHeaders]);
  ordersSheet.getRange(2, 1, orderRows.length, orderHeaders.length).setValues(orderRows);
  styleHeader(ordersSheet, orderHeaders.length);
  ordersSheet.setRowHeights(2, orderRows.length, 80);
  ordersSheet.autoResizeColumns(1, orderHeaders.length);

  // ── PRICING TABS ─────────────────────────────────────────────
  function makeTab(name, headers, rows) {
    const sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    styleHeader(sh, headers.length);
    sh.autoResizeColumns(1, headers.length);
  }

  makeTab('Headstones',
    ['Type','Size','Cost (£)','Sell (£)','Margin (£)','Margin (%)'],
    [
      ['Ogee','1.9ft (Base 2ft)',750,900,150,16.7],
      ['Ogee','2ft (Base 2.6ft)',1000,1150,150,13.0],
      ['Ogee','2.6ft (Base 3ft)',1200,1400,200,14.3],
      ['Ogee','3ft (Base 3.6ft)',1400,1600,200,12.5],
      ['Ogee','3.6ft (Base 4ft)',1600,1800,200,11.1],
      ['G3','1.9ft (Base 2ft)',750,900,150,16.7],
      ['G3','2ft (Base 2.6ft)',1000,1150,150,13.0],
      ['G3','2.6ft (Base 3ft)',1200,1400,200,14.3],
      ['G3','3ft (Base 3.6ft)',1400,1600,200,12.5],
      ['G3','3.6ft (Base 4ft)',1600,1800,200,11.1],
      ['Denmore','3ft (Base 3.6ft)',1400,1600,200,12.5],
      ['Denmore','3.6ft (Base 4ft)',1600,1800,200,11.1],
      ['Half Denmore','2ft (Base 2.6ft)',1000,1150,150,13.0],
      ['Half Denmore','2.6ft (Base 3ft)',1200,1400,200,14.3],
      ['Half Denmore','3ft (Base 3.6ft)',1400,1600,200,12.5],
      ['Murphy','36"x30" / Base 42"x12"x5"',1400,2600,1200,46.2],
    ]
  );

  makeTab('Headstone_Colours',
    ['Colour Name','Cost Adjustment (£)','Sell Adjustment (£)','Margin (£)','Notes'],
    [
      ['Black',0,0,0,'Standard'],
      ['G603 Light Grey',-100,0,100,'Mason discount, customer pays standard'],
      ['Bahamas Blue (Visac Blue)',0,100,100,'Customer premium'],
      ['SA Impala',50,150,100,'Premium granite'],
    ]
  );

  makeTab('Surrounds',
    ['Type','Base Cost (£)','Base Sell (£)','Granite Cost Add (£)','Granite Sell Add (£)','Base Margin (£)','With Granite Margin (£)'],
    [
      ['Full Surround',1400,1600,300,400,200,300],
      ['Half Surround',900,1200,300,275,300,275],
      ['Tree Surround',1050,1400,300,275,350,325],
    ]
  );

  makeTab('Stones',
    ['Type','Standalone Cost (£)','With Surround Cost (£)','Sell Price (£)','Standalone Margin (£)','With Surround Margin (£)'],
    [
      ['Grey',60,0,100,40,100],
      ['White Quartz',140,40,200,60,160],
      ['Black Pebbles',195,95,300,105,205],
      ['White Pebbles',195,95,300,105,205],
      ['Green Pebbles',210,110,300,90,190],
      ['Blue Pebbles',210,110,300,90,190],
      ['Blue Glass Chippings',210,110,300,90,190],
      ['Green Glass Chippings',210,110,300,90,190],
      ['Black Glass Chippings',210,110,300,90,190],
    ]
  );

  makeTab('Accessories',
    ['Item Name','Size','Cost (£)','Sell (£)','Margin (£)','Margin (%)'],
    [
      ['Martin Vase','Standard',160,210,50,23.8],
      ['Chamfered Top Vase','Standard',150,210,60,28.6],
      ['Round Vase 4','Standard',180,210,30,14.3],
      ['12" x 12" Splayed Vase','12" x 12"',160,230,70,30.4],
      ['18" x 12" Splayed Vase','18" x 12"',180,250,70,28.0],
      ['6" x 6" x 12" Rose Vase','6" x 6" x 12"',180,240,60,25.0],
      ['10" x 10" Heart Vase','10" x 10"',200,250,50,20.0],
      ['16" x 12" Book','16" x 12"',180,250,70,28.0],
      ['15" x 15" Heart Plaque','15" x 15"',160,210,50,23.8],
    ]
  );

  makeTab('Cemetery_Fees',
    ['Cemetery / Location','Fee (£)','Notes'],
    [
      ['None',0,'No cemetery fee'],
      ['Roselawn',300,''],
      ['Blaris',200,''],
      ['Church Yard',300,'Varies - confirm with church'],
    ]
  );

  makeTab('Additional_Services',
    ['Service Name','Cost (£)','Sell (£)','Margin (£)','Margin (%)','Notes'],
    [['Reconcrete Full Grave',120,200,80,40.0,'Full grave foundation']]
  );

  // Done — show the new sheet URL & ID
  const url = ss.getUrl();
  SpreadsheetApp.getUi().alert(
    '✅ Master sheet created!\n\n' +
    'URL: ' + url + '\n\n' +
    'NEW SHEET ID:\n' + id + '\n\n' +
    'Copy that Sheet ID, then update SHEET_ID at the top of Code.gs and redeploy.'
  );
}

// ============================================================
// RUN THIS ONCE to create all price book tabs in your Sheet
// Open Apps Script editor → select setupPriceBook → click Run
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
    // Style header row
    const hdr = sheet.getRange(1, 1, 1, headers.length);
    hdr.setBackground('#1e3a5f');
    hdr.setFontColor('#ffffff');
    hdr.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  // ── Headstones ──
  writeTab(
    getOrCreate('Headstones'),
    ['Type', 'Size', 'Cost (£)', 'Sell (£)', 'Margin (£)', 'Margin (%)'],
    [
      ['Ogee', '1.9ft (Base 2ft)',          750,  900,  150, 16.7],
      ['Ogee', '2ft (Base 2.6ft)',          1000, 1150, 150, 13.0],
      ['Ogee', '2.6ft (Base 3ft)',          1200, 1400, 200, 14.3],
      ['Ogee', '3ft (Base 3.6ft)',          1400, 1600, 200, 12.5],
      ['Ogee', '3.6ft (Base 4ft)',          1600, 1800, 200, 11.1],
      ['G3',   '1.9ft (Base 2ft)',          750,  900,  150, 16.7],
      ['G3',   '2ft (Base 2.6ft)',          1000, 1150, 150, 13.0],
      ['G3',   '2.6ft (Base 3ft)',          1200, 1400, 200, 14.3],
      ['G3',   '3ft (Base 3.6ft)',          1400, 1600, 200, 12.5],
      ['G3',   '3.6ft (Base 4ft)',          1600, 1800, 200, 11.1],
      ['Denmore',      '3ft (Base 3.6ft)',  1400, 1600, 200, 12.5],
      ['Denmore',      '3.6ft (Base 4ft)',  1600, 1800, 200, 11.1],
      ['Half Denmore', '2ft (Base 2.6ft)',  1000, 1150, 150, 13.0],
      ['Half Denmore', '2.6ft (Base 3ft)',  1200, 1400, 200, 14.3],
      ['Half Denmore', '3ft (Base 3.6ft)',  1400, 1600, 200, 12.5],
      ['Murphy', '36"x30" / Base 42"x12"x5"', 1400, 2600, 1200, 46.2],
    ]
  );

  // ── Headstone_Colours ──
  writeTab(
    getOrCreate('Headstone_Colours'),
    ['Colour Name', 'Cost Adjustment (£)', 'Sell Adjustment (£)', 'Margin (£)', 'Notes'],
    [
      ['Black',                  0,    0,    0,    'Standard - no adjustment'],
      ['G603 Light Grey',       -100,  0,   100,   'Mason discount, customer pays standard'],
      ['Bahamas Blue (Visac Blue)', 0, 100,  100,  'Same cost as black, customer premium'],
      ['SA Impala',              50,  150,   100,  'Premium granite'],
    ]
  );

  // ── Surrounds ──
  writeTab(
    getOrCreate('Surrounds'),
    ['Type', 'Base Cost (£)', 'Base Sell (£)', 'Granite Cost Add (£)', 'Granite Sell Add (£)', 'Base Margin (£)', 'With Granite Margin (£)'],
    [
      ['Full Surround', 1400, 1600, 300, 400, 200, 300],
      ['Half Surround',  900, 1200, 300, 275, 300, 275],
      ['Tree Surround', 1050, 1400, 300, 275, 350, 325],
    ]
  );

  // ── Stones ──
  writeTab(
    getOrCreate('Stones'),
    ['Type', 'Standalone Cost (£)', 'With Surround Cost (£)', 'Sell Price (£)', 'Standalone Margin (£)', 'With Surround Margin (£)'],
    [
      ['Grey',                  60,   0,  100, 40,  100],
      ['White Quartz',         140,  40,  200, 60,  160],
      ['Black Pebbles',        195,  95,  300, 105, 205],
      ['White Pebbles',        195,  95,  300, 105, 205],
      ['Green Pebbles',        210, 110,  300, 90,  190],
      ['Blue Pebbles',         210, 110,  300, 90,  190],
      ['Blue Glass Chippings', 210, 110,  300, 90,  190],
      ['Green Glass Chippings',210, 110,  300, 90,  190],
      ['Black Glass Chippings',210, 110,  300, 90,  190],
    ]
  );

  // ── Accessories ──
  writeTab(
    getOrCreate('Accessories'),
    ['Item Name', 'Size', 'Cost (£)', 'Sell (£)', 'Margin (£)', 'Margin (%)'],
    [
      ['Martin Vase',            'Standard',     160, 210, 50, 23.8],
      ['Chamfered Top Vase',     'Standard',     150, 210, 60, 28.6],
      ['Round Vase 4',           'Standard',     180, 210, 30, 14.3],
      ['12" x 12" Splayed Vase', '12" x 12"',   160, 230, 70, 30.4],
      ['18" x 12" Splayed Vase', '18" x 12"',   180, 250, 70, 28.0],
      ['6" x 6" x 12" Rose Vase','6" x 6" x 12"',180,240, 60, 25.0],
      ['10" x 10" Heart Vase',   '10" x 10"',   200, 250, 50, 20.0],
      ['16" x 12" Book',         '16" x 12"',   180, 250, 70, 28.0],
      ['15" x 15" Heart Plaque', '15" x 15"',   160, 210, 50, 23.8],
    ]
  );

  // ── Cemetery_Fees ──
  writeTab(
    getOrCreate('Cemetery_Fees'),
    ['Cemetery / Location', 'Fee (£)', 'Notes'],
    [
      ['None',        0,   'No cemetery fee'],
      ['Roselawn',    300, ''],
      ['Blaris',      200, ''],
      ['Church Yard', 300, 'Varies - confirm with church before quoting'],
    ]
  );

  // ── Additional_Services ──
  writeTab(
    getOrCreate('Additional_Services'),
    ['Service Name', 'Cost (£)', 'Sell (£)', 'Margin (£)', 'Margin (%)', 'Notes'],
    [
      ['Reconcrete Full Grave', 120, 200, 80, 40.0, 'Full grave foundation'],
    ]
  );

  SpreadsheetApp.getUi().alert(
    '✅ Price book tabs created successfully!\n\n' +
    'Tabs created:\n' +
    '• Headstones (16 products)\n' +
    '• Headstone_Colours (4 colours)\n' +
    '• Surrounds (3 types)\n' +
    '• Stones (9 types)\n' +
    '• Accessories (9 items)\n' +
    '• Cemetery_Fees (4 locations)\n' +
    '• Additional_Services (1 service)\n\n' +
    'You can now edit any prices directly in the tabs.\n' +
    'Redeploy the Apps Script as a new version, then refresh the tracker.'
  );
}

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
  const ss = SpreadsheetApp.openById(PRICE_BOOK_ID);
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
