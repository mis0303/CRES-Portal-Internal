var FOLDER_NAME = "CRES Estimates";
var INDEX_SHEET = "Index";
var LIBRARY_SHEET = "Library";

// Default library (used first time setupLibrary is run)
var DEFAULT_LIBRARY = [
  { name: "Venues", subs: [
    { title: "Room Category", items: ["Single Occupancy", "Double Occupancy", "Triple Occupancy"] },
    { title: "Banquets/Ballrooms", items: [] }
  ]},
  { name: "Manpower", subs: [
    { title: "", items: ["Housekeeping", "Hostess", "Volunteers", "supervisors"] }
  ]},
  { name: "Welcome (Hotel Registration)", subs: [
    { title: "Branding", items: ["Registration Backdrop", "Registration Desk Facade", "Info Desk in Lobby", "Carpet for pathway"] }
  ]},
  { name: "Artist & Entertainment", subs: [
    { title: "Performances", items: ["Live Band / Base Band", "DJ (Bollywood / English)", "Magicians", "Dancers"] }
  ]},
  { name: "Fabrication", subs: [
    { title: "Main Entrance Gate Arch", items: [
      "Top with frame and flex", "Pillars with frame and flex", "Ballroom Entrance Arch"
    ]},
    { title: "Stage", items: [
      "Platform for Stage", "Skirting for Stage", "LED Riser", "Skirting for LED Riser",
      "Steps", "Side Panel", "Side View Cutter", "Podium",
      "Console with Black Masking", "Camera Riser", "LED Side Riser", "Riser for Side Panel"
    ]},
    { title: "Branding", items: ["Standees", "Customized Photo Booth", "Platform for photo booth"] }
  ]},
  { name: "Sound", subs: [
    { title: "Equipment", items: [
      "Tops (JBL VRX)", "Dual Base for Stage", "Stage Monitors", "Cordless Mics",
      "Headset Mics", "Podium Mics", "Mixer / Pioneer Player", "Cabling and Technicians",
      "Tech rider for band"
    ]}
  ]},
  { name: "Lights", subs: [
    { title: "Fixtures", items: [
      "LED Par Cans RGB", "LED Par Cans Face Light", "Sharpy", "Moving LED Wash",
      "Profile", "Smoke", "Light Technician & Operator"
    ]},
    { title: "Structures", items: ["Front Goal Post", "T truss", "Box truss", "Light stand"] }
  ]},
  { name: "LED Wall (2.9)", subs: [
    { title: "Screens", items: [
      "Main LED Wall", "Side Screen", "Comfort Plasma", "Timer Plasma"
    ]},
    { title: "Control", items: [
      "Switcher", "Clicker", "Laptops", "Resolume Server", "Operator", "Watchout"
    ]}
  ]},
  { name: "sfx (special effects)", subs: [
    { title: "", items: ["pyros"] }
  ]},
  { name: "Security and safety", subs: [
    { title: "", items: ["bouncer", "lady bouncer", "security guard", "gunmen"] }
  ]},
  { name: "Permission and liscences", subs: [
    { title: "", items: [
      "IPRS / PWD / Police intimation (Band)",
      "PWD / IPRS / NOVEX / PPL (DJ)",
      "Graphic Designer"
    ] }
  ]},
  { name: "Food and Beverage", subs: [
    { title: "", items: [] }
  ]},
  { name: "Power Supply", subs: [
    { title: "", items: ["generator", "power socket", "wire managers", "cabling"] }
  ]},
  { name: "Directional Signages", subs: [
    { title: "", items: [] }
  ]},
  { name: "collaterals", subs: [
    { title: "", items: [
      "que cards", "tent cards(name card)", "mic branding", "table numbers",
      "badges/lenyard", "winnner envelope", "wrist bands/paper bands"
    ] }
  ]},
  { name: "Rental", subs: [
    { title: "", items: [
      "Queue manager", "Sofa", "Coffee table", "Easel stand", "Catering table",
      "Round table", "Plastic chairs", "Sofa chairs", "Cocktail / high table",
      "mist fans", "air coolers", "portable ac", "tower ac/Standing ac",
      "electric heater", "pedesteal fans", "jumbo fan", "coffee machine",
      "printer", "water dispensor"
    ] }
  ]},
  { name: "Hampers And GiveAway", subs: [
    { title: "", items: [] }
  ]},
  { name: "Event Coverage", subs: [
    { title: "Photo & Video", items: [
      "Photographer", "Videographer (Wide Angle + Gimble)", "Montage / Highlight Video"
    ]}
  ]},
  { name: "Activities", subs: [
    { title: "Engagement Booths", items: [
      "360° Photo Booth", "Instant Photo Strip", "Word Cloud Selfie Booth"
    ]}
  ]},
  { name: "Communication", subs: [
    { title: "Equipment", items: ["Clearcom / Walkies"] }
  ]},
  { name: "Misc & Transport", subs: [
    { title: "Setup", items: [
      "Setup Transport (labour, mathadi, equipment)",
      "Crew Food / Promoters / Hostess",
      "Miscellaneous Expenses",
      "TBL (team stay + food)"
    ]}
  ]}
];

// =================== SETUP ===================
function setup() {
  var folder = getOrCreateFolder_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var idx = ss.getSheetByName(INDEX_SHEET);
  if (!idx) {
    idx = ss.insertSheet(INDEX_SHEET);
    idx.getRange(1, 1, 1, 8).setValues([[
      "EventID", "EventName", "Client", "Venue", "FileID", "URL", "LatestVersion", "UpdatedAt"
    ]]).setFontWeight("bold").setBackground("#10243e").setFontColor("white");
    idx.setFrozenRows(1);
    idx.setColumnWidth(1, 180);
    idx.setColumnWidth(2, 200);
    idx.setColumnWidth(3, 180);
    idx.setColumnWidth(5, 280);
    idx.setColumnWidth(6, 280);
  }
  setupLibrary_();
  Browser.msgBox("Setup complete. Folder ready: " + folder.getName());
}

// =================== LIBRARY ===================
function setupLibrary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (lib) return; // don't overwrite existing library

  lib = ss.insertSheet(LIBRARY_SHEET);
  lib.getRange(1, 1, 1, 3).setValues([["Category", "SubSection", "Item"]])
    .setFontWeight("bold").setBackground("#10243e").setFontColor("white");
  lib.setFrozenRows(1);
  lib.setColumnWidth(1, 240);
  lib.setColumnWidth(2, 200);
  lib.setColumnWidth(3, 260);

  var rows = [];
  DEFAULT_LIBRARY.forEach(function(cat) {
    cat.subs.forEach(function(sub) {
      if (sub.items.length === 0) {
        // Empty sub-section — still emit a row for structure
        rows.push([cat.name, sub.title, ""]);
      } else {
        sub.items.forEach(function(item) {
          rows.push([cat.name, sub.title, item]);
        });
      }
    });
  });
  if (rows.length > 0) {
    lib.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

// Reset Library tab to fresh DEFAULT_LIBRARY structure — run this manually
// to overwrite existing Library data with the latest template.
function resetLibrary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib) {
    setupLibrary_();
    Browser.msgBox("Library tab created with default template.");
    return;
  }
  // Clear existing rows (except header)
  var lastRow = lib.getLastRow();
  if (lastRow > 1) {
    lib.getRange(2, 1, lastRow - 1, 3).clearContent();
  }
  // Write fresh from DEFAULT_LIBRARY
  var rows = [];
  DEFAULT_LIBRARY.forEach(function(cat) {
    cat.subs.forEach(function(sub) {
      if (sub.items.length === 0) {
        rows.push([cat.name, sub.title, ""]);
      } else {
        sub.items.forEach(function(item) {
          rows.push([cat.name, sub.title, item]);
        });
      }
    });
  });
  if (rows.length > 0) {
    lib.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  Browser.msgBox("Library reset. " + rows.length + " rows written.");
}

function getLibrary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib || lib.getLastRow() < 2) return { categories: [] };
  var data = lib.getRange(2, 1, lib.getLastRow() - 1, 3).getValues();

  var cats = {};
  var order = [];
  data.forEach(function(r) {
    var cat = String(r[0] || "").trim();
    var sub = String(r[1] || "").trim(); // blank means "direct items"
    var item = String(r[2] || "").trim();
    if (!cat) return;
    if (!cats[cat]) { cats[cat] = { name: cat, subOrder: [], subs: {} }; order.push(cat); }
    if (!cats[cat].subs[sub]) { cats[cat].subs[sub] = { title: sub, items: [] }; cats[cat].subOrder.push(sub); }
    if (item) cats[cat].subs[sub].items.push(item);
  });
  var result = order.map(function(c) {
    return {
      name: c,
      subs: cats[c].subOrder.map(function(s) {
        return { title: s, items: cats[c].subs[s].items.map(function(i){ return { name: i }; }) };
      })
    };
  });
  return { categories: result };
}

function renameLibrary_(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib) return { status: "error", message: "Library not set up" };
  var data = lib.getDataRange().getValues();
  var level = payload.level;
  var oldName = payload.oldName;
  var newName = payload.newName;
  var scope = payload.scope || {};

  if (!oldName || !newName || oldName === newName) return { status: "error", message: "Invalid rename request" };

  var updated = 0;
  for (var i = 1; i < data.length; i++) {
    var cat = data[i][0], sub = data[i][1], item = data[i][2];
    if (level === "category" && cat === oldName) {
      lib.getRange(i + 1, 1).setValue(newName); updated++;
    } else if (level === "subsection" && cat === (scope.category || cat) && sub === oldName) {
      if (!scope.category || cat === scope.category) {
        lib.getRange(i + 1, 2).setValue(newName); updated++;
      }
    } else if (level === "item" && item === oldName) {
      if ((!scope.category || cat === scope.category) &&
          (!scope.subSection || sub === scope.subSection)) {
        lib.getRange(i + 1, 3).setValue(newName); updated++;
      }
    }
  }
  return { status: "success", updated: updated };
}

function deleteLibrarySubsection_(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib) return { status: "error", message: "Library not set up" };
  var data = lib.getDataRange().getValues();
  var category = payload.category;
  var subSection = payload.subSection;
  var keepItems = !!payload.keepItems;

  if (!category || !subSection) return { status: "error", message: "Missing category/subsection" };

  var deleted = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === category && data[i][1] === subSection) {
      if (keepItems) {
        lib.getRange(i + 1, 2).setValue(""); // clear sub-section but keep row
      } else {
        lib.deleteRow(i + 1);
        deleted++;
      }
    }
  }
  return { status: "success", deleted: deleted };
}

function reorderLibrary_(payload) {
  // payload.library = full new library structure (frontend sends after drag reorder)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib) return { status: "error", message: "Library not set up" };

  var newLib = payload.library || [];
  var rows = [];
  newLib.forEach(function(cat) {
    (cat.subs || []).forEach(function(sub) {
      (sub.items || []).forEach(function(item) {
        var itemName = typeof item === "string" ? item : (item.name || "");
        rows.push([cat.name, sub.title, itemName]);
      });
    });
  });

  // Clear existing (except header) and write fresh
  var lastRow = lib.getLastRow();
  if (lastRow > 1) lib.getRange(2, 1, lastRow - 1, 3).clearContent();
  if (rows.length > 0) lib.getRange(2, 1, rows.length, 3).setValues(rows);

  return { status: "success", rowsWritten: rows.length };
}

function getOrCreateFolder_() {
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// =================== GET ===================
function doGet(e) {
  try {
    var action = e.parameter.action;
    if (action === "listEvents") return jsonOut_(listEvents_());
    if (action === "listVersions") return jsonOut_(listVersions_(e.parameter.eventId));
    if (action === "getVersion") return jsonOut_(getVersion_(e.parameter.eventId, e.parameter.version));
    if (action === "getLibrary") return jsonOut_(getLibrary_());
    return jsonOut_({ error: "Unknown action" });
  } catch (err) {
    return jsonOut_({ error: err.message, stack: err.stack });
  }
}

function listEvents_() {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  if (!idx || idx.getLastRow() < 2) return { events: [] };
  var data = idx.getRange(2, 1, idx.getLastRow() - 1, 8).getValues();
  var events = data.map(function(r) {
    return { eventId: r[0], eventName: r[1], client: r[2], venue: r[3],
             fileId: r[4], url: r[5], latestVersion: r[6], updatedAt: r[7] };
  }).filter(function(ev) { return ev.eventId; });
  return { events: events };
}

function listVersions_(eventId) {
  var info = findEventFile_(eventId);
  if (!info) return { versions: [] };
  var ss = SpreadsheetApp.openById(info.fileId);
  var versions = ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return /^V\d+$/.test(n); })
    .sort(function(a, b) { return parseInt(a.slice(1)) - parseInt(b.slice(1)); });
  return { versions: versions };
}

function getVersion_(eventId, version) {
  var info = findEventFile_(eventId);
  if (!info) return { error: "Event not found" };
  var ss = SpreadsheetApp.openById(info.fileId);
  var sheet = ss.getSheetByName(version);
  if (!sheet) return { error: "Version not found" };

  // Read the hidden JSON snapshot we wrote into the sheet's developer metadata
  // (simpler: read from a hidden tab _Data_<version>)
  var dataSheet = ss.getSheetByName("_Data_" + version);
  if (dataSheet) {
    var raw = dataSheet.getRange(1, 1).getValue();
    try {
      return JSON.parse(raw);
    } catch (e) {}
  }
  return { header: {}, days: [], lineItems: [], logistics: [], summary: {} };
}

// =================== POST ===================
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === "create") return jsonOut_(createEvent_(payload));
    if (payload.action === "updateCurrent") return jsonOut_(updateCurrent_(payload));
    if (payload.action === "saveAsNew") return jsonOut_(saveAsNew_(payload));
    if (payload.action === "deleteVersion") return jsonOut_(deleteVersion_(payload));
    if (payload.action === "deleteEvent") return jsonOut_(deleteEvent_(payload));
    if (payload.action === "renameLibrary") return jsonOut_(renameLibrary_(payload));
    if (payload.action === "deleteLibrarySubsection") return jsonOut_(deleteLibrarySubsection_(payload));
    if (payload.action === "reorderLibrary") return jsonOut_(reorderLibrary_(payload));
    return jsonOut_({ status: "error", message: "Unknown action" });
  } catch (err) {
    return jsonOut_({ status: "error", message: err.message });
  }
}

function deleteVersion_(payload) {
  var info = findEventFile_(payload.eventId);
  if (!info) return { status: "error", message: "Event not found" };
  var ss = SpreadsheetApp.openById(info.fileId);

  // Safety: don't allow deleting the only version
  var versionSheets = ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return /^V\d+$/.test(n); });
  if (versionSheets.length <= 1) {
    return { status: "error", message: "Cannot delete the only version. Delete the entire event instead." };
  }

  var version = payload.version;
  var sheet = ss.getSheetByName(version);
  if (!sheet) return { status: "error", message: "Version not found" };

  ss.deleteSheet(sheet);
  var dataSheet = ss.getSheetByName("_Data_" + version);
  if (dataSheet) ss.deleteSheet(dataSheet);

  // Recompute latest V and update Index
  var remaining = ss.getSheets()
    .map(function(s) { return s.getName(); })
    .filter(function(n) { return /^V\d+$/.test(n); })
    .sort(function(a, b) { return parseInt(a.slice(1)) - parseInt(b.slice(1)); });
  var newLatest = remaining[remaining.length - 1];

  updateIndexLatest_(payload.eventId, newLatest);

  return { status: "success", eventId: payload.eventId, newLatest: newLatest };
}

function deleteEvent_(payload) {
  var info = findEventFile_(payload.eventId);
  if (!info) return { status: "error", message: "Event not found" };

  // Move file to Trash (30-day recovery)
  var file = DriveApp.getFileById(info.fileId);
  file.setTrashed(true);

  // Remove from Index
  removeFromIndex_(payload.eventId);

  return { status: "success", eventId: payload.eventId };
}

function updateIndexLatest_(eventId, newLatest) {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  var data = idx.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === eventId) {
      idx.getRange(i + 1, 7).setValue(newLatest);
      idx.getRange(i + 1, 8).setValue(new Date());
      return;
    }
  }
}

function removeFromIndex_(eventId) {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  var data = idx.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === eventId) {
      idx.deleteRow(i + 1);
      return;
    }
  }
}

function createEvent_(payload) {
  var folder = getOrCreateFolder_();
  var fileName = "Estimate - " + (payload.header.eventName || "Untitled") + " - " + (payload.header.client || "");
  var newSS = SpreadsheetApp.create(fileName);
  var file = DriveApp.getFileById(newSS.getId());
  // Move file to our folder
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // Remove the default blank "Sheet1"
  var defaultSheet = newSS.getSheets()[0];

  // Create V1
  var version = "V1";
  writeVersion_(newSS, version, payload);
  newSS.deleteSheet(defaultSheet);

  var eventId = "EVT-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  appendToIndex_(eventId, payload.header, newSS.getId(), newSS.getUrl(), version);

  return { status: "success", eventId: eventId, version: version, sheetUrl: newSS.getUrl() };
}

function updateCurrent_(payload) {
  var info = findEventFile_(payload.eventId);
  if (!info) return { status: "error", message: "Event not found" };
  var ss = SpreadsheetApp.openById(info.fileId);
  var version = payload.fromVersion;
  if (!version) return { status: "error", message: "Version missing" };

  var sheet = ss.getSheetByName(version);
  if (sheet) ss.deleteSheet(sheet);
  var dataSheet = ss.getSheetByName("_Data_" + version);
  if (dataSheet) ss.deleteSheet(dataSheet);

  writeVersion_(ss, version, payload);
  updateIndex_(payload.eventId, payload.header, version);
  return { status: "success", eventId: payload.eventId, version: version, sheetUrl: ss.getUrl() };
}

function saveAsNew_(payload) {
  var info = findEventFile_(payload.eventId);
  if (!info) return { status: "error", message: "Event not found" };
  var ss = SpreadsheetApp.openById(info.fileId);

  var maxV = 0;
  ss.getSheets().forEach(function(s) {
    var m = s.getName().match(/^V(\d+)$/);
    if (m) maxV = Math.max(maxV, parseInt(m[1]));
  });
  var newVersion = "V" + (maxV + 1);

  writeVersion_(ss, newVersion, payload);
  updateIndex_(payload.eventId, payload.header, newVersion);
  return { status: "success", eventId: payload.eventId, version: newVersion, sheetUrl: ss.getUrl() };
}

// =================== WRITE VERSION SHEET ===================
function writeVersion_(ss, version, payload) {
  var sheet = ss.insertSheet(version);
  var row = 1;
  var NCOL = 11; // Amount(9) + Remark Client(10) + Remark Doer(11)

  var NOTE_LINES = [
    "GST @ 18% will be charge on the final bill",
    "Payment: 75% advance on confirmation and 25% immediately after the event against the final invoice.",
    "Payment Details: Cheques/DDs in favour of Creative Rays Entertainment Solutions, payable at Mumbai.",
    "Bank: Bank of Baroda | A/c No.: 17580200001303 | IFSC: BARB0GOREAS (5th character is Zero)",
    "Delayed Payment: Delayed payments shall attract interest as applicable under the MSMED Act, 2006.",
    "Additional Scope: Any additional elements, services, manpower, equipment or changes requested beyond the approved scope will be charged extra, subject to feasibility.",
    "Rental: Equipment/materials supplied are on a rental basis unless specifically mentioned otherwise.",
    "TDS & Taxes: TDS, wherever legally applicable, may be deducted at the prescribed rate. GST and applicable taxes shall be charged extra.",
    "Cancellation: Cancellation/postponement charges and non-refundable commitments/expenses shall be payable by the client.",
    "Force Majeure: We shall not be liable for delays caused by circumstances beyond our reasonable control, including natural calamities, strikes, government restrictions, venue issues or power failures.",
    "Client Responsibility: Timely approvals, venue access, permissions, power, security and other client-side requirements shall be the client's responsibility.",
    "Liability: Any claim/deduction for a specific defective, delayed or non-delivered element shall be limited to the invoiced value of that particular element and not the total project value.",
    "Media Usage: Event photographs/videos may be used for portfolio and promotional purposes unless restricted by the client in writing.",
    "Jurisdiction: All disputes shall be subject to Mumbai, Maharashtra jurisdiction."
  ];

  // Title bar
  sheet.getRange(row, 1, 1, NCOL).merge()
    .setValue("CRES \u2014 EVENT ESTIMATE  (" + version + ")")
    .setFontWeight("bold").setFontSize(14).setFontColor("white")
    .setBackground("#10243e").setHorizontalAlignment("center");
  sheet.setRowHeight(row, 28);
  row += 2;

  // Header info
  var h = payload.header || {};
  [
    ["Company:", h.company || ""],
    ["Client:", h.client || ""],
    ["Event:", h.eventName || ""],
    ["Venue:", h.venue || ""],
    ["Pax:", h.pax || ""],
    ["Version:", version],
    ["Updated:", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")]
  ].forEach(function(pair) {
    sheet.getRange(row, 1).setValue(pair[0]).setFontWeight("bold");
    sheet.getRange(row, 2).setValue(pair[1]);
    row++;
  });
  row++;

  // Column headers
  var COLS = ["Sr.", "Particulars", "Description", "Size", "Sqft", "Days", "Qty", "Rate", "Amount", "Remark (Client)", "Remark (Doer)"];
  sheet.getRange(row, 1, 1, COLS.length).setValues([COLS])
    .setFontWeight("bold").setBackground("#10243e").setFontColor("white")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(row, 22);
  row++;

  // Group line items by dayId
  var byDay = {};
  (payload.lineItems || []).forEach(function(li) {
    if (!byDay[li.dayId]) byDay[li.dayId] = [];
    byDay[li.dayId].push(li);
  });

  var subtotal = 0;

  (payload.days || []).forEach(function(d) {
    var items = byDay[d.id] || [];
    if (items.length === 0) return;

    sheet.getRange(row, 1, 1, NCOL).merge()
      .setValue("DAY: " + d.label + (d.date ? " \u2014 " + d.date : ""))
      .setFontWeight("bold").setFontColor("white").setBackground("#c89b3c");
    sheet.setRowHeight(row, 20);
    row++;

    var byCat = {};
    items.forEach(function(it) {
      if (!byCat[it.category]) byCat[it.category] = {};
      if (!byCat[it.category][it.subSection]) byCat[it.category][it.subSection] = [];
      byCat[it.category][it.subSection].push(it);
    });

    var catLetter = 65;
    var dayTotal = 0;

    Object.keys(byCat).forEach(function(catName) {
      sheet.getRange(row, 1, 1, NCOL)
        .setValues([[String.fromCharCode(catLetter++), catName, "", "", "", "", "", "", "", "", ""]])
        .setFontWeight("bold").setBackground("#fafaf6");
      row++;

      Object.keys(byCat[catName]).forEach(function(subName) {
        if (subName) {
          sheet.getRange(row, 2).setValue(subName)
            .setFontWeight("bold").setFontColor("#c89b3c").setFontStyle("italic");
          row++;
        }

        byCat[catName][subName].forEach(function(it) {
          var rc = (it.remarkClient != null ? it.remarkClient : (it.remark || ""));
          sheet.getRange(row, 1, 1, NCOL).setValues([[
            "", it.item, it.description || "", it.size || "",
            it.sqft || "", it.days || "", it.qty || 0, it.rate || 0, it.amount || 0,
            rc, it.remarkDoer || ""
          ]]);
          dayTotal += (it.amount || 0);
          row++;
        });
      });
    });

    sheet.getRange(row, 2).setValue(d.label + " Total").setFontWeight("bold");
    sheet.getRange(row, 9).setValue(dayTotal).setFontWeight("bold").setNumberFormat("#,##0");
    sheet.getRange(row, 1, 1, NCOL).setBackground("#fbf8f0");
    row++;
    subtotal += dayTotal;
    row++;
  });

  // Logistics
  if ((payload.logistics || []).length > 0) {
    sheet.getRange(row, 1, 1, NCOL).merge()
      .setValue("LOGISTICS")
      .setFontWeight("bold").setFontColor("white").setBackground("#c89b3c");
    row++;
    var logTotal = 0;
    payload.logistics.forEach(function(lg) {
      var label = (lg.vehicle || "") + (lg.route ? " \u2014 " + lg.route : "");
      var rc = (lg.remarkClient != null ? lg.remarkClient : (lg.remark || ""));
      sheet.getRange(row, 1, 1, NCOL).setValues([[
        "", label, "", "", "", "", lg.qty || 0, lg.rate || 0, lg.amount || 0, rc, lg.remarkDoer || ""
      ]]);
      logTotal += (lg.amount || 0);
      row++;
    });
    sheet.getRange(row, 2).setValue("Logistics Total").setFontWeight("bold");
    sheet.getRange(row, 9).setValue(logTotal).setFontWeight("bold").setNumberFormat("#,##0");
    sheet.getRange(row, 1, 1, NCOL).setBackground("#fbf8f0");
    row++;
    subtotal += logTotal;
    row++;
  }

  // Summary — IGST  YA  SGST+CGST
  var s = payload.summary || {};
  var mgmtPct = s.mgmtPct || 0, hotel = s.hotelCoord || 0;
  var taxType = s.taxType || 'igst';
  var igstPct = taxType === 'igst' ? (s.igstPct || 0) : 0;
  var sgstPct = taxType === 'cs' ? (s.sgstPct || 0) : 0;
  var cgstPct = taxType === 'cs' ? (s.cgstPct || 0) : 0;
  var mgmt = subtotal * mgmtPct / 100;
  var preGst = subtotal + mgmt + hotel;
  var igst = preGst * igstPct / 100;
  var sgst = preGst * sgstPct / 100;
  var cgst = preGst * cgstPct / 100;
  var grand = preGst + igst + sgst + cgst;

  sheet.getRange(row, 1, 1, NCOL).merge()
    .setValue("SUMMARY")
    .setFontWeight("bold").setFontColor("white").setBackground("#10243e");
  row++;

  var sumRows = [
    ["Subtotal", subtotal],
    ["Management Fee (" + mgmtPct + "%)", mgmt],
    ["Hotel Coordination Charges", hotel],
    ["Sub Total before GST", preGst]
  ];
  if (taxType === 'cs') {
    sumRows.push(["SGST (" + sgstPct + "%)", sgst]);
    sumRows.push(["CGST (" + cgstPct + "%)", cgst]);
  } else {
    sumRows.push(["IGST (" + igstPct + "%)", igst]);
  }
  sumRows.forEach(function(p) {
    sheet.getRange(row, 2).setValue(p[0]);
    sheet.getRange(row, 9).setValue(p[1]).setNumberFormat("#,##0");
    row++;
  });
  sheet.getRange(row, 2).setValue("GRAND TOTAL").setFontWeight("bold").setFontColor("white");
  sheet.getRange(row, 9).setValue(grand).setFontWeight("bold").setFontColor("white").setNumberFormat("#,##0");
  sheet.getRange(row, 1, 1, NCOL).setBackground("#10243e");
  sheet.setRowHeight(row, 24);
  row += 2;

  // ===== NOTE (fixed Payment Terms & Conditions) =====
  sheet.getRange(row, 1, 1, NCOL).merge()
    .setValue("PAYMENT TERMS & CONDITIONS")
    .setFontWeight("bold").setFontColor("white").setBackground("#10243e")
    .setHorizontalAlignment("left");
  row++;
  NOTE_LINES.forEach(function(line) {
    var cell = sheet.getRange(row, 1, 1, NCOL).merge();
    cell.setValue(line).setWrap(true).setVerticalAlignment("top").setFontSize(10);
    row++;
  });

  // Column widths (10 & 11 = the two Remark columns)
  var widths = [6, 32, 22, 18, 8, 8, 8, 12, 14, 20, 20];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w * 8); });

  // Raw JSON snapshot — reliable round-trip
  var dataSheetName = "_Data_" + version;
  var existing = ss.getSheetByName(dataSheetName);
  if (existing) ss.deleteSheet(existing);
  var ds = ss.insertSheet(dataSheetName);
  ds.getRange(1, 1).setValue(JSON.stringify(payload));
  ds.hideSheet();
}


// =================== INDEX OPS ===================
function appendToIndex_(eventId, header, fileId, url, version) {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  idx.appendRow([
    eventId, header.eventName || "", header.client || "", header.venue || "",
    fileId, url, version, new Date()
  ]);
}

function updateIndex_(eventId, header, version) {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  var data = idx.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === eventId) {
      idx.getRange(i + 1, 2).setValue(header.eventName || data[i][1]);
      idx.getRange(i + 1, 3).setValue(header.client || data[i][2]);
      idx.getRange(i + 1, 4).setValue(header.venue || data[i][3]);
      // bump latest version only if new version > stored
      var stored = data[i][6] || "V0";
      var storedN = parseInt(String(stored).replace("V","")) || 0;
      var newN = parseInt(String(version).replace("V","")) || 0;
      if (newN >= storedN) idx.getRange(i + 1, 7).setValue(version);
      idx.getRange(i + 1, 8).setValue(new Date());
      return;
    }
  }
}

function findEventFile_(eventId) {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  if (!idx || idx.getLastRow() < 2) return null;
  var data = idx.getRange(2, 1, idx.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === eventId) {
      return { fileId: data[i][4], url: data[i][5] };
    }
  }
  return null;
}

// =================== UTIL ===================
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}