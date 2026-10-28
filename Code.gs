var FOLDER_NAME = "CRES Estimates";
var INDEX_SHEET = "Index";
var LIBRARY_SHEET = "Library";

// ===== VENDOR COSTING (new) =====
// Index columns 9 & 10 hold the separate Vendor Costing file per event.
var IDX_VENDOR_ID_COL = 9;   // I: VendorFileID
var IDX_VENDOR_URL_COL = 10; // J: VendorURL
var LIB_GUIDE_COL = 4;         // D: How to fill
var SNAPSHOT_CHUNK = 45000;    // a Sheets cell holds max 50,000 chars

// ===== LOGO (top-right of every costing sheet, columns J-K, rows 2-10) =====
// Logo is served from the GitHub Pages site (CRES_LOGO.png in the repo root).
// Optional: a Drive file named CRES_LOGO.png (same account) takes priority if present.
// If neither is reachable, sheets are written without a logo.
var LOGO_FILE_NAME = "CRES_LOGO.png";
var LOGO_URL = "https://mis0303.github.io/CRES-Portal-Internal/CRES_LOGO.png";

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
  ensureIndexVendorCols_();
  ensureLibraryGuideCol_();
  Browser.msgBox("Setup complete. Folder ready: " + folder.getName());
}

// Adds "VendorFileID" / "VendorURL" headers to Index if missing (safe to run many times)
function ensureIndexVendorCols_() {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  if (!idx) return;
  var h = idx.getRange(1, IDX_VENDOR_ID_COL, 1, 2).getValues()[0];
  if (h[0] !== "VendorFileID" || h[1] !== "VendorURL") {
    idx.getRange(1, IDX_VENDOR_ID_COL, 1, 2).setValues([["VendorFileID", "VendorURL"]])
      .setFontWeight("bold").setBackground("#0f6b5c").setFontColor("white");
    idx.setColumnWidth(IDX_VENDOR_ID_COL, 280);
    idx.setColumnWidth(IDX_VENDOR_URL_COL, 280);
  }
}

// Adds "How to fill" header to Library column D if missing
function ensureLibraryGuideCol_() {
  var lib = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LIBRARY_SHEET);
  if (!lib) return;
  if (lib.getRange(1, LIB_GUIDE_COL).getValue() !== "How to fill") {
    lib.getRange(1, LIB_GUIDE_COL).setValue("How to fill")
      .setFontWeight("bold").setBackground("#10243e").setFontColor("white");
    lib.setColumnWidth(LIB_GUIDE_COL, 420);
  }
}

// Map of "category|subsection|item" (lowercase) -> guide text, read from Library col D
function readGuideMap_(lib) {
  var map = {};
  if (!lib || lib.getLastRow() < 2) return map;
  var data = lib.getRange(2, 1, lib.getLastRow() - 1, LIB_GUIDE_COL).getValues();
  data.forEach(function(r) {
    var g = String(r[3] || "").trim();
    if (g) map[guideKey_(r[0], r[1], r[2])] = g;
  });
  return map;
}
function guideKey_(cat, sub, item) {
  return [cat, sub, item].map(function(s) { return String(s || "").trim().toLowerCase(); }).join("|");
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
  // Keep existing "How to fill" text so a reset does not lose it
  var guideMap = readGuideMap_(lib);
  // Clear existing rows (except header) — columns A..D so guides stay aligned
  var lastRow = lib.getLastRow();
  if (lastRow > 1) {
    lib.getRange(2, 1, lastRow - 1, LIB_GUIDE_COL).clearContent();
  }
  // Write fresh from DEFAULT_LIBRARY
  var rows = [];
  DEFAULT_LIBRARY.forEach(function(cat) {
    cat.subs.forEach(function(sub) {
      if (sub.items.length === 0) {
        rows.push([cat.name, sub.title, "", ""]);
      } else {
        sub.items.forEach(function(item) {
          rows.push([cat.name, sub.title, item, guideMap[guideKey_(cat.name, sub.title, item)] || ""]);
        });
      }
    });
  });
  if (rows.length > 0) {
    lib.getRange(2, 1, rows.length, LIB_GUIDE_COL).setValues(rows);
  }
  ensureLibraryGuideCol_();
  Browser.msgBox("Library reset. " + rows.length + " rows written.");
}

function getLibrary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib || lib.getLastRow() < 2) return { categories: [] };
  // Columns A..D (D = "How to fill" guide; blank if not filled)
  var nCols = Math.max(3, Math.min(LIB_GUIDE_COL, lib.getLastColumn()));
  var data = lib.getRange(2, 1, lib.getLastRow() - 1, nCols).getValues();

  var cats = {};
  var order = [];
  data.forEach(function(r) {
    var cat = String(r[0] || "").trim();
    var sub = String(r[1] || "").trim(); // blank means "direct items"
    var item = String(r[2] || "").trim();
    var guide = nCols >= LIB_GUIDE_COL ? String(r[3] || "").trim() : "";
    if (!cat) return;
    if (!cats[cat]) { cats[cat] = { name: cat, subOrder: [], subs: {} }; order.push(cat); }
    if (!cats[cat].subs[sub]) { cats[cat].subs[sub] = { title: sub, items: [] }; cats[cat].subOrder.push(sub); }
    if (item) cats[cat].subs[sub].items.push({ name: item, guide: guide });
  });
  var result = order.map(function(c) {
    return {
      name: c,
      subs: cats[c].subOrder.map(function(s) {
        return { title: s, items: cats[c].subs[s].items.map(function(i){ return { name: i.name, guide: i.guide }; }) };
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
  var guideMap = readGuideMap_(lib); // keep "How to fill" text attached to its item
  var rows = [];
  newLib.forEach(function(cat) {
    (cat.subs || []).forEach(function(sub) {
      (sub.items || []).forEach(function(item) {
        var itemName = typeof item === "string" ? item : (item.name || "");
        rows.push([cat.name, sub.title, itemName, guideMap[guideKey_(cat.name, sub.title, itemName)] || ""]);
      });
    });
  });

  // Clear existing (except header) and write fresh — A..D so guides stay aligned
  var lastRow = lib.getLastRow();
  if (lastRow > 1) lib.getRange(2, 1, lastRow - 1, LIB_GUIDE_COL).clearContent();
  if (rows.length > 0) lib.getRange(2, 1, rows.length, LIB_GUIDE_COL).setValues(rows);

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
    // Snapshot may be split across A1, A2, ... (long events); old files only have A1
    var n = Math.max(1, dataSheet.getLastRow());
    var raw = dataSheet.getRange(1, 1, n, 1).getValues()
      .map(function(r) { return String(r[0] || ""); }).join("");
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

  // Vendor file: remove the same version tab (never the last remaining tab)
  try {
    var iss = openVendorFile_(payload.eventId);
    if (iss) {
      var isheet = iss.getSheetByName(version);
      if (isheet && iss.getSheets().length > 1) iss.deleteSheet(isheet);
    }
  } catch (e) {}

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

  // Vendor file goes to Trash too
  var intInfo = findVendorFile_(payload.eventId);
  if (intInfo && intInfo.fileId) {
    try { DriveApp.getFileById(intInfo.fileId).setTrashed(true); } catch (e) {}
  }

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
  writeVersion_(newSS, version, clientPayload_(payload), "client");
  writeSnapshot_(newSS, version, payload); // full payload (incl. vendor amounts) for reloading
  newSS.deleteSheet(defaultSheet);

  var eventId = "EVT-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  appendToIndex_(eventId, payload.header, newSS.getId(), newSS.getUrl(), version);

  var vendorUrl = writeVendor_(eventId, version, payload);
  return { status: "success", eventId: eventId, version: version, sheetUrl: newSS.getUrl(), vendorUrl: vendorUrl };
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

  writeVersion_(ss, version, clientPayload_(payload), "client");
  writeSnapshot_(ss, version, payload);
  updateIndex_(payload.eventId, payload.header, version);
  var vendorUrl = writeVendor_(payload.eventId, version, payload);
  return { status: "success", eventId: payload.eventId, version: version, sheetUrl: ss.getUrl(), vendorUrl: vendorUrl };
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

  writeVersion_(ss, newVersion, clientPayload_(payload), "client");
  writeSnapshot_(ss, newVersion, payload);
  updateIndex_(payload.eventId, payload.header, newVersion);
  var vendorUrl = writeVendor_(payload.eventId, newVersion, payload);
  return { status: "success", eventId: payload.eventId, version: newVersion, sheetUrl: ss.getUrl(), vendorUrl: vendorUrl };
}

// =================== VENDOR COSTING (new) ===================
// Every line item / logistics row from the portal carries `amount` (client) and
// `vendorAmount` (typed in the Vendor Amount column). Both sheets use writeVersion_():
//   client sheet -> title "CLIENT COSTING", Amount column = amount
//   vendor sheet -> title "VENDOR COSTING", Amount column = vendorAmount
// Everything else (rows, days, logistics, summary maths, notes) is identical.

// Client sheet skips rows that only have a Vendor Amount and nothing else.
function hasClientValues_(r) {
  return !!(r.description || r.size || r.sqft || r.days || r.qty || r.rate ||
            (Number(r.amount) || 0) || r.remarkClient || r.remarkDoer || r.vehicle || r.route);
}
function clientPayload_(p) {
  var out = {};
  Object.keys(p).forEach(function(k) { out[k] = p[k]; });
  out.lineItems = (p.lineItems || []).filter(hasClientValues_);
  out.logistics = (p.logistics || []).filter(hasClientValues_);
  return out;
}

// Writes the version tab into the event's Vendor Costing file (created the first time).
// Returns the file URL ("" if something failed - the client save still stands).
function writeVendor_(eventId, version, payload) {
  try {
    var vss = getOrCreateVendorFile_(eventId, payload.header || {});
    var old = vss.getSheetByName(version);
    if (old) {
      if (vss.getSheets().length === 1) vss.insertSheet("_tmp");
      vss.deleteSheet(old);
    }
    writeVersion_(vss, version, payload, "vendor");
    // drop blank default / temp tabs once a real version exists
    vss.getSheets().forEach(function(s) {
      var n = s.getName();
      if ((n === "Sheet1" || n === "_tmp") && vss.getSheets().length > 1) vss.deleteSheet(s);
    });
    // keep version tabs in order V1, V2, ...
    var vs = vss.getSheets().filter(function(s) { return /^V\d+$/.test(s.getName()); })
      .sort(function(a, b) { return parseInt(a.getName().slice(1)) - parseInt(b.getName().slice(1)); });
    vs.forEach(function(s, i) { vss.setActiveSheet(s); vss.moveActiveSheet(i + 1); });
    return vss.getUrl();
  } catch (err) {
    Logger.log("Vendor sheet failed: " + err.message);
    return "";
  }
}

function getOrCreateVendorFile_(eventId, header) {
  var existing = openVendorFile_(eventId);
  if (existing) return existing;

  var folder = getOrCreateFolder_();
  var name = "Vendor Costing - " + (header.eventName || "Untitled") + " - " + (header.client || "");
  var vss = SpreadsheetApp.create(name);
  var file = DriveApp.getFileById(vss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  ensureIndexVendorCols_();
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  var data = idx.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === eventId) {
      idx.getRange(i + 1, IDX_VENDOR_ID_COL, 1, 2).setValues([[vss.getId(), vss.getUrl()]]);
      break;
    }
  }
  return vss;
}

function openVendorFile_(eventId) {
  var info = findVendorFile_(eventId);
  if (!info || !info.fileId) return null;
  try {
    var f = DriveApp.getFileById(info.fileId);
    if (f.isTrashed()) return null;
    return SpreadsheetApp.openById(info.fileId);
  } catch (e) { return null; }
}

function findVendorFile_(eventId) {
  var idx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INDEX_SHEET);
  if (!idx || idx.getLastRow() < 2 || idx.getLastColumn() < IDX_VENDOR_ID_COL) return null;
  var data = idx.getRange(2, 1, idx.getLastRow() - 1, IDX_VENDOR_URL_COL).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === eventId) {
      return { fileId: data[i][IDX_VENDOR_ID_COL - 1], url: data[i][IDX_VENDOR_URL_COL - 1] };
    }
  }
  return null;
}

// Full JSON snapshot into hidden _Data_<version>, split into 45k-char chunks down
// column A so long events never hit the 50,000-char cell limit.
function writeSnapshot_(ss, version, payload) {
  var name = "_Data_" + version;
  var existing = ss.getSheetByName(name);
  if (existing) ss.deleteSheet(existing);
  var ds = ss.insertSheet(name);
  var json = JSON.stringify(payload);
  var chunks = [];
  for (var i = 0; i < json.length; i += SNAPSHOT_CHUNK) chunks.push([json.substring(i, i + SNAPSHOT_CHUNK)]);
  if (chunks.length === 0) chunks.push([""]);
  ds.getRange(1, 1, chunks.length, 1).setNumberFormat("@").setValues(chunks);
  ds.hideSheet();
}

// =================== WRITE VERSION SHEET ===================
// mode: "client" (default) -> CLIENT COSTING with Amount
//       "vendor"            -> VENDOR COSTING with Vendor Amount (everything else identical)
function writeVersion_(ss, version, payload, mode) {
  var isVendor = mode === "vendor";
  var amtOf = function(r) { return isVendor ? (Number(r.vendorAmount) || 0) : (Number(r.amount) || 0); };
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
    .setValue("CRES \u2014 " + (isVendor ? "VENDOR COSTING" : "CLIENT COSTING") + "  (" + version + ")")
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

  addLogo_(sheet);

  // Column headers
  var COLS = ["Sr.", "Particulars", "Description", "Size", "Sqft", "Days", "Qty", "Rate", (isVendor ? "Vendor Amount" : "Amount"), "Remark (Client)", "Remark (Doer)"];
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
            it.sqft || "", it.days || "", it.qty || 0, it.rate || 0, amtOf(it),
            rc, it.remarkDoer || ""
          ]]);
          dayTotal += amtOf(it);
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
        "", label, "", "", "", "", lg.qty || 0, lg.rate || 0, amtOf(lg), rc, lg.remarkDoer || ""
      ]]);
      logTotal += amtOf(lg);
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
  var mgmt = isVendor ? 0 : subtotal * mgmtPct / 100;   // vendor costing: no management fee
  var preGst = subtotal + mgmt + hotel;
  var igst = preGst * igstPct / 100;
  var sgst = preGst * sgstPct / 100;
  var cgst = preGst * cgstPct / 100;
  var grand = preGst + igst + sgst + cgst;

  sheet.getRange(row, 1, 1, NCOL).merge()
    .setValue("SUMMARY")
    .setFontWeight("bold").setFontColor("white").setBackground("#10243e");
  row++;

  var sumRows = isVendor ? [
    ["Total", subtotal],
    ["Hotel Coordination Charges", hotel],
    ["Sub Total before GST", preGst]
  ] : [
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
  var widths = [6, 32, 22, 18, 8, 8, 10, 12, 14, 20, 20];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w * 8); });
  // (hidden _Data_ snapshot is written by the caller, client file only)
}


// =================== LOGO ===================
function getLogoBlob_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("LOGO_FILE_ID");
  if (id) {
    try { return DriveApp.getFileById(id).getBlob(); } catch (e) { props.deleteProperty("LOGO_FILE_ID"); }
  }
  var files = DriveApp.getFilesByName(LOGO_FILE_NAME);
  if (files.hasNext()) {
    var f = files.next();
    props.setProperty("LOGO_FILE_ID", f.getId());
    return f.getBlob();
  }
  try {
    var res = UrlFetchApp.fetch(LOGO_URL, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return res.getBlob().setName(LOGO_FILE_NAME);
  } catch (e) {}
  return null;
}

// Places the logo centred in the blank block J2:K10 (right of Company ... Updated).
function addLogo_(sheet) {
  try {
    var blob = getLogoBlob_();
    if (!blob) return;
    var H = 176, W = 200;                 // logo aspect ~1.14 : 1, fits rows 2-10
    var BLOCK_W = 20 * 8 * 2;             // columns J + K width (px), see widths below
    var img = sheet.insertImage(blob, 10, 2, Math.round((BLOCK_W - W) / 2), 6);
    img.setWidth(W).setHeight(H);
  } catch (err) {
    Logger.log("Logo skipped: " + err.message);
  }
}

// Run once from the editor to check the logo is found (View -> Logs).
function testLogo() {
  Logger.log(getLogoBlob_() ? "Logo found: " + LOGO_FILE_NAME : "Logo NOT found - check " + LOGO_URL);
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

// =================== HOW TO FILL (new) ===================
// Run seedHowToFill() once from the Apps Script editor. It writes guidance into
// Library column D for every item whose D cell is still empty (it never overwrites
// text you have edited). To rewrite everything from this list, run seedHowToFillOverwrite().
// Guidance follows the portal's columns for each category:
//   Size+Sqft hidden for service categories, Days hidden for Manpower / Artist /
//   Security / Collaterals / Power Supply, Size shows as "kW" for Power Supply,
//   Qty shows as "Pax" for Food & Beverage and Permissions, Rental hides only Sqft.
//   Amount = Days (blank = 1) x Qty x Rate (x Sqft only when the Sqft box is ticked).
//   Typing in Amount overrides the auto amount; you can type sums like 15000*3.
var HOW_TO_FILL = {
  "venues": {
    "_": "Qty = no. of units, Days = no. of days/nights, Rate = per unit per day.",
    "single occupancy": "Qty = no. of rooms, Days = no. of nights, Rate = per room per night (single). Write meal plan (CP/MAP) in Description.",
    "double occupancy": "Qty = no. of rooms, Days = no. of nights, Rate = per room per night (twin sharing). Write meal plan (CP/MAP) in Description.",
    "triple occupancy": "Qty = no. of rooms, Days = no. of nights, Rate = per room per night (3 sharing). Write meal plan in Description.",
    "banquets/ballrooms": "Hall name in Description. Qty = 1, Days = event days, Rate = hall rental per day (or 0 if included in F&B MPG)."
  },
  "manpower": {
    "_": "Days column is hidden here: Qty = no. of people x no. of days (e.g. 4 people x 2 days = 8). Rate = per person per day. Shift hours in Description."
  },
  "artist & entertainment": {
    "_": "Days column is hidden: Qty = no. of shows/sets (usually 1). Rate = full fee per show. Set length, language in Description. Artist travel & stay go under Travel, Lodging & Boarding.",
    "live band / base band": "Qty = 1 per show. Rate = band fee. No. of members & set length in Description. Tech rider goes under Technical > Sound.",
    "dj (bollywood / english)": "Qty = 1 per night. Rate = DJ fee incl. console if included (say so in Description). Playing hours in Description.",
    "emcee (male)": "Qty = 1 per event/day. Rate = anchor fee. Language & hours in Description.",
    "emcee (female)": "Qty = 1 per event/day. Rate = anchor fee. Language & hours in Description.",
    "dancers": "Qty = no. of acts (or no. of dancers if priced per head — say which in Description). Rate = per act / per dancer.",
    "singer": "Qty = no. of sets. Rate = fee per set. Accompanists & set length in Description."
  },
  "fabrication": {
    "_": "Size = W x H in ft (e.g. 20ft*10ft) -> Sqft fills itself. Tick the Sqft box when the rate is per sqft. Qty = no. of units. Days blank = 1 (one-time build).",
    "platform for stage": "Size = W x D in ft (e.g. 40ft*24ft), height in Description. Tick Sqft box, Rate = per sqft. Qty = 1.",
    "skirting for stage": "Size = running length x height (e.g. 88ft*3ft). Tick Sqft box, Rate = per sqft. Qty = 1.",
    "led riser": "Size = W x D in ft. Tick Sqft box if rate per sqft, else Qty = 1 and Rate = lump sum.",
    "steps": "Qty = no. of step sets. Rate = per set. Width in Description.",
    "podium": "Qty = no. of podiums. Rate = per podium incl. branding (say so in Description).",
    "standees": "Size = W x H in ft (e.g. 3ft*6ft). Qty = no. of standees. Rate = per standee (leave Sqft box unticked) or per sqft (tick it).",
    "carpet for pathway": "Size = length x width in ft. Tick Sqft box, Rate = per sqft. Colour in Description.",
    "registration backdrop": "Size = W x H in ft (e.g. 12ft*8ft). Tick Sqft box if rate per sqft. Qty = 1."
  },
  "technical": {
    "_": "Size/Sqft are hidden here. Qty = no. of units, Days = no. of show days (blank = 1), Rate = per unit per day. Model/spec in Description.",
    "main led wall": "Write size in Description (e.g. 24ft x 12ft, P2.9). Qty = total sqft (24x12 = 288), Rate = per sqft per day, Days = show days.",
    "side screen": "Size in Description. Qty = total sqft of all side screens, Rate = per sqft per day, Days = show days.",
    "comfort plasma": "Qty = no. of TVs, Rate = per TV per day, screen size (55\"/65\") in Description.",
    "timer plasma": "Qty = no. of TVs, Rate = per TV per day, screen size in Description.",
    "tops (jbl vrx)": "Qty = no. of boxes, Days = show days, Rate = per box per day.",
    "cordless mics": "Qty = no. of mics, Days = show days, Rate = per mic per day.",
    "cabling and technicians": "Qty = 1 (lump sum) or no. of technicians, Days = show days, Rate accordingly.",
    "tech rider for band": "Qty = 1, Rate = lump sum as per band's rider. Attach/mention rider in Remark (Doer).",
    "light technician & operator": "Qty = no. of people, Days = show days, Rate = per person per day.",
    "operator": "Qty = no. of operators, Days = show days, Rate = per operator per day.",
    "front goal post": "Qty = 1 set (span x height in Description), Days = show days, Rate = per set.",
    "box truss": "Qty = running feet, Days = show days, Rate = per ft. Layout in Description.",
    "t truss": "Qty = no. of T-truss, Days = show days, Rate = per truss per day."
  },
  "sfx (special effects)": {
    "_": "Qty = no. of units/machines (or no. of shots), Days = show days, Rate = per unit per day. Cue / timing in Remark (Doer).",
    "pyros": "Qty = no. of shots/pieces, Rate = per shot. Venue permission needed — note in Remark (Doer).",
    "fire show": "Qty = 1 act, Rate = per act. Fire NOC & venue OK needed.",
    "fire extinguisher": "Qty = no. of cylinders, Days = event days, Rate = per cylinder."
  },
  "security and safety": {
    "_": "Days column is hidden: Qty = no. of guards x no. of shifts/days (e.g. 6 x 2 = 12). Rate = per guard per shift (8/12 hrs — write in Description)."
  },
  "permission and liscences": {
    "_": "Pax column = no. of guests (licence slabs depend on pax). If the fee is fixed, put Pax = 1 and the full fee in Rate. Days blank = 1.",
    "iprs": "Fee depends on pax & event type: Pax = guest count with Rate per pax, or Pax = 1 with total fee in Rate.",
    "ppl": "Fee depends on pax & event type: Pax = guest count with Rate per pax, or Pax = 1 with total fee in Rate.",
    "novex": "Pax = 1, Rate = licence fee quoted by Novex (based on pax/area).",
    "police intimation": "Pax = 1, Rate = fee + liaisoning charge.",
    "foreign artist noc (mantralaya)": "Pax = 1, Rate = NOC fee + liaisoning. Start early — note lead time in Remark (Doer).",
    "fire brigade noc": "Pax = 1, Rate = NOC fee + liaisoning."
  },
  "food and beverage": {
    "_": "Pax = no. of guests, Days = no. of days this meal is served, Rate = per person per meal (taxes as per venue). Menu in Description.",
    "per plate charges": "Pax = guaranteed covers (MPG), Days = event days, Rate = per plate. Veg/non-veg split in Description.",
    "cocktail per person (imfs / imported)": "Pax = no. of drinking guests, Rate = per person for the bar duration. Hours & brands in Description.",
    "bartender": "Pax = no. of bartenders, Days = event days, Rate = per bartender per day.",
    "bar equipment": "Pax = 1 set, Rate = lump sum. List items in Description.",
    "buffet / table setup": "Pax = no. of counters/tables, Days = event days, Rate = per counter per day."
  },
  "power supply": {
    "_": "kW = capacity, Qty = no. of units, Rate = per unit. Days column is hidden — fold days into Qty or Rate.",
    "generator": "kW = genset capacity (e.g. 125). Days column is hidden: Qty = no. of gensets x days. Rate = per set per day; write if diesel is extra in Remark.",
    "inhouse power (venue charges)": "Qty = units/hours as billed by venue, Rate = per unit. Or Qty = 1 with lump sum.",
    "cabling": "Qty = 1, Rate = lump sum (or running metres in Qty).",
    "power socket": "Leave kW blank. Qty = no. of points, Rate = per point.",
    "wire managers": "Leave kW blank. Qty = no. of pieces / running ft, Rate = per piece / per ft.",
    "fire extinguisher": "Leave kW blank. Qty = no. of cylinders, Rate = per cylinder."
  },
  "collaterals": {
    "_": "Qty = no. of pieces, Rate = per piece. Size / material / print in Description. Days column is hidden."
  },
  "rental": {
    "_": "Qty = no. of units, Days = no. of rental days (blank = 1), Rate = per unit per day. Size is optional (e.g. table size).",
    "plastic chairs": "Qty = no. of chairs (with covers? say in Description), Days = rental days, Rate = per chair.",
    "round table": "Qty = no. of tables, size in Size (e.g. 5ft dia), Days = rental days, Rate = per table.",
    "portable ac": "Qty = no. of units (tonnage in Size/Description), Days = rental days, Rate = per unit per day."
  },
  "hampers and giveaway": {
    "_": "Qty = no. of pieces (usually = pax), Rate = per piece incl. branding. Leave Days blank. Spec & branding in Description."
  },
  "event coverage": {
    "_": "Qty = no. of people/cameras, Days = shoot days, Rate = per person per day. Deliverables in Description.",
    "montage / highlight video": "Qty = no. of videos, Rate = per video. Duration & delivery date in Description.",
    "online editor with setup": "Qty = 1, Days = event days, Rate = per day incl. setup.",
    "drone camera": "Qty = 1, Days = shoot days, Rate = per day. Venue/police permission note in Remark (Doer)."
  },
  "activities": {
    "_": "Qty = no. of booths/activities, Days = event days, Rate = per booth per day (say if prints are included in Description)."
  },
  "communication": {
    "_": "Qty = no. of sets/people, Days = event days, Rate = per set per day."
  },
  "misc & transport": {
    "_": "Usually a lump sum: Qty = 1, Rate = estimated amount. Write the break-up in Description.",
    "truck": "Qty = no. of trips/trucks, Rate = per trip. Route & size in Description.",
    "tbl (team stay + food)": "Qty = no. of people x nights, Rate = per person per night incl. food.",
    "crew food / promoters / hostess": "Qty = no. of meals (people x meals x days), Rate = per meal."
  },
  "awards & recognition": {
    "_": "Qty = no. of pieces, Rate = per piece. Design/engraving in Description. Leave Days blank."
  },
  "travel, lodging & boarding": {
    "_": "Qty = no. of people (or trips), Days = no. of nights/days where relevant, Rate = per person per trip/night. Route & mode in Description.",
    "travel of artists": "Qty = no. of tickets, Rate = per ticket (flight/train, class in Description). Leave Days blank.",
    "lodging for artist & food": "Qty = no. of rooms, Days = no. of nights, Rate = per room per night incl. food."
  }
};

function seedHowToFill() { seedHowToFill_(false); }
function seedHowToFillOverwrite() { seedHowToFill_(true); }

function seedHowToFill_(overwrite) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib || lib.getLastRow() < 2) { Browser.msgBox("Library sheet not found / empty."); return; }
  ensureLibraryGuideCol_();
  var n = lib.getLastRow() - 1;
  var data = lib.getRange(2, 1, n, LIB_GUIDE_COL).getValues();
  var written = 0;
  var out = data.map(function(r) {
    var cat = String(r[0] || "").trim().toLowerCase();
    var item = String(r[2] || "").trim().toLowerCase();
    var current = String(r[3] || "").trim();
    if (!item) return [current];
    if (current && !overwrite) return [current];
    var g = HOW_TO_FILL[cat];
    var text = g ? (g[item] || g["_"] || "") : "";
    if (!text) text = "Qty = no. of units, Days = no. of days (blank = 1), Rate = per unit per day. Details in Description.";
    if (text !== current) written++;
    return [text];
  });
  lib.getRange(2, LIB_GUIDE_COL, n, 1).setValues(out);
  Browser.msgBox("How to fill: " + written + " cells written.");
}
