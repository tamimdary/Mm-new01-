/**
 * Madrasah Management System - Google Sheets Cloud Sync & Backup Apps Script
 *
 * This script is deployed as a Web App on Google Apps Script and acts as the fully
 * compatible cloud database backup provider for the Madrasah Management System.
 * 
 * It supports connection tests, full state/database synchronization (backup and restore),
 * and molecular row CRUD operations (addRow, updateRow, deleteRow) to maintain absolute data integrity
 * in the backing spreadsheet.
 */

// Define all required sheet tabs and their exact column structures
var SHEET_CONFIGS = {
  "Config": [
    "Setting Key", 
    "Value"
  ],
  "Raw_Backup": [
    "Timestamp / Date", 
    "Backup Version", 
    "Raw JSON State Data"
  ],
  "Students": [
    "ID (দাখেলা আইডি)", 
    "Name (নাম)", 
    "Gender (লিঙ্গ)", 
    "Section (বিভাগ)", 
    "Jamat (জামাত)", 
    "Father (পিতা)", 
    "Mother (মাতা)", 
    "Guardian (অভিভাবক)", 
    "DOB (জন্ম তারিখ)", 
    "POB (জন্মস্থান)", 
    "Phone (ফোন নম্বর)", 
    "Address (ঠিকানা)", 
    "Monthly Fee (মাসিক বেতন)", 
    "Admission Date (ভর্তি তারিখ)", 
    "Resident (আবাসিক)", 
    "NID / Birth Cert (এনআইডি)", 
    "Blood Group (রক্তের গ্রুপ)", 
    "Status (অবস্থা)"
  ],
  "Teachers": [
    "ID (আইডি)", 
    "Name (নাম)", 
    "Designation (পদবি)", 
    "Section (বিভাগ)", 
    "Join Date (যোগদানের তারিখ)", 
    "Phone (ফোন নম্বর)", 
    "Address (ঠিকানা)", 
    "Salary (মাসিক বেতন)", 
    "Status (অবস্থা)"
  ],
  "Hajira_Attendance": [
    "Date (তারিখ)", 
    "Student ID (ছাত্র আইডি)", 
    "Status (হাজিরা অবস্থা)"
  ],
  "Fees": [
    "Receipt ID (রশিদ আইডি)", 
    "Student ID (ছাত্র আইডি)", 
    "Month (মাস)", 
    "Amount (টাকা)", 
    "Payment Date (তারিখ)", 
    "Method (মাধ্যম)", 
    "Receipt No (রশিদ নং)"
  ],
  "Salaries": [
    "Salary ID (বেতন আইডি)", 
    "Teacher ID (শিক্ষক আইডি)", 
    "Teacher Name (শিক্ষকের নাম)", 
    "Month (মাস)", 
    "Amount (টাকা)", 
    "Payment Date (তারিখ)", 
    "Notes (মন্তব্য)"
  ],
  "Expenses": [
    "ID (আইডি)", 
    "Transaction Type (ধরণ)", 
    "Category (খাত)", 
    "Description (বিবরণ)", 
    "Amount (পরিমাণ)", 
    "Date (তারিখ)"
  ],
  "Boarding": [
    "Boarding ID (আইডি)", 
    "Date (তারিখ)", 
    "Type (বাজার ধরণ)", 
    "Total Cost (মোট বাজার খরচ)", 
    "Purchase Items List (ক্রয়কৃত সামগ্রী বিবরণী)"
  ],
  "Donors": [
    "ID (আইডি)", 
    "Name (দাতার নাম)", 
    "Phone (ফোন)", 
    "Address (ঠিকানা)", 
    "Membership Type (দাতার ধরণ)", 
    "Committed Amount (অঙ্গীকারকৃত পরিমাণ)"
  ],
  "Donations": [
    "Donation ID (আইডি)", 
    "Donor ID (দাতা আইডি)", 
    "Amount (টাকার পরিমাণ)", 
    "Date (তারিখ)"
  ],
  "Notices": [
    "Notice ID (আইডি)", 
    "Title (শিরোনাম)", 
    "Content (বিবরণ)", 
    "Severity Type (ধরণ)", 
    "Target Audience (লক্ষ্য গ্রুপ)", 
    "Notification Channel (মাধ্যম)", 
    "Date (তারিখ)"
  ],
  "Hifz_Records": [
    "Record ID (আইডি)", 
    "Student ID (ছাত্র আইডি)", 
    "Date (তারিখ)", 
    "Teacher Name (শিক্ষক)", 
    "Reading Type (ধরণ)", 
    "Sabak Para (সবক পারা)", 
    "Sabak From (সবক শুরু আয়াত/পৃষ্ঠা)", 
    "Sabak To (সবক শেষ আয়াত/পৃষ্ঠা)", 
    "Sabak Qty (পরিমাণ)", 
    "Sabki Para (সবকী পারা)", 
    "Sabki From (সবকী শুরু)", 
    "Sabki To (সবকী শেষ)", 
    "Sabki Qty (পরিমাণ)", 
    "Amukhta Para (আমুখতা পারা)", 
    "Amukhta From (আমুখতা শুরু)", 
    "Amukhta To (আমুখতা শেষ)", 
    "Amukhta Qty (পরিমাণ)", 
    "Performance Status (মান)", 
    "Remark (মন্তব্য)"
  ],
  "Sync_Logs": [
    "Timestamp", 
    "Operation Type", 
    "IP / Caller Context", 
    "Status Summary"
  ]
};

/**
 * Automatically sets up all requested spreadsheet tabs and configures headers.
 * Safe to be executed repeatedly without destroying existing data.
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var name in SHEET_CONFIGS) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    var headers = SHEET_CONFIGS[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
         .setFontWeight("bold")
         .setBackground("#e2f0d9")
         .setFontColor("#385723");
    try {
      sheet.setFrozenRows(1);
    } catch(e) {}
  }
}

/**
 * Handle HTTP GET Requests
 * Supports Action: "test", and Action: "load" to fetch the state
 */
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : "";
  
  // Make sure all target tables are initialized beautifully
  setupSheets();
  
  if (action === "test") {
    return ContentService.createTextOutput(JSON.stringify({
      "success": true,
      "status": "ok",
      "message": "Google Sheets Connection Tested and Initial Setup Configured Successfully!",
      "timestamp": new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "loadAll" || action === "load") {
    try {
      var fullState = loadAll();
      return ContentService.createTextOutput(JSON.stringify({
        "success": true,
        "status": "ok",
        "data": fullState,
        "timestamp": new Date().toISOString()
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        "success": false,
        "status": "error",
        "message": "Failed to load database state: " + err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({
    "success": false,
    "status": "error",
    "message": "Invalid or missing GET action parameter."
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle HTTP POST Requests
 * Processes backend state updates ('sync'), custom saves ('saveAll'), and atomic row manipulation operations (addRow, updateRow, deleteRow).
 */
function doPost(e) {
  var responsePayload = {};
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Missing parameters or empty post request payload");
    }
    
    var requestBody = JSON.parse(e.postData.contents);
    var action = requestBody.action;
    
    // Auto-bootstrap sheets to ensure tabs are valid
    setupSheets();
    
    if (action === "sync" || action === "saveAll") {
      var rawData = requestBody.data;
      if (!rawData) {
        throw new Error("Missing required 'data' field containing serialized application state");
      }
      
      saveAll(rawData);
      
      responsePayload = {
        "success": true,
        "status": "ok",
        "message": "Cloud backup and state sync completed successfully at " + new Date().toLocaleString(),
        "timestamp": new Date().toISOString()
      };
      
    } else if (action === "addRow") {
      var sheetName = requestBody.sheetName;
      var rowData = requestBody.rowData;
      if (!sheetName || !rowData) throw new Error("Missing sheetName or rowData parameters");
      
      addRow(sheetName, rowData);
      responsePayload = { "success": true, "message": "Successfully appended row to " + sheetName };
      
    } else if (action === "updateRow") {
      var sheetName = requestBody.sheetName;
      var idColumnIndex = requestBody.idColumnIndex !== undefined ? requestBody.idColumnIndex : 0;
      var rowId = requestBody.rowId;
      var rowData = requestBody.rowData;
      if (!sheetName || rowId === undefined || !rowData) throw new Error("Missing correct update specifications");
      
      var updated = updateRow(sheetName, idColumnIndex, rowId, rowData);
      responsePayload = { "success": updated, "message": updated ? "Successfully updated row in " + sheetName : "Row matching ID " + rowId + " not found" };
      
    } else if (action === "deleteRow") {
      var sheetName = requestBody.sheetName;
      var idColumnIndex = requestBody.idColumnIndex !== undefined ? requestBody.idColumnIndex : 0;
      var rowId = requestBody.rowId;
      if (!sheetName || rowId === undefined) throw new Error("Missing details for action 'deleteRow'");
      
      var deleted = deleteRow(sheetName, idColumnIndex, rowId);
      responsePayload = { "success": deleted, "message": deleted ? "Successfully deleted row from " + sheetName : "Row matching ID " + rowId + " not found" };
      
    } else {
      throw new Error("Unsupported POST action action requested: " + action);
    }
    
  } catch (err) {
    appendLog("sync_error", "Error", err.toString());
    responsePayload = {
      "success": false,
      "status": "error",
      "message": "Apps Script Operation Failure: " + err.toString(),
      "timestamp": new Date().toISOString()
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(responsePayload))
                       .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Re-reads current state from spreadsheet tabs and parses back into a cleanly structured JSON database matching S object format.
 */
function loadAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var state = {
    cfg: {},
    students: [],
    teachers: [],
    hajira: [],
    fees: [],
    salaries: [],
    expenses: [],
    boarding: [],
    donors: [],
    donations: [],
    notices: [],
    hifzRecs: []
  };
  
  // 1. Config loading
  var configSheet = ss.getSheetByName("Config");
  if (configSheet) {
    var rawConfig = configSheet.getDataRange().getValues();
    state.cfg.pw = {};
    state.cfg.receipt = {};
    for (var i = 1; i < rawConfig.length; i++) {
      var k = rawConfig[i][0];
      var v = rawConfig[i][1];
      if (k) {
        // Simple type reconstruction
        if (v === "true") v = true;
        else if (v === "false") v = false;
        
        state.cfg[k] = v;
      }
    }
  }
  
  // Helper to map values easily
  var parseRows = function(sheetName, mapper) {
    var s = ss.getSheetByName(sheetName);
    if (!s) return [];
    var vals = s.getDataRange().getValues();
    var list = [];
    for (var r = 1; r < vals.length; r++) {
      var item = mapper(vals[r]);
      if (item) list.push(item);
    }
    return list;
  };

  // 2. Load entities
  state.students = parseRows("Students", function(r) {
    return {
      id: r[0], name: r[1], gender: r[2], section: r[3], jamat: r[4], 
      father: r[5], mother: r[6], guardian: r[7], dob: r[8], pob: r[9],
      phone: r[10], address: r[11], salary: r[12], ভর্তি_তারিখ: r[13],
      resident: r[14], nid: r[15], blood: r[16], status: r[17]
    };
  });
  
  state.teachers = parseRows("Teachers", function(r) {
    return {
      id: r[0], name: r[1], designation: r[2], section: r[3],
      joinDate: r[4], phone: r[5], address: r[6], salary: r[7], status: r[8]
    };
  });
  
  state.hajira = parseRows("Hajira_Attendance", function(r) {
    return { date: r[0], studentId: r[1], status: r[2] };
  });
  
  state.fees = parseRows("Fees", function(r) {
    return { id: r[0], studentId: r[1], month: r[2], amount: r[3], date: r[4], method: r[5], receiptNo: r[6] };
  });
  
  state.salaries = parseRows("Salaries", function(r) {
    return { id: r[0], teacherId: r[1], teacherName: r[2], month: r[3], amount: r[4], paymentDate: r[5], notes: r[6] };
  });
  
  state.expenses = parseRows("Expenses", function(r) {
    return { id: r[0], type: r[1], category: r[2], description: r[3], amount: r[4], date: r[5] };
  });
  
  state.boarding = parseRows("Boarding", function(r) {
    // Attempt parsing list items string back to structured array
    var itemsList = [];
    try {
      if (r[4]) {
        var parts = r[4].split(", ");
        itemsList = parts.map(function(p) {
          return { name: p }; // Simple recovery representation
        });
      }
    } catch(e) {}
    return { id: r[0], date: r[1], type: r[2], total: r[3], items: itemsList };
  });
  
  state.donors = parseRows("Donors", function(r) {
    return { id: r[0], name: r[1], phone: r[2], address: r[3], type: r[4], amount: r[5] };
  });
  
  state.donations = parseRows("Donations", function(r) {
    return { id: r[0], donorId: r[1], amount: r[2], date: r[3] };
  });
  
  state.notices = parseRows("Notices", function(r) {
    return { id: r[0], title: r[1], content: r[2], type: r[3], target: r[4], channel: r[5], date: r[6] };
  });
  
  state.hifzRecs = parseRows("Hifz_Records", function(r) {
    return {
      id: r[0], studentId: r[1], date: r[2], teacher: r[3], type: r[4],
      sabakPara: r[5], sabakFrom: r[6], sabakTo: r[7], sabakQty: r[8],
      sabkiPara: r[9], sabkiFrom: r[10], sabkiTo: r[11], sabkiQty: r[12],
      amukhtaPara: r[13], amukhtaFrom: r[14], amukhtaTo: r[15], amukhtaQty: r[16],
      status: r[17], remark: r[18]
    };
  });
  
  return state;
}

/**
 * Parses and writes raw DB state data cleanly across all relative spreadsheet tabs.
 */
function saveAll(rawStateString) {
  var stateS = JSON.parse(rawStateString);
  
  // Backup raw JSON record state first
  overwriteTabWithData("Raw_Backup", [
    [new Date().toISOString(), "v1.0.0", rawStateString]
  ]);
  
  // 1. Sync Config block
  if (stateS.cfg) {
    var configRows = [];
    for (var key in stateS.cfg) {
      if (stateS.cfg[key] !== null && typeof stateS.cfg[key] !== 'object') {
        configRows.push([key, stateS.cfg[key]]);
      }
    }
    overwriteTabWithData("Config", configRows);
  }
  
  // 2. Sync Student records
  if (stateS.students && Array.isArray(stateS.students)) {
    overwriteTabWithData("Students", stateS.students.map(function(s) {
      return [
        s.id || "", s.name || "", s.gender || "", s.section || "", s.jamat || "",
        s.father || "", s.mother || "", s.guardian || "", s.dob || "", s.pob || "",
        s.phone || "", s.address || "", s.salary || "০", s.ভর্তি_তারিখ || "",
        s.resident || "না", s.nid || "", s.blood || "", s.status || "সক্রিয়"
      ];
    }));
  }
  
  // 3. Sync Teacher records
  if (stateS.teachers && Array.isArray(stateS.teachers)) {
    overwriteTabWithData("Teachers", stateS.teachers.map(function(t) {
      return [
        t.id || "", t.name || "", t.designation || "", t.section || "",
        t.joinDate || "", t.phone || "", t.address || "", t.salary || "০", t.status || "সক্রিয়"
      ];
    }));
  }
  
  // 4. Sync Attendance Log
  if (stateS.hajira && Array.isArray(stateS.hajira)) {
    overwriteTabWithData("Hajira_Attendance", stateS.hajira.map(function(h) {
      return [h.date || "", h.studentId || "", h.status || ""];
    }));
  }
  
  // 5. Sync Fees Receipt collection history
  if (stateS.fees && Array.isArray(stateS.fees)) {
    overwriteTabWithData("Fees", stateS.fees.map(function(f) {
      return [f.id || "", f.studentId || "", f.month || "", f.amount || "০", f.date || "", f.method || "", f.receiptNo || ""];
    }));
  }
  
  // 6. Sync Salary transactions logs
  if (stateS.salaries && Array.isArray(stateS.salaries)) {
    overwriteTabWithData("Salaries", stateS.salaries.map(function(sa) {
      return [sa.id || "", sa.teacherId || "", sa.teacherName || "", sa.month || "", sa.amount || "০", sa.paymentDate || "", sa.notes || ""];
    }));
  }
  
  // 7. Sync Expenses logs
  if (stateS.expenses && Array.isArray(stateS.expenses)) {
    overwriteTabWithData("Expenses", stateS.expenses.map(function(ex) {
      return [ex.id || "", ex.type || "ব্যয়", ex.category || "", ex.description || "", ex.amount || "০", ex.date || ""];
    }));
  }
  
  // 8. Boarding Marketplace Bazaar Purchase List
  if (stateS.boarding && Array.isArray(stateS.boarding)) {
    overwriteTabWithData("Boarding", stateS.boarding.map(function(b) {
      var itemization = "";
      try {
        if (b.items && Array.isArray(b.items)) {
          itemization = b.items.map(function(it) {
            return it.name + " (" + (it.qty || "") + (it.unit || "") + "x" + (it.price || "") + "=" + (it.total || "") + ")";
          }).join(", ");
        }
      } catch(e) {}
      return [b.id || "", b.date || "", b.type || "", b.total || "০", itemization];
    }));
  }
  
  // 9. Sync Donors details
  if (stateS.donors && Array.isArray(stateS.donors)) {
    overwriteTabWithData("Donors", stateS.donors.map(function(d) {
      return [d.id || "", d.name || "", d.phone || "", d.address || "", d.type || "", d.amount || "০"];
    }));
  }
  
  // 10. Sync Donations histories
  if (stateS.donations && Array.isArray(stateS.donations)) {
    overwriteTabWithData("Donations", stateS.donations.map(function(dn) {
      return [dn.id || "", dn.donorId || "", dn.amount || "০", dn.date || ""];
    }));
  }
  
  // 11. Sync Official Notices
  if (stateS.notices && Array.isArray(stateS.notices)) {
    overwriteTabWithData("Notices", stateS.notices.map(function(n) {
      return [n.id || "", n.title || "", n.content || "", n.type || "", n.target || "", n.channel || "", n.date || ""];
    }));
  }
  
  // 12. Hifz Diary Progress Records
  if (stateS.hifzRecs && Array.isArray(stateS.hifzRecs)) {
    overwriteTabWithData("Hifz_Records", stateS.hifzRecs.map(function(r) {
      return [
        r.id || "", r.studentId || "", r.date || "", r.teacher || "", r.type || "",
        r.sabakPara || "", r.sabakFrom || "", r.sabakTo || "", r.sabakQty || "",
        r.sabkiPara || "", r.sabkiFrom || "", r.sabkiTo || "", r.sabkiQty || "",
        r.amukhtaPara || "", r.amukhtaFrom || "", r.amukhtaTo || "", r.amukhtaQty || "",
        r.status || "", r.remark || ""
      ];
    }));
  }
  
  // Append standard system success entry Log
  appendLog("sync", "Success", "Synced " + Object.keys(SHEET_CONFIGS).length + " standard table tabs successfully.");
}

/**
 * Appends a new molecular row data item format to the target tab sheet
 */
function addRow(sheetName, dataRowArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheets();
  }
  sheet.appendRow(dataRowArray);
  appendLog("addRow", "Success", "Appended atomic row inside " + sheetName);
}

/**
 * Molecular updates row parameters matching unique identifier inside the column index
 * Returns boolean indication of update action status.
 */
function updateRow(sheetName, idColumnIndex, rowId, updatedRowArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idColumnIndex]).trim() === String(rowId).trim()) {
      var rowNumber = r + 1;
      // Align array elements to fit target spreadsheet range dimensions cleanly
      var headersCount = SHEET_CONFIGS[sheetName].length;
      var writeValues = [];
      for (var col = 0; col < headersCount; col++) {
        writeValues.push(updatedRowArray[col] !== undefined ? updatedRowArray[col] : "");
      }
      sheet.getRange(rowNumber, 1, 1, headersCount).setValues([writeValues]);
      appendLog("updateRow", "Success", "Updated record ID: " + rowId + " in " + sheetName);
      return true;
    }
  }
  return false;
}

/**
 * Molecular row removal matches unique identifier key inside the column index
 * Returns status indication of row purge request.
 */
function deleteRow(sheetName, idColumnIndex, rowId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  
  var values = sheet.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idColumnIndex]).trim() === String(rowId).trim()) {
      sheet.deleteRow(r + 1);
      appendLog("deleteRow", "Success", "Purged record row master matching ID: " + rowId + " from sheet " + sheetName);
      return true;
    }
  }
  return false;
}

/**
 * Overwrites individual sheet data cleanly while maintaining active headers
 */
function overwriteTabWithData(sheetName, dataRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheets();
  }
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  
  if (dataRows && dataRows.length > 0) {
    var numRows = dataRows.length;
    var numCols = SHEET_CONFIGS[sheetName].length;
    var formattedRows = dataRows.map(function(row) {
      var arr = [];
      for (var i = 0; i < numCols; i++) {
        arr.push(row[i] !== undefined ? row[i] : "");
      }
      return arr;
    });
    sheet.getRange(2, 1, numRows, numCols).setValues(formattedRows);
  }
}

/**
 * Appends standard entry in sync logger tab for quick system monitoring
 */
function appendLog(opType, status, summary) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Sync_Logs");
    if (!logSheet) return;
    
    logSheet.appendRow([
      new Date().toLocaleString('bn-BD'),
      opType,
      "API/AppsScript Endpoint",
      summary
    ]);
  } catch(e) {}
}
