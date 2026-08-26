/**
 * =========================================================================
 * EXIUM MUPS - SWEATER CAMPAIGN 2026 (4Q'26)
 * GOOGLE APPS SCRIPT BACKEND (WITH JSONP + DIRECT CORS SUPPORT)
 * =========================================================================
 */

function getOrCreateSheets(ss) {
  var sheets = ss.getSheets();
  var sheet1 = ss.getSheetByName("Gyne Core Doctor (Family)");
  var sheet2 = ss.getSheetByName("Core Doctor Maximization");

  if (!sheet1) {
    if (sheets.length > 0) {
      sheet1 = sheets[0];
      sheet1.setName("Gyne Core Doctor (Family)");
    } else {
      sheet1 = ss.insertSheet("Gyne Core Doctor (Family)");
    }
  }

  if (!sheet2) {
    if (sheets.length > 1 && sheets[1] !== sheet1) {
      sheet2 = sheets[1];
      sheet2.setName("Core Doctor Maximization");
    } else {
      sheet2 = ss.insertSheet("Core Doctor Maximization");
    }
  }

  return { sheet1: sheet1, sheet2: sheet2 };
}

/**
 * Run this function once from Apps Script editor to instantly restore 
 * the Territory Status column and set 3-Doctor layout on Google Sheet!
 */
function fixAndRestoreStatusColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pair = getOrCreateSheets(ss);
  var sheet1 = pair.sheet1;
  var sheet2 = pair.sheet2;

  restoreHeaders();

  // Populate Territory Status Formula in Column S (Column 19) for Core Doctor Maximization
  if (sheet2) {
    var lr2 = sheet2.getLastRow();
    if (lr2 > 2) {
      var formulas2 = [];
      for (var r = 3; r <= lr2; r++) {
        formulas2.push([
          '=IF(AND(G' + r + '<>"",LEN(H' + r + ')=6,I' + r + '<>"",J' + r + '<>"",K' + r + '<>"",LEN(L' + r + ')=6,M' + r + '<>"",N' + r + '<>"",O' + r + '<>"",LEN(P' + r + ')=6,Q' + r + '<>"",R' + r + '<>""), "Complete", IF(OR(G' + r + '<>"",H' + r + '<>"",I' + r + '<>"",J' + r + '<>"",K' + r + '<>"",L' + r + '<>"",M' + r + '<>"",N' + r + '<>"",O' + r + '<>"",P' + r + '<>"",Q' + r + '<>"",R' + r + '<>""), "In Progress", "Not Started"))'
        ]);
      }
      sheet2.getRange(3, 19, lr2 - 2, 1).setFormulas(formulas2)
        .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
    }
  }

  // Populate Territory Status Formula in Column Q (Column 17) for Gyne Core Doctor (Family)
  if (sheet1) {
    var lr1 = sheet1.getLastRow();
    if (lr1 > 2) {
      var formulas1 = [];
      for (var r1 = 3; r1 <= lr1; r1++) {
        formulas1.push([
          '=IF(AND(G' + r1 + '<>"",LEN(H' + r1 + ')=6,I' + r1 + '<>"",J' + r1 + '<>"",K' + r1 + '<>"",L' + r1 + '<>"",M' + r1 + '<>"",N' + r1 + '<>""), "Complete", IF(OR(G' + r1 + '<>"",H' + r1 + '<>"",I' + r1 + '<>"",K' + r1 + '<>"",M' + r1 + '<>""), "In Progress", "Not Started"))'
        ]);
      }
      sheet1.getRange(3, 17, lr1 - 2, 1).setFormulas(formulas1)
        .setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
    }
  }

  ss.toast("Territory Status column & formulas successfully restored!", "Success", 5);
}

function restoreHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pair = getOrCreateSheets(ss);
  var sheet1 = pair.sheet1;
  var sheet2 = pair.sheet2;

  if (sheet1) {
    sheet1.getRange("A1:Z1").breakApart();
    sheet1.getRange("A1:F1").merge().setValue("TERRITORY INFORMATION (EXIUM FIELD FORCE LIST)")
      .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet1.getRange("G1:P1").merge().setValue("CAMPAIGN 1: GYNE CORE DOCTOR DEVELOPMENT (FAMILY PACKAGE - 4 SWEATERS / TERRITORY)")
      .setBackground("#0F766E").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet1.getRange("Q1").setValue("STATUS")
      .setBackground("#047857").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    var h1 = [
      "Zone", "SAP Region Code", "Region", "Regional Head", "SAP Territory Code", "Territory",
      "Doctor Name", "Doctor RPL ID",
      "Sweater 1", "Size 1", "Sweater 2", "Size 2", "Sweater 3", "Size 3", "Sweater 4", "Size 4",
      "Territory Status"
    ];
    sheet1.getRange(2, 1, 1, h1.length).setValues([h1])
      .setBackground("#F1F5F9").setFontColor("#0F172A").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    
    sheet1.setFrozenRows(2);
    sheet1.setRowHeights(1, 2, 28);
  }

  if (sheet2) {
    sheet2.getRange("A1:Z1").breakApart();

    // Ensure exactly 19 columns
    var maxCol = sheet2.getMaxColumns();
    if (maxCol < 19) {
      sheet2.insertColumnsAfter(maxCol, 19 - maxCol);
    } else if (maxCol > 19) {
      sheet2.deleteColumns(20, maxCol - 19);
    }

    sheet2.getRange("A1:F1").merge().setValue("TERRITORY INFORMATION (EXIUM FIELD FORCE LIST)")
      .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet2.getRange("G1:R1").merge().setValue("CAMPAIGN 2: CORE DOCTOR MAXIMIZATION (1 SWEATER / DOCTOR - 3 DOCTORS / TERRITORY)")
      .setBackground("#6B21A8").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet2.getRange("S1").setValue("STATUS")
      .setBackground("#047857").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");

    var h2 = [
      "Zone", "SAP Region Code", "Region", "Regional Head", "SAP Territory Code", "Territory",
      "Doctor 1 Name", "Doctor 1 RPL ID", "Sweater 1", "Size 1",
      "Doctor 2 Name", "Doctor 2 RPL ID", "Sweater 2", "Size 2",
      "Doctor 3 Name", "Doctor 3 RPL ID", "Sweater 3", "Size 3",
      "Territory Status"
    ];
    sheet2.getRange(2, 1, 1, h2.length).setValues([h2])
      .setBackground("#F1F5F9").setFontColor("#0F172A").setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle");
    
    sheet2.setFrozenRows(2);
    sheet2.setRowHeights(1, 2, 28);
  }
}

function doGet(e) {
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : "";
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "pull_data";

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pair = getOrCreateSheets(ss);
    var sheet1 = pair.sheet1;
    var sheet2 = pair.sheet2;

    if (action === "test") {
      return sendResponse({ status: "success", message: "Google Sheet connection successful!" }, callback);
    }

    var store = {};
    var populatedCount = 0;

    if (sheet1) {
      var lr1 = sheet1.getLastRow();
      if (lr1 > 2) {
        var v1 = sheet1.getRange(3, 1, lr1 - 2, 16).getValues();
        for (var i = 0; i < v1.length; i++) {
          var code = String(v1[i][4]).trim();
          if (code) {
            if (!store[code]) store[code] = {};
            store[code].c1_doc_name = String(v1[i][6] || "").trim();
            store[code].c1_doc_rpl = String(v1[i][7] || "").trim();
            store[code].c1_m1_sweater = String(v1[i][8] || "").trim();
            store[code].c1_m1_size = String(v1[i][9] || "").trim();
            store[code].c1_m2_sweater = String(v1[i][10] || "").trim();
            store[code].c1_m2_size = String(v1[i][11] || "").trim();
            store[code].c1_m3_sweater = String(v1[i][12] || "").trim();
            store[code].c1_m3_size = String(v1[i][13] || "").trim();
            store[code].c1_m4_sweater = String(v1[i][14] || "").trim();
            store[code].c1_m4_size = String(v1[i][15] || "").trim();
          }
        }
      }
    }

    if (sheet2) {
      var lr2 = sheet2.getLastRow();
      if (lr2 > 2) {
        var v2 = sheet2.getRange(3, 1, lr2 - 2, 18).getValues();
        for (var j = 0; j < v2.length; j++) {
          var code2 = String(v2[j][4]).trim();
          if (code2) {
            if (!store[code2]) store[code2] = {};
            store[code2].c2_d1_name = String(v2[j][6] || "").trim();
            store[code2].c2_d1_rpl = String(v2[j][7] || "").trim();
            store[code2].c2_d1_sweater = String(v2[j][8] || "").trim();
            store[code2].c2_d1_size = String(v2[j][9] || "").trim();
            store[code2].c2_d2_name = String(v2[j][10] || "").trim();
            store[code2].c2_d2_rpl = String(v2[j][11] || "").trim();
            store[code2].c2_d2_sweater = String(v2[j][12] || "").trim();
            store[code2].c2_d2_size = String(v2[j][13] || "").trim();
            store[code2].c2_d3_name = String(v2[j][14] || "").trim();
            store[code2].c2_d3_rpl = String(v2[j][15] || "").trim();
            store[code2].c2_d3_sweater = String(v2[j][16] || "").trim();
            store[code2].c2_d3_size = String(v2[j][17] || "").trim();
            store[code2].c2_d4_name = "";
            store[code2].c2_d4_rpl = "";
            store[code2].c2_d4_sweater = "";
            store[code2].c2_d4_size = "";
          }
        }
      }
    }

    for (var k in store) {
      var item = store[k];
      if (item.c1_doc_name || item.c1_doc_rpl || item.c1_m1_sweater || item.c2_d1_name || item.c2_d1_rpl || item.c2_d1_sweater) {
        populatedCount++;
      }
    }

    return sendResponse({
      status: "success",
      store: store,
      total_territories: Object.keys(store).length,
      populated_territories: populatedCount
    }, callback);

  } catch (err) {
    return sendResponse({ status: "error", message: err.toString() }, callback);
  }
}

function sendResponse(obj, callback) {
  var jsonStr = JSON.stringify(obj);
  if (callback && callback.trim()) {
    return ContentService
      .createTextOutput(callback.trim() + "(" + jsonStr + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawText = (e && e.postData && e.postData.contents) ? e.postData.contents : "";
    var postData = rawText ? JSON.parse(rawText) : (e ? e.parameter : {});
    var action = postData.action || "save_territory";
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "save_territory") {
      updateTerritoryInSheets(ss, String(postData.sap_territory_code), postData.data);
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "save_batch") {
      var batch = postData.batch || {};
      for (var tCode in batch) {
        updateTerritoryInSheets(ss, String(tCode), batch[tCode]);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success", count: Object.keys(batch).length })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "delete_region") {
      clearRegionSheetData(ss, String(postData.sap_region_code));
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Region cleared" })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "reset_all") {
      clearAllSheetData(ss);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "All sheet data reset" })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unknown action" })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function updateTerritoryInSheets(ss, terrCode, d) {
  if (!d || !terrCode) return;
  var pair = getOrCreateSheets(ss);
  var sheet1 = pair.sheet1;
  var sheet2 = pair.sheet2;

  if (sheet1) {
    var lr1 = sheet1.getLastRow();
    if (lr1 > 2) {
      var codes1 = sheet1.getRange(3, 5, lr1 - 2, 1).getValues();
      for (var i = 0; i < codes1.length; i++) {
        if (String(codes1[i][0]).trim() === String(terrCode).trim()) {
          var c1Mandatory = Boolean(
            d.c1_doc_name && d.c1_doc_rpl && String(d.c1_doc_rpl).length === 6 &&
            d.c1_m1_sweater && d.c1_m1_size &&
            d.c1_m2_sweater && d.c1_m2_size &&
            d.c1_m3_sweater && d.c1_m3_size
          );
          var c1M4HasAny = Boolean(d.c1_m4_sweater || d.c1_m4_size);
          var c1M4Ok = !c1M4HasAny || Boolean(d.c1_m4_sweater && d.c1_m4_size);
          var c1Complete = c1Mandatory && c1M4Ok;
          var c1HasAny = Boolean(d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater || d.c1_m1_size || d.c1_m2_sweater || d.c1_m2_size || d.c1_m3_sweater || d.c1_m3_size || d.c1_m4_sweater || d.c1_m4_size);
          var c1Status = c1Complete ? "Complete" : (c1HasAny ? "In Progress" : "Not Started");

          var v1 = [
            d.c1_doc_name||"", d.c1_doc_rpl||"",
            d.c1_m1_sweater||"", d.c1_m1_size||"",
            d.c1_m2_sweater||"", d.c1_m2_size||"",
            d.c1_m3_sweater||"", d.c1_m3_size||"",
            d.c1_m4_sweater||"", d.c1_m4_size||"",
            c1Status
          ];
          sheet1.getRange(i + 3, 7, 1, 11).setValues([v1]);
          break;
        }
      }
    }
  }

  if (sheet2) {
    var lr2 = sheet2.getLastRow();
    if (lr2 > 2) {
      var codes2 = sheet2.getRange(3, 5, lr2 - 2, 1).getValues();
      for (var j = 0; j < codes2.length; j++) {
        if (String(codes2[j][0]).trim() === String(terrCode).trim()) {
          var c2Doc1Ok = Boolean(d.c2_d1_name && d.c2_d1_rpl && String(d.c2_d1_rpl).length === 6 && d.c2_d1_sweater && d.c2_d1_size);
          var c2Doc2Ok = Boolean(d.c2_d2_name && d.c2_d2_rpl && String(d.c2_d2_rpl).length === 6 && d.c2_d2_sweater && d.c2_d2_size);
          var c2Doc3Ok = Boolean(d.c2_d3_name && d.c2_d3_rpl && String(d.c2_d3_rpl).length === 6 && d.c2_d3_sweater && d.c2_d3_size);
          var c2Complete = c2Doc1Ok && c2Doc2Ok && c2Doc3Ok;
          var c2HasAny = Boolean(d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater || d.c2_d1_size || d.c2_d2_name || d.c2_d2_rpl || d.c2_d2_sweater || d.c2_d2_size || d.c2_d3_name || d.c2_d3_rpl || d.c2_d3_sweater || d.c2_d3_size);
          var c2Status = c2Complete ? "Complete" : (c2HasAny ? "In Progress" : "Not Started");

          var v2 = [
            d.c2_d1_name||"", d.c2_d1_rpl||"", d.c2_d1_sweater||"", d.c2_d1_size||"",
            d.c2_d2_name||"", d.c2_d2_rpl||"", d.c2_d2_sweater||"", d.c2_d2_size||"",
            d.c2_d3_name||"", d.c2_d3_rpl||"", d.c2_d3_sweater||"", d.c2_d3_size||"",
            c2Status
          ];
          sheet2.getRange(j + 3, 7, 1, 13).setValues([v2]);
          break;
        }
      }
    }
  }
}

function clearRegionSheetData(ss, regCode) {
  if (!regCode) return;
  var pair = getOrCreateSheets(ss);
  var sheet1 = pair.sheet1;
  var sheet2 = pair.sheet2;

  if (sheet1) {
    var lr1 = sheet1.getLastRow();
    if (lr1 > 2) {
      var rCodes1 = sheet1.getRange(3, 2, lr1 - 2, 1).getValues();
      for (var i = 0; i < rCodes1.length; i++) {
        if (String(rCodes1[i][0]).trim() === String(regCode).trim()) {
          var v1 = ["", "", "", "", "", "", "", "", "", "", "Not Started"];
          sheet1.getRange(i + 3, 7, 1, 11).setValues([v1]);
        }
      }
    }
  }

  if (sheet2) {
    var lr2 = sheet2.getLastRow();
    if (lr2 > 2) {
      var rCodes2 = sheet2.getRange(3, 2, lr2 - 2, 1).getValues();
      for (var j = 0; j < rCodes2.length; j++) {
        if (String(rCodes2[j][0]).trim() === String(regCode).trim()) {
          var v2 = ["", "", "", "", "", "", "", "", "", "", "", "", "Not Started"];
          sheet2.getRange(j + 3, 7, 1, 13).setValues([v2]);
        }
      }
    }
  }
}

function clearAllSheetData(ss) {
  var pair = getOrCreateSheets(ss);
  var sheet1 = pair.sheet1;
  var sheet2 = pair.sheet2;

  fixAndRestoreStatusColumn();

  if (sheet1) {
    var lr1 = sheet1.getLastRow();
    if (lr1 > 2) {
      sheet1.getRange(3, 7, lr1 - 2, 11).clearContent();
    }
  }

  if (sheet2) {
    var lr2 = sheet2.getLastRow();
    if (lr2 > 2) {
      sheet2.getRange(3, 7, lr2 - 2, 13).clearContent();
    }
  }
}
