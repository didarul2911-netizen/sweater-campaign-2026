/**
 * =========================================================================
 * EXIUM MUPS - SWEATER CAMPAIGN 2026 (4Q'26)
 * GOOGLE APPS SCRIPT BACKEND (FOR GOOGLE DRIVE / GOOGLE SHEETS)
 * =========================================================================
 * 
 * Instructions to deploy on Google Drive:
 * 1. Open Google Drive (https://drive.google.com).
 * 2. Create a new Google Sheet named "Exium MUPS Sweater Campaign 2026".
 * 3. In the top menu, go to: Extensions -> Apps Script.
 * 4. Delete any code in Code.gs and paste ALL the code below.
 * 5. Click "Save" (Ctrl+S) and then click "Run" -> select "setupSheets" to create initial headers.
 * 6. Click "Deploy" (top right) -> "New deployment".
 * 7. Select type: "Web app".
 * 8. Configuration:
 *    - Description: "Exium Sweater Sync API"
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone" (allows field force browsers to send data)
 * 9. Click "Deploy", authorize permissions, and COPY the "Web App URL" (ends with /exec).
 * 10. Open the Sweater Portal -> Click "Admin" -> "Cloud Setup" -> Paste the URL!
 * =========================================================================
 */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sheet 1: Gyne Core Doctor (Family)
  var sheet1 = ss.getSheetByName("Gyne Core Doctor (Family)");
  if (!sheet1) {
    sheet1 = ss.insertSheet("Gyne Core Doctor (Family)");
  }
  var h1 = [
    "Zone", "SAP Region Code", "Region", "Regional Head", 
    "SAP Territory Code", "Territory", "Doctor Name", "Doctor RPL ID", 
    "Sweater 1", "Size 1", "Sweater 2", "Size 2", 
    "Sweater 3", "Size 3", "Sweater 4", "Size 4", "Status", "Last Updated"
  ];
  sheet1.getRange(1, 1, 1, h1.length).setValues([h1]).setFontWeight("bold").setBackground("#0f766e").setFontColor("#ffffff");
  sheet1.setFrozenRows(1);

  // 2. Sheet 2: Core Doctor Maximization
  var sheet2 = ss.getSheetByName("Core Doctor Maximization");
  if (!sheet2) {
    sheet2 = ss.insertSheet("Core Doctor Maximization");
  }
  var h2 = [
    "Zone", "SAP Region Code", "Region", "Regional Head", 
    "SAP Territory Code", "Territory", 
    "Doctor 1 Name", "Doctor 1 RPL ID", "Sweater 1", "Size 1", 
    "Doctor 2 Name", "Doctor 2 RPL ID", "Sweater 2", "Size 2", 
    "Doctor 3 Name", "Doctor 3 RPL ID", "Sweater 3", "Size 3", 
    "Doctor 4 Name", "Doctor 4 RPL ID", "Sweater 4", "Size 4", "Status", "Last Updated"
  ];
  sheet2.getRange(1, 1, 1, h2.length).setValues([h2]).setFontWeight("bold").setBackground("#581c87").setFontColor("#ffffff");
  sheet2.setFrozenRows(1);
  
  // 3. Sheet 3: Raw Store KV (for instant portal JSON sync)
  var sheet3 = ss.getSheetByName("Portal_Store_DB");
  if (!sheet3) {
    sheet3 = ss.insertSheet("Portal_Store_DB");
  }
  sheet3.getRange(1, 1, 1, 3).setValues([["SAP Territory Code", "Data JSON", "Last Updated"]]).setFontWeight("bold");
  sheet3.setFrozenRows(1);
  sheet3.hideSheet();
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dbSheet = ss.getSheetByName("Portal_Store_DB");
    if (!dbSheet) {
      setupSheets();
      dbSheet = ss.getSheetByName("Portal_Store_DB");
    }

    var store = {};
    var lastRow = dbSheet.getLastRow();
    if (lastRow > 1) {
      var values = dbSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        var code = String(values[i][0]);
        var jsonStr = values[i][1];
        if (code && jsonStr) {
          try {
            store[code] = JSON.parse(jsonStr);
          } catch(err) {}
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", store: store, total_territories: Object.keys(store).length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action || "save_territory";
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "save_territory") {
      var terrCode = String(postData.sap_territory_code);
      var terrData = postData.data;
      var meta = postData.meta || {}; // zone, region, rh, territory name

      saveSingleTerritoryToSheets(ss, terrCode, terrData, meta);

      return ContentService
        .createTextOutput(JSON.stringify({ status: "success", message: "Territory saved to Google Sheets successfully", terrCode: terrCode }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "save_batch") {
      var batch = postData.batch || {};
      for (var tCode in batch) {
        saveSingleTerritoryToSheets(ss, String(tCode), batch[tCode].data, batch[tCode].meta || {});
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: "success", message: "Batch saved successfully" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveSingleTerritoryToSheets(ss, terrCode, d, meta) {
  var dbSheet = ss.getSheetByName("Portal_Store_DB");
  if (!dbSheet) {
    setupSheets();
    dbSheet = ss.getSheetByName("Portal_Store_DB");
  }

  var nowStr = Utilities.formatDate(new Date(), "Asia/Dhaka", "yyyy-MM-dd HH:mm:ss");

  // 1. Update Portal_Store_DB (Fast KV Store)
  var lastRow = dbSheet.getLastRow();
  var foundRow = 0;
  if (lastRow > 1) {
    var codes = dbSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < codes.length; i++) {
      if (String(codes[i][0]) === terrCode) {
        foundRow = i + 2;
        break;
      }
    }
  }

  var jsonStr = JSON.stringify(d);
  if (foundRow > 0) {
    dbSheet.getRange(foundRow, 2, 1, 2).setValues([[jsonStr, nowStr]]);
  } else {
    dbSheet.appendRow([terrCode, jsonStr, nowStr]);
  }

  // 2. Update Human-Readable Sheets (Sheet 1 & Sheet 2)
  updateSheet1Row(ss, terrCode, d, meta, nowStr);
  updateSheet2Row(ss, terrCode, d, meta, nowStr);
}

function updateSheet1Row(ss, terrCode, d, meta, nowStr) {
  var sheet = ss.getSheetByName("Gyne Core Doctor (Family)");
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var rowIdx = 0;
  if (lastRow > 1) {
    var codes = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    for (var i = 0; i < codes.length; i++) {
      if (String(codes[i][0]) === terrCode) {
        rowIdx = i + 2;
        break;
      }
    }
  }

  var isComplete = (d.c1_doc_name && d.c1_doc_rpl && d.c1_doc_rpl.length === 6 && d.c1_m1_sweater && d.c1_m1_size && d.c1_m2_sweater && d.c1_m2_size && d.c1_m3_sweater && d.c1_m3_size && d.c1_m4_sweater && d.c1_m4_size);
  var status = isComplete ? "Complete" : ((d.c1_doc_name || d.c1_doc_rpl || d.c1_m1_sweater) ? "In Progress" : "Not Started");

  var rowData = [
    meta.zone || "", meta.sap_region_code || "", meta.region_name || "", meta.regional_head || "",
    terrCode, meta.territory_name || "",
    d.c1_doc_name || "", d.c1_doc_rpl || "",
    d.c1_m1_sweater || "", d.c1_m1_size || "",
    d.c1_m2_sweater || "", d.c1_m2_size || "",
    d.c1_m3_sweater || "", d.c1_m3_size || "",
    d.c1_m4_sweater || "", d.c1_m4_size || "",
    status, nowStr
  ];

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

function updateSheet2Row(ss, terrCode, d, meta, nowStr) {
  var sheet = ss.getSheetByName("Core Doctor Maximization");
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var rowIdx = 0;
  if (lastRow > 1) {
    var codes = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
    for (var i = 0; i < codes.length; i++) {
      if (String(codes[i][0]) === terrCode) {
        rowIdx = i + 2;
        break;
      }
    }
  }

  var isComplete = (d.c2_d1_name && d.c2_d1_rpl && d.c2_d1_rpl.length === 6 && d.c2_d1_sweater && d.c2_d1_size && 
                    d.c2_d2_name && d.c2_d2_rpl && d.c2_d2_rpl.length === 6 && d.c2_d2_sweater && d.c2_d2_size && 
                    d.c2_d3_name && d.c2_d3_rpl && d.c2_d3_rpl.length === 6 && d.c2_d3_sweater && d.c2_d3_size && 
                    d.c2_d4_name && d.c2_d4_rpl && d.c2_d4_rpl.length === 6 && d.c2_d4_sweater && d.c2_d4_size);
  var status = isComplete ? "Complete" : ((d.c2_d1_name || d.c2_d1_rpl || d.c2_d1_sweater) ? "In Progress" : "Not Started");

  var rowData = [
    meta.zone || "", meta.sap_region_code || "", meta.region_name || "", meta.regional_head || "",
    terrCode, meta.territory_name || "",
    d.c2_d1_name || "", d.c2_d1_rpl || "", d.c2_d1_sweater || "", d.c2_d1_size || "",
    d.c2_d2_name || "", d.c2_d2_rpl || "", d.c2_d2_sweater || "", d.c2_d2_size || "",
    d.c2_d3_name || "", d.c2_d3_rpl || "", d.c2_d3_sweater || "", d.c2_d3_size || "",
    d.c2_d4_name || "", d.c2_d4_rpl || "", d.c2_d4_sweater || "", d.c2_d4_size || "",
    status, nowStr
  ];

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}
