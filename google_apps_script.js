/**
 * =================================================================================
 * GOOGLE APPS SCRIPT FOR REAL-TIME CLOUD DATABASE (GOOGLE SHEETS SYNC)
 * =================================================================================
 * Instructions to deploy in 2 minutes:
 * 1. Open your Google Drive -> New -> Google Sheets (or upload 'Sweater_Campaign_2026_Master_Format.xlsx').
 * 2. In Google Sheets, click "Extensions" -> "Apps Script".
 * 3. Delete any default code and paste this entire file content.
 * 4. Click "Deploy" -> "New deployment".
 * 5. Select type: "Web app".
 * 6. Set "Execute as": "Me" and "Who has access": "Anyone".
 * 7. Click "Deploy" and copy the "Web app URL".
 * 8. Paste your Web App URL into the Web Portal settings to sync live data directly to Google Sheets!
 * =================================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    
    if (action === "save_territory") {
      var terrCode = String(postData.sap_territory_code);
      var data = postData.data;
      
      // Update Campaign 1 Sheet: "Gyne Core Doctor (Family)"
      var sheetC1 = ss.getSheetByName("Gyne Core Doctor (Family)");
      if (sheetC1) {
        var valuesC1 = sheetC1.getDataRange().getValues();
        for (var i = 2; i < valuesC1.length; i++) {
          if (String(valuesC1[i][4]) === terrCode) { // Column E is SAP Territory Code (0-indexed 4)
            var r = i + 1;
            sheetC1.getRange(r, 7).setValue(data.c1_doc_name || ""); // G: Doc Name
            sheetC1.getRange(r, 8).setValue(data.c1_m1_sweater || "");
            sheetC1.getRange(r, 9).setValue(data.c1_m1_size || "");
            sheetC1.getRange(r, 10).setValue(data.c1_m2_sweater || "");
            sheetC1.getRange(r, 11).setValue(data.c1_m2_size || "");
            sheetC1.getRange(r, 12).setValue(data.c1_m3_sweater || "");
            sheetC1.getRange(r, 13).setValue(data.c1_m3_size || "");
            sheetC1.getRange(r, 14).setValue(data.c1_m4_sweater || "");
            sheetC1.getRange(r, 15).setValue(data.c1_m4_size || "");
            break;
          }
        }
      }
      
      // Update Campaign 2 Sheet: "Core Doctor Maximization"
      var sheetC2 = ss.getSheetByName("Core Doctor Maximization");
      if (sheetC2) {
        var valuesC2 = sheetC2.getDataRange().getValues();
        for (var j = 2; j < valuesC2.length; j++) {
          if (String(valuesC2[j][4]) === terrCode) {
            var r2 = j + 1;
            sheetC2.getRange(r2, 7).setValue(data.c2_d1_name || "");
            sheetC2.getRange(r2, 8).setValue(data.c2_d1_sweater || "");
            sheetC2.getRange(r2, 9).setValue(data.c2_d1_size || "");
            sheetC2.getRange(r2, 10).setValue(data.c2_d2_name || "");
            sheetC2.getRange(r2, 11).setValue(data.c2_d2_sweater || "");
            sheetC2.getRange(r2, 12).setValue(data.c2_d2_size || "");
            sheetC2.getRange(r2, 13).setValue(data.c2_d3_name || "");
            sheetC2.getRange(r2, 14).setValue(data.c2_d3_sweater || "");
            sheetC2.getRange(r2, 15).setValue(data.c2_d3_size || "");
            sheetC2.getRange(r2, 16).setValue(data.c2_d4_name || "");
            sheetC2.getRange(r2, 17).setValue(data.c2_d4_sweater || "");
            sheetC2.getRange(r2, 18).setValue(data.c2_d4_size || "");
            break;
          }
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Territory " + terrCode + " saved successfully in Google Sheets"
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Unknown action"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    message: "Exium Sweater Campaign Google Cloud API is running!"
  })).setMimeType(ContentService.MimeType.JSON);
}
