var SPREADSHEET_ID = "1ZOrwkS-_8-ngI3tvLqJBo3VOOg7Hp3csemw3boA15os"; 
var SHEET_NAME = "Sheet1"; 

// --- ENVIRONMENT CONFIGURATIONS ---
var ENV_CONFIG = {
  "STAGING": {
    SPREADSHEET_ID: "1ZOrwkS-_8-ngI3tvLqJBo3VOOg7Hp3csemw3boA15os",
    SHEET_NAME: "Sheet_Staging",
    DRIVE_FOLDER: "POS_Activations_Staging",
    SUPABASE_URL: "https://pcawzzubhylrplrrgpfv.supabase.co",
    SUPABASE_KEY: "YOUR_SUPABASE_SECRET_KEY", 
    RAZORPAY_KEY_ID: "YOUR_TEST_KEY_ID",
    RAZORPAY_KEY_SECRET: "YOUR_TEST_KEY_SECRET"
  },
  "LIVE": {
    SPREADSHEET_ID: "1ZOrwkS-_8-ngI3tvLqJBo3VOOg7Hp3csemw3boA15os",
    SHEET_NAME: "Sheet1",
    DRIVE_FOLDER: "POS_Activations",
    SUPABASE_URL: "https://xwlzckxsspgyvgszueht.supabase.co",
    SUPABASE_KEY: "sb_publishable_USFK0tj1qga8bDoNxdx06w_-uXLDQ5K",
    RAZORPAY_KEY_ID: "YOUR_KEY_ID",
    RAZORPAY_KEY_SECRET: "YOUR_KEY_SECRET"
  }
};

function getConfig(env) {
  var requestedEnv = (env && String(env).toUpperCase() === "STAGING") ? "STAGING" : "LIVE";
  return ENV_CONFIG[requestedEnv];
}

function getInvoiceVisibility() {
  var props = PropertiesService.getScriptProperties();
  var val = props.getProperty('SHOW_INVOICE');
  return val !== 'false'; 
}

function doGet(e) {
  var currentEnv = (e && e.parameter && e.parameter.env) ? String(e.parameter.env).toUpperCase() : "LIVE";
  var currentAppMode = (e && e.parameter && e.parameter.app) ? String(e.parameter.app).toLowerCase() : "desktop";

  if (e && e.parameter && e.parameter.page === "policies") {
    return HtmlService.createTemplateFromFile('Policies').evaluate()
        .setTitle('Store Policies & Legal Terms')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (e && e.parameter && e.parameter.mode === "admin") {
    var adminTemplate = HtmlService.createTemplateFromFile('Admin');
    adminTemplate.showInvoice = getInvoiceVisibility();
    adminTemplate.env = currentEnv;
    return adminTemplate.evaluate()
        .setTitle('Admin Dashboard - License Manager')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var indexTemplate = HtmlService.createTemplateFromFile('Index');
  indexTemplate.showInvoice = getInvoiceVisibility();
  indexTemplate.env = currentEnv;
  indexTemplate.appMode = currentAppMode;
  
  return indexTemplate.evaluate()
      .setTitle('Hotel POS License Generator')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- VISITOR & LOCATION ANALYTICS BACKEND ---
function recordWebsiteVisit(clientLocation, env) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const tz = "Asia/Kolkata";
    const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    const props = PropertiesService.getScriptProperties();

    let dailyMap = {};
    const rawMap = props.getProperty("VISITOR_DAILY_MAP");
    if (rawMap) {
      try { dailyMap = JSON.parse(rawMap); } catch (e) { dailyMap = {}; }
    }

    dailyMap[todayStr] = (Number(dailyMap[todayStr]) || 0) + 1;
    props.setProperty("VISITOR_DAILY_MAP", JSON.stringify(dailyMap));

    var config = getConfig(env);
    syncVisitToSheet_(config.SPREADSHEET_ID, todayStr, dailyMap[todayStr], clientLocation || "Unknown Location", timestamp);

    return buildVisitorStatsResponse_(dailyMap, todayStr);
  } catch (err) {
    console.error("Visitor logging error: " + err);
    return getAdminVisitorStats();
  } finally {
    lock.releaseLock();
  }
}

function getAdminVisitorStats() {
  const tz = "Asia/Kolkata";
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const props = PropertiesService.getScriptProperties();
  let dailyMap = {};

  const rawMap = props.getProperty("VISITOR_DAILY_MAP");
  if (rawMap) {
    try { dailyMap = JSON.parse(rawMap); } catch (e) { dailyMap = {}; }
  }
  return buildVisitorStatsResponse_(dailyMap, todayStr);
}

function buildVisitorStatsResponse_(dailyMap, todayStr) {
  let totalVisitors = 0;
  Object.keys(dailyMap).forEach(function(dateKey) {
    totalVisitors += Number(dailyMap[dateKey]) || 0;
  });

  return {
    totalVisitors: totalVisitors,
    todayVisitors: Number(dailyMap[todayStr]) || 0,
    todayDate: todayStr,
    dailyCounts: dailyMap
  };
}

function syncVisitToSheet_(sheetId, dateStr, count, location, timestamp) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    let sheet = ss.getSheetByName("Visitors");
    if (!sheet) {
      sheet = ss.insertSheet("Visitors");
      sheet.appendRow(["Date", "Visitor Count", "Visitor Location", "Timestamp"]);
    }
    sheet.appendRow([dateStr, count, location, timestamp]);
  } catch (e) {
    console.warn("Visitor Sheet sync skipped: " + e.message);
  }
}

function checkRequestStatus(requestId, env) {
  var config = getConfig(env);
  var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) { 
      return {
        status: String(data[i][3]), 
        key: String(data[i][4])     
      };
    }
  }
  return { status: "NOT_FOUND" };
}

function getAllRequests(env) {
  var config = getConfig(env);
  var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var requestList = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) { 
      var timeData = data[i][6];
      var formattedTime = "N/A";
      if (timeData) {
        if (timeData instanceof Date) {
          formattedTime = Utilities.formatDate(timeData, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else {
          formattedTime = String(timeData);
        }
      }

      requestList.push({
        rowIndex: i + 1, 
        requestId: String(data[i][0]),
        systemId: String(data[i][1]),
        duration: String(data[i][2]),
        status: String(data[i][3]), 
        key: String(data[i][4] || ""),    
        location: String(data[i][5] || "Unknown"),
        timestamp: formattedTime,
        photoUrl: String(data[i][7] || "No Photo"),
        shopName: String(data[i][8] || "N/A"),
        shopAddress: String(data[i][9] || "N/A"),
        machineType: String(data[i][10] || "N/A"),
        areaPin: String(data[i][11] || "N/A"),
        paymentMethod: String(data[i][12] || "N/A"),
        paymentAmount: String(data[i][13] || "0"),
        invoiceType: String(data[i][14] || "N/A")
      });
    }
  }
  return requestList.reverse(); 
}

function approveFromAdminPanel(rowIndex, clientIdentifier, durationCode, machineType, env) {
  var config = getConfig(env);
  var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  var approvedKey = generateLicenseKeyLogic(clientIdentifier, durationCode, machineType);
  sheet.getRange(rowIndex, 4).setValue("APPROVED");
  sheet.getRange(rowIndex, 5).setValue(approvedKey);
  return approvedKey;
}

function adminGenerateKey(clientIdentifier, durationChoice, locationStr, machineType, appType, env) {
  var config = getConfig(env);
  var durationCodes = { '1': 'DEMO', '2': '3MON', '3': '6MON', '4': '12MON', '5': 'LIFE' };
  var durationCode = durationCodes[durationChoice] || '12MON';
  var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  var requestId = "REQ-" + Utilities.getUuid().substring(0, 8).toUpperCase();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var generatedKey = generateLicenseKeyLogic(clientIdentifier, durationCode, machineType);
  var loc = locationStr ? locationStr : "Admin Generated";
  var finalAppType = appType ? appType : "Standalone Application (Desktop)";
  sheet.appendRow([requestId, String(clientIdentifier || "").trim(), durationCode, "APPROVED", generatedKey, loc, timestamp, "N/A", "N/A", "N/A", machineType, "N/A", "Cash", "0", "Non-GST", finalAppType]);
  return generatedKey;
}

function deleteMultipleRequests(requestIds, env) {
  var config = getConfig(env);
  var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  for (var i = data.length - 1; i > 0; i--) {
    if (requestIds.indexOf(data[i][0]) !== -1) {
      sheet.deleteRow(i + 1);
    }
  }
  return true;
}

function requestKeyFromPython(systemId, durationChoice, locationStr, photoBase64, shopName, shopAddress, machineType, areaPin, paymentMethod, paymentAmount, invoiceType, env, appType) {
  var config = getConfig(env);
  var durationCodes = { '1': 'DEMO', '2': '3MON', '3': '6MON', '4': '12MON', '5': 'LIFE' };
  var durationCode = durationCodes[durationChoice] || '12MON';
  
  var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  var requestId = "REQ-" + Utilities.getUuid().substring(0, 8).toUpperCase();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var finalAppType = appType ? appType : "Standalone Application (Desktop)";
  
  var savedPhotoUrl = "No Photo";
  if (photoBase64) {
    try {
      var base64Data = photoBase64.split(',')[1];
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', requestId + '_photo.png');
      var folders = DriveApp.getFoldersByName(config.DRIVE_FOLDER);
      var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(config.DRIVE_FOLDER);
      var file = targetFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      savedPhotoUrl = file.getUrl();
    } catch (e) {
      savedPhotoUrl = "Error Saving Photo";
    }
  }
  
  if (durationCode === 'DEMO') {
    var demoKey = generateLicenseKeyLogic(systemId, durationCode, machineType);
    sheet.appendRow([requestId, systemId.trim(), durationCode, "APPROVED", demoKey, locationStr, timestamp, savedPhotoUrl, shopName, shopAddress, machineType, areaPin, paymentMethod, paymentAmount, invoiceType, finalAppType]);
    return { status: "APPROVED", key: demoKey };
  } else {
    sheet.appendRow([requestId, systemId.trim(), durationCode, "PENDING", "", locationStr, timestamp, savedPhotoUrl, shopName, shopAddress, machineType, areaPin, paymentMethod, paymentAmount, invoiceType, finalAppType]);
    
    if (paymentMethod === 'Online') {
      var orderResponse = createPaymentOrder(durationChoice, systemId.trim(), requestId, shopName, machineType, areaPin, paymentAmount, env);
      return { 
        status: "PAYMENT_REQUIRED", 
        requestId: requestId, 
        orderId: orderResponse.id || "",
        amount: orderResponse.amount || (paymentAmount * 100), 
        message: "Please complete your payment to automatically generate the key. Request ID: " + requestId 
      };
    } else {
      return { 
        status: "PENDING", 
        requestId: requestId, 
        message: "Cash payment requested. Please pay the admin directly. Your key will be generated upon admin approval. Request ID: " + requestId 
      };
    }
  }
}

function createPaymentOrder(durationChoice, systemId, requestId, shopName, machineType, areaPin, paymentAmount, env) {
  var config = getConfig(env);
  var amount = parseInt(paymentAmount, 10) * 100; 
  if (isNaN(amount) || amount <= 0) amount = 500000; 
  
  var url = "https://api.razorpay.com/v1/orders";
  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(config.RAZORPAY_KEY_ID + ":" + config.RAZORPAY_KEY_SECRET)
    },
    "payload": {
      "amount": amount,
      "currency": "INR",
      "receipt": requestId,
      "notes": {
        "system_id": systemId,
        "duration_choice": durationChoice,
        "request_id": requestId,
        "shop_name": shopName,
        "machine_type": machineType,
        "area_pin": areaPin
      }
    },
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    
    if (payload.event === "payment.captured") {
      var paymentData = payload.payload.payment.entity;
      var customerEmail = paymentData.email;
      var systemId = paymentData.notes.system_id;
      var durationChoice = paymentData.notes.duration_choice;
      var reqId = paymentData.notes.request_id;
      var shopName = paymentData.notes.shop_name;
      
      var durationCodes = { '1': 'DEMO', '2': '3MON', '3': '6MON', '4': '12MON', '5': 'LIFE' };
      var durationCode = durationCodes[durationChoice] || '12MON';
      
      var config = getConfig("LIVE");
      var generatedKey = generateLicenseKeyLogic(systemId, durationCode, "Desktop");
      
      var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] === reqId) {
          sheet.getRange(i + 1, 4).setValue("APPROVED"); 
          sheet.getRange(i + 1, 5).setValue(generatedKey); 
          break;
        }
      }
      
      MailApp.sendEmail({
        to: customerEmail,
        subject: "Your Hotel POS License Key",
        body: "Payment Successful for " + shopName + "!\n\nSystem ID: " + systemId + "\nDuration: " + durationCode + "\n\nYour License Key: " + generatedKey + "\n\nPlease enter this key into your POS application."
      });
      
      return ContentService.createTextOutput("Success");
    }
    return ContentService.createTextOutput("Ignored");
  } catch (error) {
    return ContentService.createTextOutput("Error");
  }
}

function generateLicenseKeyLogic(clientIdentifier, durationCode, machineType) {
  var SECRET_WORD = "MySuperSecretHotelApp2026";
  var today = new Date();
  var generationDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
  var rawData = String(clientIdentifier || "").trim().toUpperCase() + SECRET_WORD + String(durationCode || "").trim().toUpperCase() + String(machineType || "").trim().toUpperCase() + generationDate;
  
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawData, Utilities.Charset.UTF_8);
  var hexString = "";
  for (var i = 0; i < rawHash.length; i++) {
    var byteValue = rawHash[i];
    if (byteValue < 0) byteValue += 256;
    var byteString = byteValue.toString(16);
    if (byteString.length == 1) byteString = "0" + byteString;
    hexString += byteString;
  }
  
  var rawKey = hexString.substring(0, 16).toUpperCase();
  return rawKey.substring(0, 4) + "-" + rawKey.substring(4, 8) + "-" + rawKey.substring(8, 12) + "-" + rawKey.substring(12, 16);
}

function generateInvoicePDFBase64(requestId, env) {
  var config = getConfig(env);
  var sheet = SpreadsheetApp.openById(config.SPREADSHEET_ID).getSheetByName(config.SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var reqData = null;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === requestId) { 
      reqData = {
        durationCode: String(data[i][2]),
        key: String(data[i][4] || "PENDING"),
        date: data[i][6] ? Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), "dd-MM-yyyy") : "N/A",
        shopName: String(data[i][8] || "N/A"),
        shopAddress: String(data[i][9] || "N/A"),
        machineType: String(data[i][10] || "N/A"),
        pin: String(data[i][11] || "N/A"),
        paymentMethod: String(data[i][12] || "N/A"),
        amount: parseFloat(data[i][13]) || 0,
        invoiceType: String(data[i][14] || "Non-GST")
      };
      break;
    }
  }
  
  if (!reqData) return null;
  
  var basePrice = reqData.amount;
  var cgst = 0;
  var sgst = 0;
  var gstRows = "";
  
  if (reqData.invoiceType === 'GST') {
    basePrice = reqData.amount / 1.18;
    cgst = basePrice * 0.09;
    sgst = basePrice * 0.09;
    
    gstRows = `
      <tr>
        <td style="padding: 8px 0; font-weight: bold; color: #555;">Base Price:</td>
        <td style="text-align: right; color: #555;">Rs. ${basePrice.toFixed(2)}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; color: #555;">CGST (9%):</td>
        <td style="text-align: right; color: #555;">Rs. ${cgst.toFixed(2)}</td>
      </tr>
      <tr>
        <td style="padding: 8px 0; border-bottom: 2px solid #dee2e6; color: #555;">SGST (9%):</td>
        <td style="text-align: right; border-bottom: 2px solid #dee2e6; color: #555;">Rs. ${sgst.toFixed(2)}</td>
      </tr>
    `;
  }
  
  var html = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 30px; max-width: 800px; margin: auto;">
      <div style="text-align: center; border-bottom: 2px solid #B8860B; padding-bottom: 20px; margin-bottom: 25px;">
        <img src="https://drive.google.com/thumbnail?id=1ojCplKaANjQBes_WLSJNSI2EMi7YAOzN&sz=w800" alt="GVSS Logo" style="max-height: 80px; width: auto; object-fit: contain; margin-bottom: 10px;">
        <h1 style="margin: 0; color: #B8860B; font-size: 26px;">GVSS <span style="font-size: 16px; color: #333; font-weight: normal;">Complete Software Solutions</span></h1>
        <p style="margin: 5px 0 2px 0; color: #555; font-size: 13px;">#10/1, 1st cross E block, RHCS Layout, Nagarabhavi, Bangalore 560091.</p>
        <p style="margin: 2px 0; color: #555; font-size: 13px;">📞 +91 9986021497</p>
        <p style="margin: 2px 0; color: #555; font-size: 13px;">✉️ gvss4594@gmail.com | 🌐 https://tinyurl.com/GVSS45</p>
      </div>
      
      <div style="text-align: center; margin-bottom: 25px;">
        <h2 style="color: #333; margin: 0; font-size: 20px; letter-spacing: 1px;">TAX INVOICE</h2>
      </div>
      
      <table style="width: 100%; margin-bottom: 30px;">
        <tr>
          <td style="vertical-align: top; width: 60%;">
            <h3 style="margin: 0 0 8px 0; color: #B8860B; font-size: 15px;">Billed To:</h3>
            <p style="margin: 3px 0; font-weight: bold; font-size: 16px; color: #333;">${reqData.shopName}</p>
            <p style="margin: 3px 0; font-size: 13px; color: #555; max-width: 80%;">${reqData.shopAddress}</p>
            <p style="margin: 3px 0; font-size: 13px; color: #555;">Area PIN: ${reqData.pin}</p>
          </td>
          <td style="text-align: right; vertical-align: top; width: 40%;">
            <p style="margin: 3px 0; font-size: 13px; color: #555;"><strong>Invoice No:</strong> ${requestId}</p>
            <p style="margin: 3px 0; font-size: 13px; color: #555;"><strong>Date:</strong> ${reqData.date}</p>
            <p style="margin: 3px 0; font-size: 13px; color: #555;"><strong>Payment Method:</strong> ${reqData.paymentMethod}</p>
          </td>
        </tr>
      </table>

      <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px; border: 1px solid #dee2e6;">
        <thead>
          <tr style="background-color: #f8f9fa;">
            <th style="padding: 12px; border: 1px solid #dee2e6; text-align: left; color: #333; font-size: 14px;">Description</th>
            <th style="padding: 12px; border: 1px solid #dee2e6; text-align: center; color: #333; font-size: 14px;">Machine Type</th>
            <th style="padding: 12px; border: 1px solid #dee2e6; text-align: right; color: #333; font-size: 14px;">Amount (INR)</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="padding: 12px; border: 1px solid #dee2e6; font-size: 14px; color: #333;">
              <strong>POS Software License</strong><br>
              <small style="color: #6c757d; display: block; margin-top: 5px;">Key: ${reqData.key}</small>
              <small style="color: #6c757d; display: block;">Validity: ${reqData.durationCode}</small>
            </td>
            <td style="padding: 12px; border: 1px solid #dee2e6; text-align: center; font-size: 14px; color: #333;">${reqData.machineType}</td>
            <td style="padding: 12px; border: 1px solid #dee2e6; text-align: right; font-size: 14px; color: #333;">Rs. ${basePrice.toFixed(2)}</td>
          </tr>
        </tbody>
      </table>

      <div style="width: 100%; display: table;">
        <div style="display: table-row;">
          <div style="display: table-cell; width: 50%;"></div>
          <div style="display: table-cell; width: 50%;">
            <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
              ${gstRows}
              <tr>
                <td style="padding: 12px 0; font-weight: bold; font-size: 16px; color: #B8860B;">Total Amount:</td>
                <td style="text-align: right; font-weight: bold; font-size: 16px; color: #B8860B;">Rs. ${reqData.amount.toFixed(2)}</td>
              </tr>
            </table>
          </div>
        </div>
      </div>
      
      <div style="margin-top: 50px; font-size: 12px; color: #555; background-color: #f8f9fa; padding: 15px; border-radius: 6px; border: 1px solid #eee;">
        <h4 style="margin: 0 0 8px 0; color: #333; font-size: 13px;">Terms & Conditions:</h4>
        <ol style="margin: 0; padding-left: 18px; line-height: 1.6;">
          <li><strong>License Scope:</strong> This license key is non-transferable and valid strictly for the single machine matching the System ID provided at checkout.</li>
          <li><strong>Restrictions:</strong> Reverse engineering, decompiling, or attempting to modify the software is strictly prohibited.</li>
          <li><strong>Refund Policy:</strong> As this is a digital license key, all sales are final once the key is generated and dispatched.</li>
        </ol>
      </div>

      <div style="margin-top: 30px; text-align: center; color: #999; font-size: 11px; padding-top: 15px; border-top: 1px solid #eee;">
        <p>This is a computer-generated invoice and does not require a physical signature.</p>
      </div>
    </div>
  `;
  
  var blob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF).setName(requestId + ".pdf");
  return Utilities.base64Encode(blob.getBytes());
}