function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CN Reminder')
    .addItem('Setup Dropdowns', 'setupDropdowns')
    .addItem('Setup CN Tool', 'setupCNTool')
    .addItem('Log Sheet Names', 'logSheetNames')
    .addToUi();
}

function setupDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet name in setupDropdowns: " + ss.getName());
  const monitoringSheet = ss.getSheetByName("Monitoring JKT");
  const contactSheet = ss.getSheetByName("Import Contact List");

  if (!monitoringSheet || !contactSheet) {
    const errorMsg = "Error: Missing sheets in setupDropdowns. Expected 'Monitoring JKT' and 'Import Contact List'.";
    SpreadsheetApp.getUi().alert(errorMsg);
    Logger.log(errorMsg);
    return;
  }

  const contactData = contactSheet.getRange("B2:D" + contactSheet.getLastRow()).getValues();
  const regions = [...new Set(contactData.map(row => row[0]))].filter(Boolean);

  const regionRange = monitoringSheet.getRange("BE3:BE" + monitoringSheet.getLastRow());
  regionRange.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(regions, true)
    .build());

  monitoringSheet.getRange("BF3:BF" + monitoringSheet.getLastRow()).clearDataValidations();
  monitoringSheet.getRange("BG3:BG" + monitoringSheet.getLastRow()).clearDataValidations();

  applyColorCoding(monitoringSheet, contactData);

  monitoringSheet.getRange("BE1").setValue("Region");
  monitoringSheet.getRange("BF1").setValue("Country");
  monitoringSheet.getRange("BG1").setValue("Branch");
  monitoringSheet.getRange("BH1").setValue("Send Status");
}

function setupCNTool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet name in setupCNTool: " + ss.getName());
  let cnToolSheet = ss.getSheetByName("CN Tool");
  const contactSheet = ss.getSheetByName("Import Contact List");

  if (!contactSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Import Contact List' sheet missing.");
    return;
  }

  if (!cnToolSheet) {
    cnToolSheet = ss.insertSheet("CN Tool");
  }

  // India CN Setup (Column E)
  cnToolSheet.getRange("E2").setValue("Prefix");
  cnToolSheet.getRange("E3").setValue("IND");
  cnToolSheet.getRange("E4").setValue("MUM");
  cnToolSheet.getRange("E5").setValue("BL Number");
  cnToolSheet.getRange("E6").setValue("Branch");

  const contactData = contactSheet.getRange("B2:D" + contactSheet.getLastRow()).getValues();
  const indiaBranches = contactData
    .filter(row => row[0] === "SAS" && row[1] === "India")
    .map(row => row[2])
    .filter(Boolean);
  cnToolSheet.getRange("E6").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(indiaBranches, true)
    .build());

  // EU CN Setup (Column B)
  cnToolSheet.getRange("B2").setValue("Service Code");
  cnToolSheet.getRange("B3").setValue("AD1");
  cnToolSheet.getRange("B4").setValue("TOOT0019W");
  cnToolSheet.getRange("B5").setValue("EGDAM");
  cnToolSheet.getRange("B6").setValue("BL Number");
  cnToolSheet.getRange("B7").setValue("BL Number Input");
  cnToolSheet.getRange("B8").setValue("EU Country");
  cnToolSheet.getRange("B9").setValue("Branch");

  const euCountries = [
    "Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czech Republic", "Denmark",
    "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Ireland", "Italy",
    "Latvia", "Lithuania", "Luxembourg", "Malta", "Netherlands", "Poland", "Portugal",
    "Romania", "Slovakia", "Slovenia", "Spain", "Sweden"
  ];
  cnToolSheet.getRange("B8").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(euCountries, true)
    .build());

  const euBranches = contactData
    .filter(row => row[0] === "EUA" && euCountries.includes(row[1]))
    .map(row => row[2])
    .filter(Boolean);
  cnToolSheet.getRange("B9").setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(euBranches.length > 0 ? euBranches : ["No branches available"], true)
    .build());

  SpreadsheetApp.getUi().alert("Please add two red box drawings in 'CN Tool': 'SEND INDIA CN' (assign to 'sendIndiaCN') and 'SEND EU CN' (assign to 'sendEUCN').");
}

function logSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Spreadsheet name: " + ss.getName());
  const sheets = ss.getSheets();
  const sheetNames = sheets.map(sheet => sheet.getName());
  Logger.log("Sheet names in this spreadsheet: " + sheetNames.join(", "));
  SpreadsheetApp.getUi().alert("Sheet names logged. Check View > Logs in Apps Script editor.");
}

function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.source.getActiveSheet();
  Logger.log("onEdit - Spreadsheet: " + ss.getName() + ", Active sheet: " + sheet.getName());

  if (sheet.getName() === "Monitoring JKT") {
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    if (row < 3) return;

    if (col === 57 || col === 58) { // BE or BF
      updateDependentDropdowns(row);
    }

    if (col === 59 && range.getValue() !== "") { // BG
      Logger.log("Branch selected at row " + row + ": " + range.getValue());
      sendCorrectionNotice(row);
    }
  } else if (sheet.getName() === "CN Tool") {
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    if (row === 8 && col === 2) { // B8 (EU Country)
      onEditCNTool();
    }
  }
}

function onEditCNTool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cnToolSheet = ss.getSheetByName("CN Tool");
  const contactSheet = ss.getSheetByName("Import Contact List");

  if (!cnToolSheet || !contactSheet) return;

  const selectedCountry = cnToolSheet.getRange("B8").getValue();
  const contactData = contactSheet.getRange("B2:D" + contactSheet.getLastRow()).getValues();
  const euBranches = contactData
    .filter(row => row[0] === "EUA" && row[1] === selectedCountry)
    .map(row => row[2])
    .filter(Boolean);

  const euBranchCell = cnToolSheet.getRange("B9");
  euBranchCell.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(euBranches.length > 0 ? euBranches : ["No branches available"], true)
    .build());
}

function updateDependentDropdowns(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monitoringSheet = ss.getSheetByName("Monitoring JKT");
  const contactSheet = ss.getSheetByName("Import Contact List");

  if (!monitoringSheet || !contactSheet) return;

  const region = monitoringSheet.getRange(row, 57).getValue(); // BE
  const country = monitoringSheet.getRange(row, 58).getValue(); // BF
  const contactData = contactSheet.getRange("B2:D" + contactSheet.getLastRow()).getValues();

  if (region) {
    const countries = [...new Set(contactData
      .filter(row => row[0] === region)
      .map(row => row[1]))].filter(Boolean);
    const countryCell = monitoringSheet.getRange(row, 58); // BF
    countryCell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(countries, true)
      .build());
    monitoringSheet.getRange(row, 59).clearDataValidations(); // Clear BG
  }

  if (country) {
    const branches = contactData
      .filter(row => row[0] === region && row[1] === country)
      .map(row => row[2])
      .filter(Boolean);
    const branchCell = monitoringSheet.getRange(row, 59); // BG
    branchCell.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(branches, true)
      .build());
  }

  applyColorCoding(monitoringSheet, contactData);
}

function applyColorCoding(sheet, contactData) {
  const countryRange = sheet.getRange("BF3:BF" + sheet.getLastRow());
  const branchRange = sheet.getRange("BG3:BG" + sheet.getLastRow());
  const countryValues = countryRange.getValues();

  for (let i = 0; i < countryValues.length; i++) {
    const country = countryValues[i][0];
    if (country === "USA") {
      countryRange.getCell(i + 1, 1).setBackground("#0000FF");
      branchRange.getCell(i + 1, 1).setBackground("#ADD8E6");
    } else if (country === "Hong Kong") {
      countryRange.getCell(i + 1, 1).setBackground("#FF0000");
      branchRange.getCell(i + 1, 1).setBackground("#FFCCCC");
    } else if (country === "India") {
      countryRange.getCell(i + 1, 1).setBackground("#FF9933");
      branchRange.getCell(i + 1, 1).setBackground("#FFD700");
    } else if (country) {
      countryRange.getCell(i + 1, 1).setBackground("#D3D3D3");
      branchRange.getCell(i + 1, 1).setBackground("#F0F0F0");
    }
  }
}

function sendIndiaCN() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("sendIndiaCN - Spreadsheet: " + ss.getName());
  const cnToolSheet = ss.getSheetByName("CN Tool");
  const monitoringSheet = ss.getSheetByName("Monitoring JKT");
  const contactSheet = ss.getSheetByName("Import Contact List");

  Logger.log("CN Tool exists: " + (cnToolSheet ? "Yes" : "No"));
  Logger.log("Monitoring JKT exists: " + (monitoringSheet ? "Yes" : "No"));
  Logger.log("Import Contact List exists: " + (contactSheet ? "Yes" : "No"));

  if (!cnToolSheet || !monitoringSheet || !contactSheet) {
    const missingSheets = [];
    if (!cnToolSheet) missingSheets.push("CN Tool");
    if (!monitoringSheet) missingSheets.push("Monitoring JKT");
    if (!contactSheet) missingSheets.push("Import Contact List");
    const errorMsg = "Error: Missing required sheets - " + missingSheets.join(", ");
    SpreadsheetApp.getUi().alert(errorMsg);
    Logger.log(errorMsg);
    return;
  }

  try {
    const prefix = cnToolSheet.getRange("E3").getValue();
    const subPrefix = cnToolSheet.getRange("E4").getValue();
    const blNumberInput = cnToolSheet.getRange("E5").getValue();
    const branch = cnToolSheet.getRange("E6").getValue();

    if (!blNumberInput || !branch) {
      SpreadsheetApp.getUi().alert("Error: Please enter a BL Number in E5 and select a Branch in E6.");
      return;
    }

    const monitoringData = monitoringSheet.getRange("C3:V" + monitoringSheet.getLastRow()).getValues();
    let matchingRowData = null;
    let loiUrl = null;
    for (let i = 0; i < monitoringData.length; i++) {
      if (monitoringData[i][0] === blNumberInput) {
        matchingRowData = monitoringData[i].slice(3, 19); // F:U
        loiUrl = monitoringData[i][19]; // V
        break;
      }
    }

    if (!matchingRowData) {
      throw new Error("No matching BL Number found in Monitoring JKT.");
    }

    const contactData = contactSheet.getRange("B2:F" + contactSheet.getLastRow()).getValues();
    let email = "";
    for (let i = 0; i < contactData.length; i++) {
      if (contactData[i][0] === "SAS" && contactData[i][1] === "India" && contactData[i][2] === branch) {
        email = contactData[i][4];
        break;
      }
    }

    if (!email) {
      throw new Error("No email found for India Branch: " + branch);
    }

    const tableHtml = generateCorrectionTable(matchingRowData);
    const subject = `${prefix}//${subPrefix} CORRECTION NOTICE AFTER BDR FOR BL ${blNumberInput}`;
    const body = `
      Dear Colleagues,<br><br>
      We have updated the data in OPUS for BL ${blNumberInput}. Please find the correction details below:<br><br>
      ${tableHtml}<br><br>
      Thank you and best regards,<br>
      <b>□-----------------------------------□</b><br>
      [Your Name]<br>
      Marketing & Commercial<br>
      Sales Management: CS Desk<br>
      Export Documentation<br>
      <b>□-----------------------------------□</b><br>
      <span style="color: magenta;"><b>PT. OCEAN NETWORK EXPRESS INDONESIA</b></span><br>
      AIA Central, 22nd Floor, Jl. Jenderal Sudirman Kav. 48 A<br>
      Jakarta 12930<br>
      Phone: +62 21 50815150<br>
      DID: +62 21 50889611<br>
      www.one-line.com
    `;

    let emailOptions = {
      to: email,
      subject: subject,
      htmlBody: body
    };

    if (loiUrl) {
      try {
        const fileId = extractFileIdFromUrl(loiUrl);
        const loiFile = DriveApp.getFileById(fileId);
        emailOptions.attachments = [loiFile.getBlob()];
        Logger.log("Attaching LOI file: " + loiFile.getName());
      } catch (err) {
        Logger.log("Failed to attach LOI file: " + err.message);
      }
    }

    Logger.log(`Sending India CN email to: ${email}, Subject: ${subject}`);
    MailApp.sendEmail(emailOptions);

    SpreadsheetApp.getUi().alert("India CN sent successfully to " + email);
  } catch (error) {
    SpreadsheetApp.getUi().alert("Failed to send India CN: " + error.message);
    Logger.log("Error in sendIndiaCN: " + error.message);
  }
}

function sendEUCN() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("sendEUCN - Spreadsheet: " + ss.getName());
  const cnToolSheet = ss.getSheetByName("CN Tool");
  const monitoringSheet = ss.getSheetByName("Monitoring JKT");
  const contactSheet = ss.getSheetByName("Import Contact List");

  Logger.log("CN Tool exists: " + (cnToolSheet ? "Yes" : "No"));
  Logger.log("Monitoring JKT exists: " + (monitoringSheet ? "Yes" : "No"));
  Logger.log("Import Contact List exists: " + (contactSheet ? "Yes" : "No"));

  if (!cnToolSheet || !monitoringSheet || !contactSheet) {
    const missingSheets = [];
    if (!cnToolSheet) missingSheets.push("CN Tool");
    if (!monitoringSheet) missingSheets.push("Monitoring JKT");
    if (!contactSheet) missingSheets.push("Import Contact List");
    const errorMsg = "Error: Missing required sheets - " + missingSheets.join(", ");
    SpreadsheetApp.getUi().alert(errorMsg);
    Logger.log(errorMsg);
    return;
  }

  try {
    const serviceCode = cnToolSheet.getRange("B3").getValue();
    const vesselCode = cnToolSheet.getRange("B4").getValue();
    const podCode = cnToolSheet.getRange("B5").getValue();
    const blNumber = cnToolSheet.getRange("B6").getValue();
    const blNumberInput = cnToolSheet.getRange("B7").getValue();
    const country = cnToolSheet.getRange("B8").getValue();
    const branch = cnToolSheet.getRange("B9").getValue();

    if (!blNumberInput || !country || !branch) {
      SpreadsheetApp.getUi().alert("Error: Please enter a BL Number in B7, select a Country in B8, and a Branch in B9.");
      return;
    }

    const monitoringData = monitoringSheet.getRange("C3:V" + monitoringSheet.getLastRow()).getValues();
    let matchingRowData = null;
    let loiUrl = null;
    for (let i = 0; i < monitoringData.length; i++) {
      if (monitoringData[i][0] === blNumberInput) {
        matchingRowData = monitoringData[i].slice(3, 19);
        loiUrl = monitoringData[i][19];
        break;
      }
    }

    if (!matchingRowData) {
      throw new Error("No matching BL Number found in Monitoring JKT.");
    }

    const contactData = contactSheet.getRange("B2:F" + contactSheet.getLastRow()).getValues();
    let email = "";
    for (let i = 0; i < contactData.length; i++) {
      if (contactData[i][0] === "EUA" && contactData[i][1] === country && contactData[i][2] === branch) {
        email = contactData[i][4];
        break;
      }
    }

    if (!email) {
      throw new Error("No email found for EU Country: " + country + ", Branch: " + branch);
    }

    const tableHtml = generateCorrectionTable(matchingRowData);
    const amendmentType = getAmendmentType(matchingRowData);
    const subject = `${serviceCode}///${vesselCode}///${podCode}///${blNumber}`;
    const body = `
      Please resubmit below details.<br>
      SVVD Code: ${vesselCode}<br>
      POL Code: ${podCode}<br>
      BL Number/s: ${blNumberInput}<br>
      Reason: <b>IE Amendment in:</b> ${amendmentType}<br>
      Responsible Party: ( Shipper )<br><br>
      <b style="color: red;">Correction</b><br><br>
      ${tableHtml}<br><br>
      Thank you and best regards,<br>
      <b>□-----------------------------------□</b><br>
      [Your Name]<br>
      Marketing & Commercial<br>
      Sales Management: CS Desk<br>
      Export Documentation<br>
      <b>□-----------------------------------□</b><br>
      <span style="color: magenta;"><b>PT. OCEAN NETWORK EXPRESS INDONESIA</b></span><br>
      AIA Central, 22nd Floor, Jl. Jenderal Sudirman Kav. 48 A<br>
      Jakarta 12930<br>
      Phone: +62 21 50815150<br>
      DID: +62 21 50889611<br>
      www.one-line.com
    `;

    let emailOptions = {
      to: email,
      subject: subject,
      htmlBody: body
    };

    if (loiUrl) {
      try {
        const fileId = extractFileIdFromUrl(loiUrl);
        const loiFile = DriveApp.getFileById(fileId);
        emailOptions.attachments = [loiFile.getBlob()];
        Logger.log("Attaching LOI file: " + loiFile.getName());
      } catch (err) {
        Logger.log("Failed to attach LOI file: " + err.message);
      }
    }

    Logger.log(`Sending EU CN email to: ${email}, Subject: ${subject}`);
    MailApp.sendEmail(emailOptions);

    SpreadsheetApp.getUi().alert("EU CN sent successfully to " + email);
  } catch (error) {
    SpreadsheetApp.getUi().alert("Failed to send EU CN: " + error.message);
    Logger.log("Error in sendEUCN: " + error.message);
  }
}

function sendCorrectionNotice(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("sendCorrectionNotice - Spreadsheet: " + ss.getName() + ", Row: " + row);
  const monitoringSheet = ss.getSheetByName("Monitoring JKT");
  const contactSheet = ss.getSheetByName("Import Contact List");

  Logger.log("Monitoring JKT exists: " + (monitoringSheet ? "Yes" : "No"));
  Logger.log("Import Contact List exists: " + (contactSheet ? "Yes" : "No"));

  if (!monitoringSheet || !contactSheet) {
    const missingSheets = [];
    if (!monitoringSheet) missingSheets.push("Monitoring JKT");
    if (!contactSheet) missingSheets.push("Import Contact List");
    const errorMsg = "Error: Missing required sheets - " + missingSheets.join(", ");
    SpreadsheetApp.getUi().alert(errorMsg);
    Logger.log(errorMsg);
    return;
  }

  const statusCell = monitoringSheet.getRange(row, 60); // BH
  statusCell.setValue("Sending...");

  try {
    const blNumber = monitoringSheet.getRange(row, 3).getValue();
    const region = monitoringSheet.getRange(row, 57).getValue();
    const country = monitoringSheet.getRange(row, 58).getValue();
    const branch = monitoringSheet.getRange(row, 59).getValue();

    const contactData = contactSheet.getRange("B2:F" + contactSheet.getLastRow()).getValues();
    let email = "";
    for (let i = 0; i < contactData.length; i++) {
      if (contactData[i][0] === region && contactData[i][1] === country && contactData[i][2] === branch) {
        email = contactData[i][4];
        break;
      }
    }

    if (!email) {
      throw new Error("No email found for selected Region, Country, and Branch!");
    }

    const trialData = monitoringSheet.getRange(row, 6, 1, 15).getValues()[0]; // F:U
    const loiUrl = monitoringSheet.getRange(row, 22).getValue(); // V
    const tableHtml = generateCorrectionTable(trialData);
    const amendmentType = getAmendmentType(trialData);

    const euCountries = [
      "Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czech Republic", "Denmark",
      "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Ireland", "Italy",
      "Latvia", "Lithuania", "Luxembourg", "Malta", "Netherlands", "Poland", "Portugal",
      "Romania", "Slovakia", "Slovenia", "Spain", "Sweden"
    ];

    let subject, body;
    if (country === "USA") {
      subject = `US24 Request for AMS Resubmission BL NO: ${blNumber}`;
      body = `
        <b>Dear USCA Offshore ACMSDP Team,</b><br><br>
        <b>Please find below C/N for your reference,</b><br><br>
        <b>Please be informed that OPUS has updated the below amendment as per the customer's request. Thanks.</b><br><br>
        <b>BL No.</b>              :       ${blNumber}<br>
        <b>POL</b>                  :       JAKARTA, INDONESIA<br>
        <b>Type of amendment</b>   :       ${amendmentType}<br><br>
        <b style="color: red;">Correction</b><br><br>
        ${tableHtml}<br><br>
        Thank you and best regards,<br>
        <b>□-----------------------------------□</b><br>
        [Your Name]<br>
        Marketing & Commercial<br>
        Sales Management: CS Desk<br>
        Export Documentation<br>
        <b>□-----------------------------------□</b><br>
        <span style="color: magenta;"><b>PT. OCEAN NETWORK EXPRESS INDONESIA</b></span><br>
        AIA Central, 22nd Floor, Jl. Jenderal Sudirman Kav. 48 A<br>
        Jakarta 12930<br>
        Phone: +62 21 50815150<br>
        DID: +62 21 50889611<br>
        www.one-line.com
      `;
    } else if (euCountries.includes(country)) {
      subject = `Correction Notice After BDR for BL ${blNumber}`;
      body = `
        Dear Colleagues,<br><br>
        We have updated the data in OPUS for BL ${blNumber}. Please find the correction details below:<br><br>
        ${tableHtml}<br><br>
        Thank you and best regards,<br>
        <b>□-----------------------------------□</b><br>
        [Your Name]<br>
        Marketing & Commercial<br>
        Sales Management: CS Desk<br>
        Export Documentation<br>
        <b>□-----------------------------------□</b><br>
        <span style="color: magenta;"><b>PT. OCEAN NETWORK EXPRESS INDONESIA</b></span><br>
        AIA Central, 22nd Floor, Jl. Jenderal Sudirman Kav. 48 A<br>
        Jakarta 12930<br>
        Phone: +62 21 50815150<br>
        DID: +62 21 50889611<br>
        www.one-line.com
      `;
    } else {
      subject = `Correction Notice After BDR for BL ${blNumber}`;
      body = `
        Dear Colleagues,<br><br>
        We have updated the data in OPUS for BL ${blNumber}. Please find the correction details below:<br><br>
        ${tableHtml}<br><br>
        Thank you and best regards,<br>
        <b>□-----------------------------------□</b><br>
        [Your Name]<br>
        Marketing & Commercial<br>
        Sales Management: CS Desk<br>
        Export Documentation<br>
        <b>□-----------------------------------□</b><br>
        <span style="color: magenta;"><b>PT. OCEAN NETWORK EXPRESS INDONESIA</b></span><br>
        AIA Central, 22nd Floor, Jl. Jenderal Sudirman Kav. 48 A<br>
        Jakarta 12930<br>
        Phone: +62 21 50815150<br>
        DID: +62 21 50889611<br>
        www.one-line.com
      `;
    }

    let emailOptions = {
      to: email,
      subject: subject,
      htmlBody: body
    };

    if (loiUrl) {
      try {
        const fileId = extractFileIdFromUrl(loiUrl);
        const loiFile = DriveApp.getFileById(fileId);
        emailOptions.attachments = [loiFile.getBlob()];
        Logger.log("Attaching LOI file: " + loiFile.getName());
      } catch (err) {
        Logger.log("Failed to attach LOI file: " + err.message);
      }
    }

    Logger.log(`Sending email to: ${email}, Subject: ${subject}`);
    MailApp.sendEmail(emailOptions);

    statusCell.setValue("Sent " + new Date().toLocaleString());
  } catch (error) {
    statusCell.setValue("Failed: " + error.message);
    SpreadsheetApp.getUi().alert("Email failed: " + error.message);
    Logger.log("Error in sendCorrectionNotice: " + error.message);
  }
}

// New function to extract file ID from Google Drive URL
function extractFileIdFromUrl(url) {
  const regex = /\/file\/d\/(.+?)\/(?:view|edit)?/;
  const match = url.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  throw new Error("Invalid Google Drive URL format");
}

function generateCorrectionTable(trialData) {
  let table = `
    <table border='1' cellpadding='5'>
      <tr style="background-color: #8B008B; color: white;">
        <th>Correction Items</th>
        <th>Originally Made Out</th>
        <th>To Be Amended To Read</th>
      </tr>`;
  const corrections = [
    ["Shipper", trialData[0], trialData[1]], // F, G
    ["Consignee", trialData[2], trialData[3]], // H, I
    ["Notify Party", trialData[4], trialData[5]], // J, K
    ["Marks & No’s", trialData[6], trialData[7]], // L, M
    ["No. of Package & Description of Goods", trialData[8], trialData[9]], // N, O
    ["Container No./Seal No.", trialData[10], trialData[11]], // P, Q
    ["Gross Weight(KGS)/Measurement(m3)", trialData[12], trialData[13]], // R, S
    ["Others", trialData[14], trialData[15]] // T, U
  ];

  corrections.forEach(([item, prev, curr]) => {
    if (prev || curr) {
      table += `<tr><td>${item}</td><td>${prev || "N/A"}</td><td>${curr || "N/A"}</td></tr>`;
    }
  });
  table += "</table>";
  return table;
}

function getAmendmentType(trialData) {
  const correctionHeaders = [
    "Shipper", "Consignee", "Notify Party", "Marks & No’s",
    "No. of Package & Description of Goods", "Container No./Seal No.",
    "Gross Weight(KGS)/Measurement(m3)", "Others"
  ];

  for (let i = 0; i < correctionHeaders.length; i++) {
    const prev = trialData[i * 2];
    const curr = trialData[i * 2 + 1];
    if (prev || curr) {
      return correctionHeaders[i];
    }
  }
  return "Unknown";
}