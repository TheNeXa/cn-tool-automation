# Correction Notice Automator

## Overview
The **Correction Notice Automator** is a Google Apps Script solution designed to streamline the process of sending correction notices (CNs) for shipping documentation. Built for logistics teams, it automates email notifications based on bill of lading (BL) updates, pulling data from Google Sheets and integrating with a contact list. With dropdown menus, color-coded visuals, and region-specific email templates, it simplifies workflows for global shipping operations.

## Features
- **Custom Menu**: Adds a "CN Reminder" menu to your Google Sheet with options to set up dropdowns, configure the CN tool, and log sheet names.
- **Dynamic Dropdowns**: Populates region, country, and branch dropdowns in a monitoring sheet, with real-time updates based on user selections.
- **Color Coding**: Highlights countries (e.g., USA, Hong Kong, India) with distinct colors for quick visual reference.
- **Email Automation**: Sends tailored correction notices for India, EU, and other regions, attaching files (e.g., Letters of Indemnity) when available.
- **CN Tool**: A dedicated sheet for generating India and EU correction notices with pre-filled fields and dropdowns.
- **Error Handling**: Alerts users to missing sheets or invalid inputs, with detailed logging for troubleshooting.

## How It Works
1. **Setup**: Initialize dropdowns and the CN tool via the custom menu.
2. **Data Entry**: Input BL numbers and correction details in the monitoring or CN tool sheets.
3. **Automation**: Trigger emails manually (via buttons) or automatically (on edit) with formatted tables and attachments.
4. **Tracking**: Logs send status in the sheet with timestamps or error messages.

## Key Code Highlights
### Custom Menu Setup
```javascript
function onOpen() {
  SpreadsheetApp.getUi().createMenu('CN Reminder')
    .addItem('Setup Dropdowns', 'setupDropdowns')
    .addItem('Setup CN Tool', 'setupCNTool')
    .addItem('Log Sheet Names', 'logSheetNames')
    .addToUi();
}
```

### Dynamic Dropdowns
```javascript
function updateDependentDropdowns(row) {
  const region = monitoringSheet.getRange(row, 57).getValue();
  const countries = contactData.filter(row => row[0] === region).map(row => row[1]);
  monitoringSheet.getRange(row, 58).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(countries, true).build()
  );
}
```

### Email Sending (India CN Example)
```javascript
function sendIndiaCN() {
  const blNumber = cnToolSheet.getRange("E5").getValue();
  const branch = cnToolSheet.getRange("E6").getValue();
  const tableHtml = generateCorrectionTable(matchingRowData);
  MailApp.sendEmail({
    to: email,
    subject: `IND//MUM CORRECTION NOTICE AFTER BDR FOR BL ${blNumber}`,
    htmlBody: `Dear Colleagues,<br>${tableHtml}<br>Thank you,<br>[Your Name]`
  });
}
```

## Use Case
This tool is perfect for logistics teams managing shipping corrections across regions like India, the EU, and the USA. It:
- **Saves Time**: Automates email drafting and data lookup.
- **Reduces Errors**: Validates inputs and matches BL numbers to contact data.
- **Improves Clarity**: Sends professional, formatted emails with correction details.
- **Scales Globally**: Handles region-specific workflows with ease.

## Prerequisites
- A Google Sheets environment with Apps Script enabled.
- Permissions for Spreadsheet, Drive, and Mail services in Apps Script.

## Setup Instructions
1. **Sheets Required**:
   - `Monitoring Sheet`: Tracks BL numbers and correction data (e.g., columns C:V for data, BE:BH for dropdowns/status).
   - `Contact List`: Stores region, country, branch, and email data (e.g., columns B:F).
   - `CN Tool` (optional): Created automatically by `setupCNTool()` for manual CN generation.

2. **Steps**:
   - Copy the script into your Google Apps Script editor.
   - Run `onOpen()` to add the "CN Reminder" menu.
   - Use `Setup Dropdowns` to configure the monitoring sheet.
   - Use `Setup CN Tool` to create and configure the CN tool sheet.
   - Add two buttons in the `CN Tool` sheet:
     - "SEND INDIA CN" (assign to `sendIndiaCN`).
     - "SEND EU CN" (assign to `sendEUCN`).

3. **Permissions**: Grant access to Spreadsheet, Drive (for attachments), and Mail services when prompted.

## Usage
- **Automatic Sending**: Edit the branch column (BG) in the monitoring sheet to trigger a correction notice.
- **Manual Sending**: Use the `CN Tool` sheet to input BL numbers and send India or EU notices via buttons.
- **Debugging**: Run `Log Sheet Names` to check sheet setup in the Apps Script logs.

## Customization
- **Dropdowns**: Modify `setupDropdowns` and `setupCNTool` to adjust regions, countries, or branches.
- **Email Templates**: Edit `sendIndiaCN`, `sendEUCN`, or `sendCorrectionNotice` to tweak subjects and bodies.
- **Colors**: Update `applyColorCoding` to change country-specific highlights.

## Notes
- Replace `[Your Name]` in email templates with your actual name or a variable.
- Ensure the contact list includes valid emails in column F.
- Test with a small dataset before scaling to avoid email quotas.

Feel free to fork this repo, tweak the code, and submit pull requests with improvements or bug fixes!
