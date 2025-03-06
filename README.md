# CN Tool Automation

This Google Apps Script automates the creation and management of Correction Notices (CN) for shipment data in Google Sheets, supporting India and EU regions with dropdowns and email notifications.

## Features
- **Dropdown Setup**: Configures dynamic dropdowns in "Monitoring JKT" for regions, countries, and branches.
- **CN Tool Setup**: Initializes "CN Tool" sheet with India and EU CN configurations.
- **Email Automation**: Sends correction notices for India, EU, and other regions with attached LOI files.
- **Color Coding**: Applies visual indicators for USA, Hong Kong, and India in "Monitoring JKT."

## Prerequisites
- Google account with Sheets and Apps Script access.
- Sheets: "Monitoring JKT," "Import Contact List," and optionally "CN Tool."
- Google Drive access for LOI file attachments.

## Installation
1. Create a Google Sheet with required tabs.
2. Open `Extensions > Apps Script`, paste `cn-tool-automation.gs`, and save.
3. Set an `onEdit` trigger for `onEdit` function.
4. Run `Setup Dropdowns` and `Setup CN Tool` from the "CN Reminder" menu.

## Usage
- **Dropdowns**: Select region/country/branch in "Monitoring JKT" (BE:BG) to trigger emails.
- **CN Tool**: Input India/EU data in "CN Tool" and use red box macros to send CN emails.
- **Logging**: Use "Log Sheet Names" to debug sheet setup.

## File Structure
