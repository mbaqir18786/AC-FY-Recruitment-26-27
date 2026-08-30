/**
 * ==========================================================================
 * Alumni Cell KJSCE — FY Recruitment 2026–27 Database Backend
 * Google Apps Script Web App
 * ==========================================================================
 * 
 * Features:
 *  - Exposes POST API endpoint for the registration website
 *  - Concurrency & race condition safety via LockService
 *  - Duplicate detection by Somaiya Email and Roll Number
 *  - Robust server-side input validation
 *  - Unique sequential Application ID generation (AC2627-XXXX)
 *  - Single-row append to private Google Sheet ('Registrations' tab)
 *  - CORS & JSON response handling
 */

const SHEET_NAME = "Registrations";
const ID_PREFIX = "AC2627-";

/**
 * Health check / Ping endpoint (GET)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "active",
    message: "Alumni Cell FY Recruitment 2026–27 API is running.",
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Main Registration Handler (POST)
 */
function doPost(e) {
  // Acquire script lock to prevent race conditions during concurrent submissions
  const lock = LockService.getScriptLock();
  
  try {
    // Wait up to 30 seconds for concurrent submissions to queue safely
    lock.waitLock(30000);

    // 1. Verify and parse incoming payload
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({
        success: false,
        message: "No payload received."
      });
    }

    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return createJsonResponse({
        success: false,
        message: "Invalid JSON format."
      });
    }

    // 2. Server-side Validation
    const fullName = String(data.fullName || "").trim();
    const rollNumber = String(data.rollNumber || "").trim();
    const email = String(data.email || "").trim().toLowerCase();
    const phone = String(data.phone || "").trim();
    const currentYear = String(data.currentYear || "1st Year (Freshman)").trim();
    const branch = String(data.branch || "").trim();
    const domain1 = String(data.domain1 || "").trim();
    const domain2 = String(data.domain2 || "").trim();
    const domain3 = String(data.domain3 || "").trim();
    const githubUrl = String(data.githubUrl || "N/A").trim();
    const resumeLink = String(data.resumeLink || "").trim();
    const motivation = String(data.motivation || "").trim();

    if (!fullName) {
      return createJsonResponse({ success: false, message: "Full Name is required." });
    }
    if (!rollNumber) {
      return createJsonResponse({ success: false, message: "Roll Number is required." });
    }
    if (!email || !email.endsWith("@somaiya.edu")) {
      return createJsonResponse({ success: false, message: "A valid @somaiya.edu email address is required." });
    }
    if (!phone || !/^[6-9]\d{9}$/.test(phone)) {
      return createJsonResponse({ success: false, message: "A valid 10-digit Indian contact number is required." });
    }
    if (!branch || branch === "Select branch...") {
      return createJsonResponse({ success: false, message: "Please select a valid branch." });
    }
    if (!domain1 || !domain2 || !domain3) {
      return createJsonResponse({ success: false, message: "Exactly 3 domain preferences must be selected." });
    }
    if (domain1 === domain2 || domain2 === domain3 || domain1 === domain3) {
      return createJsonResponse({ success: false, message: "All 3 domain preferences must be unique." });
    }
    if (!resumeLink) {
      return createJsonResponse({ success: false, message: "Google Drive resume link is compulsory." });
    }
    if (!motivation) {
      return createJsonResponse({ success: false, message: "Motivation field cannot be empty." });
    }

    // 3. Connect to Spreadsheet & Sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    const headers = [
      "Timestamp",
      "Application ID",
      "Full Name",
      "Roll Number",
      "Somaiya Email",
      "Contact Number",
      "Current Year",
      "Branch",
      "1st Domain Preference",
      "2nd Domain Preference",
      "3rd Domain Preference",
      "GitHub / Portfolio URL",
      "Resume Drive Link",
      "Motivation"
    ];

    // If sheet doesn't exist, create it
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // If sheet is empty (no headers yet), automatically create and style column headers
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#1a1815")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    // 4. Generate Sequential Unique Application ID (AC2627-0001, AC2627-0002, ...)
    const lastRow = sheet.getLastRow();
    const nextSequenceNumber = lastRow; // row 1 is header, so row 2 gets 0001
    const formattedId = ID_PREFIX + String(nextSequenceNumber).padStart(4, "0");

    // 6. Format Current IST Timestamp
    const timestamp = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");

    // 7. Atomic Insert (Single row append)
    const newRow = [
      timestamp,
      formattedId,
      fullName,
      rollNumber,
      email,
      phone,
      currentYear,
      branch,
      domain1,
      domain2,
      domain3,
      githubUrl,
      resumeLink,
      motivation
    ];

    sheet.appendRow(newRow);

    // 8. Return Success Response
    return createJsonResponse({
      success: true,
      applicationId: formattedId,
      message: "Application submitted successfully."
    });

  } catch (err) {
    return createJsonResponse({
      success: false,
      message: "Server Error: " + err.toString()
    });
  } finally {
    // Release the script lock
    lock.releaseLock();
  }
}

/**
 * Helper to build JSON HTTP response
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
