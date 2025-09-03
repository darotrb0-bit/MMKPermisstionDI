/**
 * @fileoverview Google Apps Script for an advanced, high-performance Leave Request System.
 * This script powers a web application for employee leave submissions and a comprehensive admin panel.
 * It is optimized with caching, paginated data fetching, and consolidated server calls.
 * It now includes interactive Telegram notifications for direct approval/rejection.
 *
 * @version 32.3.0
 * @author Gemini
 * @changelog
 * - v32.3.0: Fixed critical bug where Telegram notifications failed due to oversized `callback_data`. Switched from using the full admin name to a short admin key (`mmk110011`) in the callback data to stay within Telegram's 64-byte limit. Updated `doPost` to look up the full name from the key. This is the definitive fix for the notification issue.
 * - v32.2.1: Enhanced `sendTelegramNotification` with robust error logging. It now uses muteHttpExceptions to capture the exact API response from Telegram, aiding in debugging delivery failures for specific bots.
 * - v32.2.0: Added a `testTelegramBots` diagnostic function to verify connectivity and permissions for each configured bot, helping to debug notification delivery issues.
 * - v32.1.1: Updated the admin action bot token to a new one as requested.
 * - v32.1.0: Fixed a bug where interactive Telegram notifications failed to send due to spaces in the callback_data. Encoded admin names for safety and improved callback parsing logic.
 * - v32.0.0: Added interactive Telegram notifications. New requests sent to a dedicated admin channel include "Approve" and "Reject" buttons. Implemented a doPost(e) webhook to handle button callbacks from Telegram, allowing admins to process requests directly from the app. Added a setWebhook() utility function for easy setup.
 * - v31.3.1: Fixed a critical bug in `getLeaveRequestDetails` where an undefined variable `row` was used instead of `rowData`, causing receipt generation to fail.
 * - v31.3.0: Updated "Number of Days" to accept text values ("មួយព្រឹក", "មួយរសៀល", "ពេលយប់") for half-day leaves. Added new disabled leave type "ច្បាប់ទៅផ្ទះ". Adjusted backend logic to handle new text-based day values for calculations and notifications.
 * - v31.2.3: Added checkEmployeeWorkStatus function to validate leave request eligibility based on work status.
 * - v31.2.2: Fixed a bug in `checkForDuplicateRequests` that incorrectly blocked users from updating a rejected request for the same day. The function now accepts an optional `requestId` to ignore during the check.
 */

// --- CONFIGURATION ---
const ADMIN_ID = "0011";
const ADMIN_ROLES = {
  "110011DR": "ជំនួយការពិសេសលោកគ្រូដារ៉ូ",
  mmk110011: "គណៈគ្រប់គ្រង លោកគ្រូ ពៅ ដារ៉ូ",
};
const DEFAULT_ADMIN_NAME = "Admin";
const SELFIE_DRIVE_FOLDER_ID = "1FfrIA-8qQOZw_DtyyQX92w-WKNfCABxM";
const DOCUMENT_DRIVE_FOLDER_ID = "1wfUbYwh6SKEqBAxhwyf5q1VMyHxLKuYn";
const PAYMENT_RECEIPT_DRIVE_FOLDER_ID = "1wfUbYwh6SKEqBAxhwyf5q1VMyHxLKuYn";

// --- TELEGRAM CONFIGURATION (SUPPORTS MULTIPLE CHANNELS) ---
const TELEGRAM_BOT_TOKENS = [
  "8482863332:AAHgcH6AjcFpsj4I-jW0s6OV31G-LTkKwFo", // Original Bot (Notification Only)
  "8251688661:AAGFtQi8pNUK9v-4KYea3p015eYGznn2h3A", // NEW Admin Action Bot (Pao Daro)
];
const TELEGRAM_CHAT_IDS = [
  "-1002558667768", // Original Chat ID
  "1487065922", // NEW Admin Action Chat ID (Pao Daro)
];
const ACTION_BOT_TOKEN = "8251688661:AAGFtQi8pNUK9v-4KYea3p015eYGznn2h3A"; // Specify the token for the bot that will have action buttons

// --- LOCATION CONFIGURATION ---
const SUBMIT_TARGET_LATITUDE = 11.414483377915474;
const SUBMIT_TARGET_LONGITUDE = 104.763828818174;
const SUBMIT_ALLOWED_RADIUS_METERS = 24;
const CHECKIN_TARGET_LATITUDE = 11.414483377915474;
const CHECKIN_TARGET_LONGITUDE = 104.763828818174;
const CHECKIN_ALLOWED_RADIUS_METERS = 85;

// --- SPREADSHEET & SHEET IDs/NAMES ---
const EMPLOYEE_DATA_SHEET_ID = "1_Kgl8UQXRsVATt_BOHYQjVWYKkRIBA12R-qnsBoSUzc";
const LEAVE_SPREADSHEET_ID = "148ZMKn2FfKIUu3oNYg-DYiOnrnqYjg67h4LzMtulcHI";

const EMPLOYEE_SHEET_NAME = "បញ្ជឺឈ្មោះរួម";
const PERMISSION_SHEET_NAME = "ច្បាប់ចេញក្រៅ";
const LEAVE_SHEET_NAME = "ច្បាប់ឈប់សម្រាក";
const HOME_LEAVE_SHEET_NAME = "ច្បាប់ទៅផ្ទះ";
const ALL_LEAVE_SHEETS = [
  PERMISSION_SHEET_NAME,
  LEAVE_SHEET_NAME,
  HOME_LEAVE_SHEET_NAME,
];

// --- SCRIPT-WIDE CONSTANTS (Column Indices & Caching) ---
const TIMESTAMP_COL = 1,
  REQUEST_ID_COL = 2,
  EMPLOYEE_ID_COL = 3,
  EMPLOYEE_NAME_COL = 4,
  LEAVE_TYPE_COL = 5,
  START_DATE_COL = 6,
  END_DATE_COL = 7,
  DAYS_COL = 8,
  REASON_COL = 9,
  STATUS_COL = 10,
  APPROVER_COL = 11,
  SELFIE_PHOTO_COL = 12,
  APPROVAL_TIMESTAMP_COL = 13,
  DOC_PHOTO_COL = 14,
  LOCATION_LINK_COL = 15,
  CHECKIN_TIMESTAMP_COL = 16,
  CHECKIN_PHOTO_COL = 17,
  CHECKIN_LOCATION_LINK_COL = 18,
  NOTIFICATION_SENT_COL = 19,
  PAYMENT_RECEIPT_COL = 20,
  ADMIN_CHECKIN_NOTE_COL = 21; // Column 21 is now for Admin Check-in Note OR Rejection Reason
const EMPLOYEE_DATA_START_ROW = 9;
const CACHE_EXPIRATION_SECONDS = 3600; // 1 hour
const CACHE_EMPLOYEE_KEY = "employee_data_map";

// --- TELEGRAM WEBHOOK & ACTIONS ---

/**
 * Handles POST requests, specifically for Telegram webhook callbacks.
 * This is triggered when an admin clicks an inline button (Approve/Reject).
 */
function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    const callbackQuery = contents.callback_query;

    if (callbackQuery) {
      const data = callbackQuery.data; // e.g., "approve_REQ12345_mmk110011"
      const message = callbackQuery.message;
      const chatId = message.chat.id;
      const messageId = message.message_id;

      // New parsing logic using the admin key
      const dataParts = data.split("_");
      const action = dataParts[0];
      const requestId = dataParts[1];
      const adminKey = dataParts[2];
      const approverName = ADMIN_ROLES[adminKey] || DEFAULT_ADMIN_NAME; // Look up the full name

      let result;
      if (action === "approve") {
        result = updateRequestStatus(requestId, "Approved", approverName);
      } else if (action === "reject") {
        result = updateRequestStatus(
          requestId,
          "Rejected",
          approverName,
          "Rejected via Telegram"
        );
      }

      if (result && result.status === "success") {
        // Edit the original message to show the result and remove buttons
        const originalMessageText = message.text;
        const newText =
          originalMessageText +
          `\n\n------------------------------------\n<b>${
            action === "approve" ? "✅ Approved" : "❌ Rejected"
          } by: ${approverName}</b>`;
        editTelegramMessage(chatId, messageId, newText);
      } else {
        // If the action failed (e.g., already processed), notify the admin
        const errorText = `⚠️ Action Failed!\n${
          result ? result.message : "Unknown error."
        }`;
        editTelegramMessage(
          chatId,
          messageId,
          message.text + "\n\n" + errorText
        );
      }
    }
  } catch (err) {
    Logger.log(`doPost Error: ${err.stack}`);
  }
  return ContentService.createTextOutput(
    JSON.stringify({ status: "ok" })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Edits an existing Telegram message to update its content, typically to remove the action buttons.
 * @param {string|number} chatId The chat ID of the message.
 * @param {string|number} messageId The message ID to edit.
 * @param {string} text The new text for the message.
 */
function editTelegramMessage(chatId, messageId, text) {
  const url = `https://api.telegram.org/bot${ACTION_BOT_TOKEN}/editMessageText`;
  const payload = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      chat_id: String(chatId),
      message_id: messageId,
      text: text,
      parse_mode: "HTML",
      reply_markup: { inline_keyboard: [] }, // Remove buttons
    }),
  };
  try {
    UrlFetchApp.fetch(url, payload);
  } catch (e) {
    Logger.log(`Could not edit Telegram message: ${e.message}`);
  }
}

/**
 * UTILITY FUNCTION: Run this manually ONCE from the script editor after deploying.
 * This tells Telegram where to send button click updates.
 */
function setWebhook() {
  const webAppUrl = ScriptApp.getService().getUrl();
  const url = `https://api.telegram.org/bot${ACTION_BOT_TOKEN}/setWebhook?url=${webAppUrl}`;
  try {
    const response = UrlFetchApp.fetch(url);
    Logger.log(response.getContentText());
    SpreadsheetApp.getUi().alert("Webhook set successfully!");
  } catch (e) {
    Logger.log(`Webhook Error: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`Webhook setup failed: ${e.message}`);
  }
}

/**
 * NEW DIAGNOSTIC FUNCTION: Run this manually to test if bots can send messages.
 */
function testTelegramBots() {
  Logger.log("--- Starting Telegram Bot Test ---");
  TELEGRAM_BOT_TOKENS.forEach((token, index) => {
    const chatId = TELEGRAM_CHAT_IDS[index];
    const botType =
      token === ACTION_BOT_TOKEN ? "Action Bot" : "Notification Bot";
    const maskedToken = `${token.substring(0, 12)}...`; // Mask token for security

    Logger.log(
      `Testing ${botType} (Token: ${maskedToken}, Chat ID: ${chatId})`
    );

    const url = `https://api.telegram.org/bot${token}/sendMessage`;
    const payload = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: chatId,
        text: `✅ This is a successful test message for the ${botType} from the Leave Request System.`,
        parse_mode: "HTML",
      }),
      muteHttpExceptions: true, // Important: Prevents script from stopping on error
    };

    try {
      const response = UrlFetchApp.fetch(url, payload);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode === 200) {
        Logger.log(
          `✅ SUCCESS: Message sent successfully to Chat ID ${chatId}.`
        );
      } else {
        Logger.log(
          `❌ FAILED: Received HTTP status ${responseCode} for Chat ID ${chatId}.`
        );
        Logger.log(`   Response from Telegram: ${responseText}`);
        Logger.log(`   Possible Causes:`);
        Logger.log(`   1. Bot Token is incorrect or has been revoked.`);
        Logger.log(`   2. Chat ID is incorrect.`);
        Logger.log(
          `   3. The bot has not been started by the user (for private chat) or is not a member of the group/channel.`
        );
      }
    } catch (e) {
      Logger.log(
        `❌ CRITICAL FAILURE: Could not send message to Chat ID ${chatId}. Error: ${e.stack}`
      );
    }
  });
  Logger.log("--- Telegram Bot Test Finished ---");
  SpreadsheetApp.getUi().alert(
    "Bot test complete. Please check the Execution Log for detailed results (View > Logs)."
  );
}

// --- HELPER FUNCTIONS ---

const dayValueMap = {
  មួយព្រឹក: 0.5,
  មួយរសៀល: 0.5,
  ពេលយប់: 0.5,
};

function getNumericDayValue(dayValue) {
  if (dayValueMap[dayValue]) {
    return dayValueMap[dayValue];
  }
  const numericValue = parseFloat(dayValue);
  return isNaN(numericValue) ? 0 : numericValue;
}

function getFromCache(key) {
  const cache = CacheService.getScriptCache();
  const cachedValue = cache.get(key);
  return cachedValue ? JSON.parse(cachedValue) : null;
}

function putInCache(key, value) {
  const cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(value), CACHE_EXPIRATION_SECONDS);
}

function isBase64Image(str) {
  return typeof str === "string" && str.startsWith("data:image");
}

function calculateDistance(lat1, lon1, lat2, lon2) {
  const R = 6371e3;
  const φ1 = (lat1 * Math.PI) / 180;
  const φ2 = (lat2 * Math.PI) / 180;
  const Δφ = ((lat2 - lat1) * Math.PI) / 180;
  const Δλ = ((lon2 - lon1) * Math.PI) / 180;
  const a =
    Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
    Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function saveImageToDrive(base64Data, fileName, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const decoded = Utilities.base64Decode(base64Data.split(",")[1]);
    const blob = Utilities.newBlob(decoded, MimeType.JPEG, `${fileName}.jpg`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log(`saveImageToDrive Error: ${e.stack}`);
    return null;
  }
}

function findRequestRow(ss, requestId) {
  for (const sheetName of ALL_LEAVE_SHEETS) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
      const requestIds = sheet
        .getRange(2, REQUEST_ID_COL, sheet.getLastRow() - 1, 1)
        .getValues();
      const rowIndex = requestIds.findIndex((row) => row[0] === requestId);
      if (rowIndex !== -1) {
        return { sheet: sheet, row: rowIndex + 2 };
      }
    }
  }
  return null;
}

function getUserStatus(employeeId) {
  if (!employeeId) {
    return { status: "Error", message: "Employee ID is required." };
  }
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    let pendingRequest = null;
    let approvedUncheckedRequest = null;
    let adminCheckedInRequest = null;
    let rejectedRequest = null; // Added to find the latest rejected request

    for (const sheetName of ALL_LEAVE_SHEETS) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet
          .getRange(2, 1, sheet.getLastRow() - 1, ADMIN_CHECKIN_NOTE_COL)
          .getValues();
        for (let i = data.length - 1; i >= 0; i--) {
          const row = data[i];
          const reqEmployeeId = row[EMPLOYEE_ID_COL - 1]
            ? row[EMPLOYEE_ID_COL - 1].toString().trim()
            : "";
          if (reqEmployeeId === employeeId) {
            const reqId = row[REQUEST_ID_COL - 1];
            const reqStatus = row[STATUS_COL - 1];
            const leaveType = row[LEAVE_TYPE_COL - 1];
            const checkInTimestamp = row[CHECKIN_TIMESTAMP_COL - 1];
            const adminCheckinNote = row[ADMIN_CHECKIN_NOTE_COL - 1];

            if (reqStatus === "Pending") {
              pendingRequest = { status: "Pending", requestId: reqId };
              break;
            }
            if (reqStatus === "Rejected" && !rejectedRequest) {
              // Find the most recent rejection
              rejectedRequest = {
                status: "Rejected",
                requestId: reqId,
                reason: adminCheckinNote || "",
              };
            }

            if (
              reqStatus === "Approved" &&
              leaveType === PERMISSION_SHEET_NAME &&
              checkInTimestamp &&
              adminCheckinNote
            ) {
              if (!adminCheckedInRequest) {
                adminCheckedInRequest = {
                  status: "AdminCheckedIn",
                  requestId: reqId,
                };
              }
            }

            if (
              reqStatus === "Approved" &&
              leaveType === PERMISSION_SHEET_NAME &&
              !checkInTimestamp
            ) {
              if (!approvedUncheckedRequest) {
                approvedUncheckedRequest = {
                  status: "Approved",
                  requestId: reqId,
                };
              }
            }
          }
        }
      }
      if (pendingRequest) {
        break;
      }
    }

    if (pendingRequest) return pendingRequest;
    if (approvedUncheckedRequest) return approvedUncheckedRequest;
    if (adminCheckedInRequest) return adminCheckedInRequest;
    if (rejectedRequest) return rejectedRequest; // Return rejected status if found

    return { status: "Clear" };
  } catch (e) {
    Logger.log(`getUserStatus Error: ${e.stack}`);
    return { status: "Error", message: e.message };
  }
}

// FUNCTION TO CHECK FOR DUPLICATE REQUESTS ON THE SAME DAY FOR THE SAME LEAVE TYPE
function checkForDuplicateRequests(details) {
  const {
    employeeId,
    startDate,
    leaveType,
    requestId: currentRequestId,
  } = details;
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(leaveType);
    if (!sheet || sheet.getLastRow() < 2) {
      return { isDuplicate: false };
    }

    const data = sheet
      .getRange(
        2,
        REQUEST_ID_COL,
        sheet.getLastRow() - 1,
        START_DATE_COL - REQUEST_ID_COL + 1
      )
      .getValues();
    const requestDate = new Date(startDate);
    requestDate.setHours(0, 0, 0, 0);

    for (const row of data) {
      const existingRequestId = row[REQUEST_ID_COL - REQUEST_ID_COL];
      const existingEmployeeId = row[EMPLOYEE_ID_COL - REQUEST_ID_COL]
        ? row[EMPLOYEE_ID_COL - REQUEST_ID_COL].toString().trim()
        : "";
      const existingDateValue = row[START_DATE_COL - REQUEST_ID_COL];
      if (existingEmployeeId === employeeId) {
        const existingDate = new Date(existingDateValue);
        existingDate.setHours(0, 0, 0, 0);

        if (existingDate.getTime() === requestDate.getTime()) {
          // If we are updating, and the found request is the one we are updating, it's not a duplicate.
          if (currentRequestId && existingRequestId === currentRequestId) {
            continue;
          }
          return {
            isDuplicate: true,
            message: `អត្តលេខ ${employeeId} បានស្នើសុំ '${leaveType}' សម្រាប់ថ្ងៃនេះរួចហើយ។`,
          };
        }
      }
    }
    return { isDuplicate: false };
  } catch (e) {
    Logger.log(`Error in checkForDuplicateRequests: ${e.stack}`);
    return { isDuplicate: false, error: e.message };
  }
}

/**
 * Checks if an employee is eligible to submit a leave request based on their work status.
 *
 * @param {string} employeeId The ID of the employee to check.
 * @return {object} An object containing a boolean `allowed` and an optional message.
 */
function checkEmployeeWorkStatus(employeeId) {
  try {
    const employeeDataSheet = SpreadsheetApp.openById(
      EMPLOYEE_DATA_SHEET_ID
    ).getSheetByName(EMPLOYEE_SHEET_NAME);
    if (!employeeDataSheet) {
      return { allowed: false, message: "Employee data sheet not found." };
    } // Fetch data from Column E (Employee ID) and F (Work Status)
    const dataRange = employeeDataSheet.getRange(
      `E${EMPLOYEE_DATA_START_ROW}:F${employeeDataSheet.getLastRow()}`
    );
    const values = dataRange.getValues();

    for (const row of values) {
      const id = row[0] ? row[0].toString().trim() : "";
      const workStatus = row[1] ? row[1].toString().trim() : ""; // Check if the current row's employee ID matches and their work status is 'No'
      if (id === employeeId && workStatus === "No") {
        return {
          allowed: false,
          message: "Your work status does not permit this type of leave.",
        };
      }
    }

    return { allowed: true };
  } catch (e) {
    Logger.log(`checkEmployeeWorkStatus Error: ${e.stack}`);
    return { allowed: false, message: `System Error: ${e.message}` };
  }
}

// --- CORE API FUNCTIONS ---

function doGet(e) {
  let page = "index";
  let title = "ប្រព័ន្ធសុំច្បាប់ស្វ័យប្រវត្ត(DI)";
  if (e && e.parameter && e.parameter.page === "admin") {
    page = "admin";
    title = "ផ្ទាំងគ្រប់គ្រង (Admin)";
  }
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(title)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// NEW: Consolidated function for faster verification
function verifyEmployeeAndGetStatus(employeeId) {
  if (!employeeId) {
    return { verificationStatus: "error", message: "សូម​បញ្ចូល​អត្តលេខ។" };
  } // 1. Verify Employee ID first (logic from checkEmployeeId)

  const employeeCheck = checkEmployeeId({ employeeId: employeeId });
  if (employeeCheck.status === "error") {
    return { verificationStatus: "error", message: employeeCheck.message };
  } // 2. If ID is valid, get user's leave status (logic from getUserStatus)

  const leaveStatus = getUserStatus(employeeId);
  if (leaveStatus.status === "Error") {
    return { verificationStatus: "error", message: leaveStatus.message };
  } // 3. Return a combined object

  return {
    verificationStatus: "success",
    employeeInfo: {
      name: employeeCheck.name,
      photoUrl: employeeCheck.photoUrl,
    },
    leaveStatus: leaveStatus,
  };
}

function checkSubmissionLocation(coords) {
  try {
    if (!coords || !coords.latitude || !coords.longitude) {
      throw new Error("Invalid coordinates provided.");
    }
    const distance = calculateDistance(
      coords.latitude,
      coords.longitude,
      SUBMIT_TARGET_LATITUDE,
      SUBMIT_TARGET_LONGITUDE
    );
    if (distance > SUBMIT_ALLOWED_RADIUS_METERS) {
      return {
        status: "error",
        message: `អ្នកនៅឆ្ងាយពីទីតាំងដែលបានកំណត់! (${Math.round(distance)}m)។`,
      };
    }
    return { status: "success", message: "Location is valid." };
  } catch (e) {
    Logger.log(`checkSubmissionLocation Error: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

function checkEmployeeId(employeeData) {
  const { employeeId } = employeeData;
  if (!employeeId || !employeeId.trim()) {
    return { status: "error", message: "សូម​បញ្ចូល​អត្តលេខ។" };
  }
  const trimmedEmployeeId = employeeId.trim();

  try {
    let employeeMap = getFromCache(CACHE_EMPLOYEE_KEY);

    if (!employeeMap) {
      employeeMap = {};
      const employeeSheet = SpreadsheetApp.openById(
        EMPLOYEE_DATA_SHEET_ID
      ).getSheetByName(EMPLOYEE_SHEET_NAME);
      if (employeeSheet.getLastRow() >= EMPLOYEE_DATA_START_ROW) {
        const dataRange = employeeSheet.getRange(
          `A${EMPLOYEE_DATA_START_ROW}:P${employeeSheet.getLastRow()}`
        );
        const formulas = dataRange.getFormulas();
        const values = dataRange.getValues();

        for (let i = 0; i < values.length; i++) {
          const id = values[i][4]; // Column E
          if (id) {
            const name = values[i][11]; // Column L
            const formula = formulas[i][15]; // Column P
            let photoUrl = "";
            if (formula && formula.toUpperCase().includes("IMAGE")) {
              const match = formula.match(/["'](https?:\/\/[^"']+)["']/);
              if (match && match[1]) {
                photoUrl = match[1];
              }
            }
            employeeMap[id.toString().trim()] = {
              name: name || "",
              photoUrl: photoUrl,
            };
          }
        }
      }
      putInCache(CACHE_EMPLOYEE_KEY, employeeMap);
    }

    const employeeInfo = employeeMap[trimmedEmployeeId];
    if (employeeInfo) {
      return {
        status: "success",
        name: employeeInfo.name,
        photoUrl: employeeInfo.photoUrl,
      };
    }
    return { status: "error", message: "អត្តលេខមិនត្រឹមត្រូវ។" };
  } catch (e) {
    Logger.log(`checkEmployeeId Error: ${e.stack}`);
    return { status: "error", message: `System Error: ${e.message}` };
  }
}

/**
 * Calculates monthly leave statistics for a given employee, counting only 'Approved' requests.
 * @param {string} employeeId The ID of the employee.
 * @returns {object} An object containing totalRequests, totalDays, permissionCount, and leaveCount.
 */
function getMonthlyLeaveStats(employeeId) {
  try {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    let stats = {
      totalRequests: 0,
      totalDays: 0,
      permissionCount: 0,
      leaveCount: 0,
    };

    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    for (const sheetName of ALL_LEAVE_SHEETS) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        // Fetch columns up to STATUS_COL to check for 'Approved' status
        const data = sheet
          .getRange(2, 1, sheet.getLastRow() - 1, STATUS_COL)
          .getValues();

        for (const row of data) {
          const rowEmployeeId = row[EMPLOYEE_ID_COL - 1]
            ? row[EMPLOYEE_ID_COL - 1].toString().trim()
            : "";
          if (rowEmployeeId === employeeId) {
            const requestDate = new Date(row[START_DATE_COL - 1]);
            const status = row[STATUS_COL - 1]; // Check if the request is in the current month and year AND is approved

            if (
              requestDate.getMonth() === currentMonth &&
              requestDate.getFullYear() === currentYear &&
              status === "Approved"
            ) {
              const dayValue = row[DAYS_COL - 1];
              const days = getNumericDayValue(dayValue); // Use helper to convert text or number
              const leaveType = row[LEAVE_TYPE_COL - 1];

              stats.totalRequests++;
              stats.totalDays += days;

              if (leaveType === PERMISSION_SHEET_NAME) {
                stats.permissionCount++;
              } else if (leaveType === LEAVE_SHEET_NAME) {
                stats.leaveCount++;
              }
            }
          }
        }
      }
    }
    return stats;
  } catch (e) {
    Logger.log(
      `Error in getMonthlyLeaveStats for employee ${employeeId}: ${e.stack}`
    ); // Return zeroed stats in case of an error
    return {
      totalRequests: 0,
      totalDays: 0,
      permissionCount: 0,
      leaveCount: 0,
    };
  }
}

function submitLeaveRequest(leaveDetails) {
  try {
    if (!leaveDetails || !leaveDetails.leaveType || !leaveDetails.employeeId) {
      throw new Error("Invalid leave details provided.");
    }
    if (!ALL_LEAVE_SHEETS.includes(leaveDetails.leaveType)) {
      // Allow submission even if sheet doesn't exist yet, it will be created.
      // But log a warning if it's not a known primary type.
      if (leaveDetails.leaveType !== HOME_LEAVE_SHEET_NAME) {
        Logger.log(
          `Warning: Submitting to a non-standard sheet: ${leaveDetails.leaveType}`
        );
      }
    }
    let locationLink = "";
    if (leaveDetails.latitude && leaveDetails.longitude) {
      locationLink = `http://maps.google.com/maps?q=${leaveDetails.latitude},${leaveDetails.longitude}`;
    }

    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const sheet =
      ss.getSheetByName(leaveDetails.leaveType) ||
      ss.insertSheet(leaveDetails.leaveType);
    const timestamp = new Date();
    const requestId = `REQ-${timestamp.getTime()}`;
    let selfieUrl = "";
    if (
      leaveDetails.selfieImageData &&
      isBase64Image(leaveDetails.selfieImageData)
    ) {
      selfieUrl = saveImageToDrive(
        leaveDetails.selfieImageData,
        `Selfie_${leaveDetails.employeeId}_${requestId}`,
        SELFIE_DRIVE_FOLDER_ID
      );
    }

    let documentUrlsJson = "";
    if (leaveDetails.documentImageData) {
      documentUrlsJson = saveMultipleImagesToDrive(
        leaveDetails.documentImageData,
        `Document_${leaveDetails.employeeId}_${requestId}`,
        DOCUMENT_DRIVE_FOLDER_ID
      );
    }
    let paymentReceiptUrl = "";
    if (leaveDetails.paymentReceiptImageData) {
      paymentReceiptUrl = saveImageToDrive(
        leaveDetails.paymentReceiptImageData,
        `Payment_${leaveDetails.employeeId}_${requestId}`,
        PAYMENT_RECEIPT_DRIVE_FOLDER_ID
      );
    } // Get stats for *already approved* leaves this month.

    const monthlyStats = getMonthlyLeaveStats(leaveDetails.employeeId);

    sheet.appendRow([
      timestamp,
      requestId,
      leaveDetails.employeeId,
      leaveDetails.employeeName,
      leaveDetails.leaveType,
      new Date(leaveDetails.startDate),
      new Date(leaveDetails.endDate),
      leaveDetails.numberOfDays,
      leaveDetails.reason,
      "Pending",
      "",
      selfieUrl,
      "",
      documentUrlsJson,
      locationLink,
      "",
      "",
      "",
      "",
      paymentReceiptUrl,
      "",
    ]); // --- Build Notification Message ---
    let daysDisplay = leaveDetails.numberOfDays;
    const numericDays = getNumericDayValue(daysDisplay);
    if (numericDays > 0 && !isNaN(parseFloat(daysDisplay))) {
      // Check if it's a number
      daysDisplay += " ថ្ងៃ";
    }

    let notificationMessage = `<b>📢 សំណើសុំច្បាប់ថ្មី</b>\n`;
    notificationMessage += `------------------------------------\n`;
    notificationMessage += `<b>ឈ្មោះ:</b> ${leaveDetails.employeeName}\n`;
    notificationMessage += `<b>ID:</b> ${leaveDetails.employeeId}\n`;
    notificationMessage += `<b>ប្រភេទច្បាប់:</b> ${leaveDetails.leaveType}\n`;
    notificationMessage += `<b>ពីថ្ងៃ:</b> ${leaveDetails.startDate} <b>ដល់</b> ${leaveDetails.endDate}\n`;
    notificationMessage += `<b>ចំនួន:</b> ${daysDisplay}\n`;
    notificationMessage += `<b>មូលហេតុ:</b> ${leaveDetails.reason}`;

    if (selfieUrl) {
      notificationMessage += `\n<b>រូបថត:</b> <a href="${selfieUrl}">មើលរូបថត</a>`;
    }
    if (locationLink) {
      notificationMessage += `\n<b>📍 ទីតាំង:</b> <a href="${locationLink}">ចុចមើលទីតាំង</a>`;
    }
    if (paymentReceiptUrl) {
      notificationMessage += `\n<b>វិក័យបត្រ:</b> <a href="${paymentReceiptUrl}">មើលវិក័យបត្របង់ប្រាក់</a>`;
    } // --- Add Monthly Stats Section (for approved leaves only) ---
    const now = new Date();
    const khmerMonthYear = now.toLocaleString("km-KH", {
      month: "long",
      year: "numeric",
    });

    let statsMessage = `\n\n<b>📊 ប្រវត្តិសុំច្បាប់ (បានអនុម័ត) ${khmerMonthYear}</b>\n`;
    statsMessage += `------------------------------------\n`;
    statsMessage += `<b>- ចំនួនដងសរុប:</b> ${monthlyStats.totalRequests} ដង\n`;
    statsMessage += `<b>- ចំនួនថ្ងៃសរុប:</b> ${monthlyStats.totalDays} ថ្ងៃ\n`;
    statsMessage += `<b>- ច្បាប់ចេញក្រៅ:</b> ${monthlyStats.permissionCount} ដង\n`;
    statsMessage += `<b>- ច្បាប់ឈប់សម្រាក:</b> ${monthlyStats.leaveCount} ដង`;

    notificationMessage += statsMessage;
    notificationMessage += `\n------------------------------------\nសូមធ្វើការសម្រេចចិត្តខាងក្រោម 👇`;

    // --- Create Interactive Keyboard ---
    const adminKey = "mmk110011"; // Use the key for "គណៈគ្រប់គ្រង លោកគ្រូ ពៅ ដារ៉ូ"
    const keyboard = {
      inline_keyboard: [
        [
          {
            text: "✅ យល់ព្រម",
            callback_data: `approve_${requestId}_${adminKey}`,
          },
          {
            text: "❌ បដិសេធ",
            callback_data: `reject_${requestId}_${adminKey}`,
          },
        ],
      ],
    };
    sendTelegramNotification(notificationMessage, keyboard);
    return { status: "success", requestId: requestId };
  } catch (e) {
    Logger.log(`submitLeaveRequest Error: ${e.stack}`);
    return { status: "error", message: `ការដាក់ស្នើបានបរាជ័យ: ${e.message}` };
  }
}

function getAdminDashboardData(options) {
  try {
    const page = options && options.page ? parseInt(options.page, 10) : 1;
    const limit = options && options.limit ? parseInt(options.limit, 10) : 12;
    const filter = options && options.filter ? options.filter : "All";
    const dateFilter =
      options && options.dateFilter ? options.dateFilter : "today_and_pending";

    const requestsData = getLeaveRequests();
    if (requestsData.status !== "success") {
      return {
        status: "error",
        message: "Failed to retrieve leave requests: " + requestsData.message,
      };
    }
    const statsData = getDashboardStats(requestsData.data);
    const uniqueDates = [
      ...new Set(requestsData.data.map((req) => req.startDate)),
    ].sort((a, b) => new Date(b) - new Date(a));

    const now = new Date();
    const todayStart = new Date(new Date().setHours(0, 0, 0, 0));
    const eveningTime = new Date(new Date().setHours(18, 0, 0, 0));

    let dataToPaginate = requestsData.data; // --- DATE FILTER LOGIC ---

    if (dateFilter === "today_and_pending") {
      const todayString = todayStart.toISOString().split("T")[0];
      dataToPaginate = requestsData.data.filter(
        (req) => req.startDate === todayString || req.status === "Pending"
      );
    } else if (dateFilter !== "all") {
      dataToPaginate = requestsData.data.filter(
        (req) => req.startDate === dateFilter
      );
    } // --- STATUS FILTER LOGIC ---

    const filteredRequests = dataToPaginate.filter((req) => {
      if (filter === "All") return true;
      if (filter === "CheckedIn") return !!req.checkInTimestamp;
      if (filter === "OverdueTime") {
        const startDate = new Date(req.startDate);
        return (
          req.leaveType === PERMISSION_SHEET_NAME &&
          req.status === "Approved" &&
          !req.checkInTimestamp &&
          startDate.getTime() === todayStart.getTime() &&
          now.getTime() > eveningTime.getTime()
        );
      }
      if (filter === "OverdueDay") {
        const startDate = new Date(req.startDate);
        return (
          req.leaveType === PERMISSION_SHEET_NAME &&
          req.status === "Approved" &&
          !req.checkInTimestamp &&
          startDate.getTime() < todayStart.getTime()
        );
      }
      return req.status === filter && !req.checkInTimestamp;
    });

    const totalRequests = filteredRequests.length;
    const totalPages = Math.ceil(totalRequests / limit);
    const startIndex = (page - 1) * limit;
    const endIndex = page * limit;
    const paginatedRequests = filteredRequests.slice(startIndex, endIndex);

    return {
      status: "success",
      data: {
        requests: paginatedRequests,
        stats: statsData.data,
        uniqueDates: uniqueDates,
        pagination: {
          currentPage: page,
          totalPages: totalPages,
          totalRequests: totalRequests,
        },
      },
    };
  } catch (e) {
    Logger.log(`getAdminDashboardData Error: ${e.stack}`);
    return {
      status: "error",
      message: `Error fetching dashboard data: ${e.message}`,
    };
  }
}

function getLeaveRequests() {
  try {
    let employeeMap = getFromCache(CACHE_EMPLOYEE_KEY);
    if (!employeeMap) {
      checkEmployeeId({ employeeId: "trigger_cache_build" });
      employeeMap = getFromCache(CACHE_EMPLOYEE_KEY) || {};
    }

    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    let allRequests = [];
    for (const sheetName of ALL_LEAVE_SHEETS) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const values = sheet
          .getRange(2, 1, sheet.getLastRow() - 1, PAYMENT_RECEIPT_COL)
          .getValues();
        const requests = values.map((row) => {
          const employeeId = row[EMPLOYEE_ID_COL - 1]
            ? row[EMPLOYEE_ID_COL - 1].toString().trim()
            : "";
          const employeeInfo = employeeMap[employeeId] || {
            name: row[EMPLOYEE_NAME_COL - 1],
            photoUrl: "",
          };
          return {
            timestamp: new Date(row[TIMESTAMP_COL - 1]).toISOString(),
            requestId: row[REQUEST_ID_COL - 1],
            employeeId: employeeId,
            employeeName: employeeInfo.name,
            leaveType: row[LEAVE_TYPE_COL - 1],
            startDate: new Date(row[START_DATE_COL - 1])
              .toISOString()
              .split("T")[0],
            endDate: new Date(row[END_DATE_COL - 1])
              .toISOString()
              .split("T")[0],
            numberOfDays: row[DAYS_COL - 1],
            reason: row[REASON_COL - 1],
            status: row[STATUS_COL - 1],
            approver: row[APPROVER_COL - 1] || "",
            selfiePhotoUrl: row[SELFIE_PHOTO_COL - 1] || "",
            approvalTimestamp: row[APPROVAL_TIMESTAMP_COL - 1]
              ? new Date(row[APPROVAL_TIMESTAMP_COL - 1]).toISOString()
              : "",
            documentPhotoUrl: row[DOC_PHOTO_COL - 1] || "",
            locationLink: row[LOCATION_LINK_COL - 1] || "",
            checkInTimestamp: row[CHECKIN_TIMESTAMP_COL - 1]
              ? new Date(row[CHECKIN_TIMESTAMP_COL - 1]).toISOString()
              : "",
            checkInPhotoUrl: row[CHECKIN_PHOTO_COL - 1] || "",
            checkInLocationLink: row[CHECKIN_LOCATION_LINK_COL - 1] || "",
            paymentReceiptUrl: row[PAYMENT_RECEIPT_COL - 1] || "",
            photoUrl: employeeInfo.photoUrl,
          };
        });
        allRequests.push(...requests);
      }
    }

    allRequests.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    return { status: "success", data: allRequests };
  } catch (e) {
    Logger.log(`getLeaveRequests Error: ${e.stack}`);
    return { status: "error", message: `Error fetching data: ${e.message}` };
  }
}

function getDashboardStats(allRequests = []) {
  try {
    let stats = {
      pending: 0,
      approved: 0,
      rejected: 0,
      checkedIn: 0,
      overdueTime: 0,
      overdueDay: 0,
      total: allRequests.length,
    };
    const now = new Date();
    const todayStart = new Date(new Date().setHours(0, 0, 0, 0));
    const eveningTime = new Date(new Date().setHours(18, 0, 0, 0));

    for (const req of allRequests) {
      if (req.checkInTimestamp) stats.checkedIn++;
      if (req.status === "Pending") stats.pending++;
      else if (req.status === "Approved") stats.approved++;
      else if (req.status === "Rejected") stats.rejected++;

      const startDate = new Date(req.startDate);
      if (
        req.leaveType === PERMISSION_SHEET_NAME &&
        req.status === "Approved" &&
        !req.checkInTimestamp
      ) {
        if (startDate.getTime() < todayStart.getTime()) {
          stats.overdueDay++;
        } else if (
          startDate.getTime() === todayStart.getTime() &&
          now.getTime() > eveningTime.getTime()
        ) {
          stats.overdueTime++;
        }
      }
    }
    return { status: "success", data: stats };
  } catch (e) {
    Logger.log(`getDashboardStats Error: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

function updateRequestStatus(
  requestId,
  newStatus,
  approverRole,
  rejectionReason = ""
) {
  if (!requestId || !newStatus)
    return {
      status: "error",
      message: "Request ID and new status are required.",
    };
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const found = findRequestRow(ss, requestId);
    if (found) {
      const currentStatus = found.sheet
        .getRange(found.row, STATUS_COL)
        .getValue();
      if (currentStatus !== "Pending") {
        return {
          status: "error",
          message: "This request has already been processed.",
        };
      }
      found.sheet.getRange(found.row, STATUS_COL).setValue(newStatus);
      let approverDisplayName = approverRole || DEFAULT_ADMIN_NAME;
      found.sheet
        .getRange(found.row, APPROVER_COL)
        .setValue(approverDisplayName);
      found.sheet
        .getRange(found.row, APPROVAL_TIMESTAMP_COL)
        .setValue(new Date());

      if (newStatus === "Rejected" && rejectionReason) {
        found.sheet
          .getRange(found.row, ADMIN_CHECKIN_NOTE_COL)
          .setValue(rejectionReason);
      }
      CacheService.getScriptCache().remove(CACHE_EMPLOYEE_KEY);

      const employeeName = found.sheet
        .getRange(found.row, EMPLOYEE_NAME_COL)
        .getValue();
      const employeeId = found.sheet
        .getRange(found.row, EMPLOYEE_ID_COL)
        .getValue();
      const statusEmoji = newStatus === "Approved" ? "✅" : "❌";
      let notificationMessage = `<b>${statusEmoji} សំណើច្បាប់ត្រូវបានសម្រេច</b>\n------------------------------------\n<b>ឈ្មោះ:</b> ${employeeName} (ID: ${employeeId})\n<b>Request ID:</b> ${requestId}\n<b>ស្ថានភាពថ្មី:</b> ${newStatus}\n<b>សម្រេចដោយ:</b> ${approverDisplayName}`;
      if (newStatus === "Rejected" && rejectionReason) {
        notificationMessage += `\n<b>មូលហេតុ:</b> ${rejectionReason}`;
      }
      notificationMessage += `\n------------------------------------`;
      sendTelegramNotification(notificationMessage, null); // Send simple text notification for updates
      return {
        status: "success",
        message: `Request ${requestId} status updated.`,
      };
    }
    return {
      status: "error",
      message: "Request ID not found for status update.",
    };
  } catch (e) {
    Logger.log(`updateRequestStatus Error: ${e.stack}`);
    return {
      status: "error",
      message: `Failed to update status: ${e.message}`,
    };
  }
}

function deleteLeaveRequest(requestId) {
  if (!requestId)
    return { status: "error", message: "Request ID is required for deletion." };
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const found = findRequestRow(ss, requestId);
    if (found) {
      found.sheet.deleteRow(found.row);
      CacheService.getScriptCache().remove(CACHE_EMPLOYEE_KEY);
      return { status: "success", message: `Request ${requestId} deleted.` };
    }
    return { status: "error", message: "Request ID not found for deletion." };
  } catch (e) {
    Logger.log(`deleteLeaveRequest Error: ${e.stack}`);
    return {
      status: "error",
      message: `Failed to delete request: ${e.message}`,
    };
  }
}

function updateLeaveRequest(details) {
  if (!details || !details.requestId) {
    return {
      status: "error",
      message: "Request ID is required for an update.",
    };
  }
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const found = findRequestRow(ss, details.requestId);
    if (!found) {
      return {
        status: "error",
        message: `Request with ID ${details.requestId} not found.`,
      };
    }

    let selfieUrl = details.selfieImageData;
    if (isBase64Image(selfieUrl)) {
      selfieUrl = saveImageToDrive(
        selfieUrl,
        `Selfie_${details.employeeId}_${details.requestId}_updated`,
        SELFIE_DRIVE_FOLDER_ID
      );
    }
    let documentUrlsJson = details.documentImageData;
    if (isBase64Image(documentUrlsJson)) {
      // Handle single new image
      const newUrl = saveImageToDrive(
        documentUrlsJson,
        `Document_${details.employeeId}_${details.requestId}_updated_1`,
        DOCUMENT_DRIVE_FOLDER_ID
      );
      documentUrlsJson = JSON.stringify([newUrl]);
    } else if (documentUrlsJson && documentUrlsJson.startsWith("[")) {
      documentUrlsJson = saveMultipleImagesToDrive(
        documentUrlsJson,
        `Document_${details.employeeId}_${details.requestId}_updated`,
        DOCUMENT_DRIVE_FOLDER_ID
      );
    }

    found.sheet.getRange(found.row, LEAVE_TYPE_COL).setValue(details.leaveType);
    found.sheet
      .getRange(found.row, START_DATE_COL)
      .setValue(new Date(details.startDate));
    found.sheet
      .getRange(found.row, END_DATE_COL)
      .setValue(new Date(details.endDate));
    found.sheet.getRange(found.row, DAYS_COL).setValue(details.numberOfDays);
    found.sheet.getRange(found.row, REASON_COL).setValue(details.reason);
    found.sheet.getRange(found.row, SELFIE_PHOTO_COL).setValue(selfieUrl);
    found.sheet.getRange(found.row, DOC_PHOTO_COL).setValue(documentUrlsJson);
    found.sheet.getRange(found.row, STATUS_COL).setValue("Pending");
    found.sheet.getRange(found.row, APPROVER_COL).clearContent();
    found.sheet.getRange(found.row, APPROVAL_TIMESTAMP_COL).clearContent();
    found.sheet.getRange(found.row, ADMIN_CHECKIN_NOTE_COL).clearContent(); // Clear old rejection reason

    sendTelegramNotification(
      `<b>🔄 សំណើច្បាប់ត្រូវបានកែសម្រួល</b>\n------------------------------------\n<b>ឈ្មោះ:</b> ${details.employeeName}\n<b>ID:</b> ${details.employeeId}\n<b>Request ID:</b> ${details.requestId}\n------------------------------------\nសំណើត្រូវបានដាក់ស្នើឡើងវិញ និងកំពុងរង់ចាំការពិនិត្យ។`
    );

    return { status: "success", requestId: details.requestId };
  } catch (e) {
    Logger.log(`updateLeaveRequest Error: ${e.stack}`);
    return { status: "error", message: `Update failed: ${e.message}` };
  }
}

// --- OTHER UTILITY FUNCTIONS ---

function sendTelegramNotification(message, keyboard = null) {
  if (!TELEGRAM_BOT_TOKENS || TELEGRAM_BOT_TOKENS.length === 0) {
    Logger.log("Telegram Bot Tokens are not set. Skipping notification.");
    return;
  }

  TELEGRAM_BOT_TOKENS.forEach((token, index) => {
    const chatId = TELEGRAM_CHAT_IDS[index];
    if (token && chatId) {
      const url = `https://api.telegram.org/bot${token}/sendMessage`;
      const payloadOptions = {
        text: message,
        parse_mode: "HTML",
        chat_id: chatId,
      };

      // Add the keyboard ONLY for the designated action bot
      if (token === ACTION_BOT_TOKEN && keyboard) {
        payloadOptions.reply_markup = keyboard;
      }

      const payload = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payloadOptions),
        muteHttpExceptions: true, // Capture API errors
      };
      try {
        const response = UrlFetchApp.fetch(url, payload);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        if (responseCode !== 200) {
          Logger.log(
            `Failed to send message to Chat ID ${chatId} (Token: ...${token.slice(
              -6
            )}). Status: ${responseCode}. Response: ${responseText}`
          );
        } else {
          Logger.log(
            `Successfully sent message to Chat ID ${chatId} (Token: ...${token.slice(
              -6
            )})`
          );
        }
      } catch (e) {
        Logger.log(
          `CRITICAL error sending to Chat ID ${chatId} (Token: ...${token.slice(
            -6
          )}): ${e.message}`
        );
      }
    }
  });
}

function getLeaveRequestDetails(requestId) {
  if (!requestId)
    return { status: "error", message: "Request ID is required." };
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const found = findRequestRow(ss, requestId);
    if (found) {
      const rowData = found.sheet
        .getRange(found.row, 1, 1, PAYMENT_RECEIPT_COL)
        .getValues()[0];
      const details = {
        employeeId: rowData[EMPLOYEE_ID_COL - 1],
        employeeName: rowData[EMPLOYEE_NAME_COL - 1],
        leaveType: rowData[LEAVE_TYPE_COL - 1],
        startDate: new Date(rowData[START_DATE_COL - 1])
          .toISOString()
          .split("T")[0],
        endDate: new Date(rowData[END_DATE_COL - 1])
          .toISOString()
          .split("T")[0],
        numberOfDays: rowData[DAYS_COL - 1],
        reason: rowData[REASON_COL - 1],
        status: rowData[STATUS_COL - 1],
        approver: rowData[APPROVER_COL - 1] || "",
        selfiePhotoUrl: rowData[SELFIE_PHOTO_COL - 1] || "",
        approvalTimestamp: rowData[APPROVAL_TIMESTAMP_COL - 1]
          ? new Date(rowData[APPROVAL_TIMESTAMP_COL - 1]).toISOString()
          : "",
        documentPhotoUrl: rowData[DOC_PHOTO_COL - 1] || "",
        locationLink: rowData[LOCATION_LINK_COL - 1] || "",
        checkInTimestamp: rowData[CHECKIN_TIMESTAMP_COL - 1]
          ? new Date(row[CHECKIN_TIMESTAMP_COL - 1]).toISOString()
          : "",
        checkInPhotoUrl: rowData[CHECKIN_PHOTO_COL - 1] || "",
        checkInLocationLink: rowData[CHECKIN_LOCATION_LINK_COL - 1] || "",
        paymentReceiptUrl: rowData[PAYMENT_RECEIPT_COL - 1] || "",
      };
      return { status: "success", data: details };
    }
    return { status: "error", message: "Request details not found." };
  } catch (e) {
    Logger.log(`getLeaveRequestDetails Error: ${e.stack}`);
    return { status: "error", message: `Error fetching details: ${e.message}` };
  }
}

function getRequestStatus(requestId, cacheBuster) {
  if (!requestId)
    return { status: "Error", message: "No Request ID provided." };
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const found = findRequestRow(ss, requestId);
    if (found) {
      const rowData = found.sheet
        .getRange(found.row, 1, 1, ADMIN_CHECKIN_NOTE_COL)
        .getValues()[0];
      const statusValue = rowData[STATUS_COL - 1];
      const leaveType = rowData[LEAVE_TYPE_COL - 1];
      const checkInTimestamp = rowData[CHECKIN_TIMESTAMP_COL - 1];
      const adminCheckinNote = rowData[ADMIN_CHECKIN_NOTE_COL - 1];

      if (statusValue === "Approved" && checkInTimestamp && adminCheckinNote) {
        return { status: "AdminCheckedIn", leaveType: leaveType };
      }
      if (statusValue === "Rejected") {
        return {
          status: "Rejected",
          leaveType: leaveType,
          reason: adminCheckinNote || "",
        };
      }

      return { status: statusValue, leaveType: leaveType };
    }
    return { status: "Not Found" };
  } catch (e) {
    Logger.log(`getRequestStatus Error: ${e.stack}`);
    return { status: "Error", message: e.message };
  }
}

function submitCheckIn(checkInData) {
  if (!checkInData || !checkInData.requestId || !checkInData.selfieData) {
    return { status: "error", message: "Invalid check-in data." };
  }
  try {
    // MODIFIED: Location validation for check-in
    let locationLink = "";
    if (checkInData.latitude && checkInData.longitude) {
      locationLink = `https://maps.google.com/?q=${checkInData.latitude},${checkInData.longitude}`;
      const distance = calculateDistance(
        checkInData.latitude,
        checkInData.longitude,
        CHECKIN_TARGET_LATITUDE,
        CHECKIN_TARGET_LONGITUDE
      );
      if (distance > CHECKIN_ALLOWED_RADIUS_METERS) {
        throw new Error(
          `អ្នកនៅឆ្ងាយពីទីតាំងដែលបានកំណត់! (${Math.round(distance)}m)។`
        );
      }
    } else {
      throw new Error(
        "Location data is missing. Please enable location services."
      );
    }

    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const found = findRequestRow(ss, checkInData.requestId);
    if (found) {
      const employeeId = found.sheet
        .getRange(found.row, EMPLOYEE_ID_COL)
        .getValue();
      const fileName = `CheckIn_${employeeId}_${checkInData.requestId}`;
      const photoUrl = saveImageToDrive(
        checkInData.selfieData,
        fileName,
        SELFIE_DRIVE_FOLDER_ID
      );

      found.sheet
        .getRange(found.row, CHECKIN_TIMESTAMP_COL)
        .setValue(new Date());
      found.sheet.getRange(found.row, CHECKIN_PHOTO_COL).setValue(photoUrl);
      found.sheet
        .getRange(found.row, CHECKIN_LOCATION_LINK_COL)
        .setValue(locationLink);
      const employeeName = found.sheet
        .getRange(found.row, EMPLOYEE_NAME_COL)
        .getValue();
      sendTelegramNotification(
        `<b>✅ បុគ្គលិកបានត្រលប់ចូលវិញ</b>\n------------------------------------\n<b>ឈ្មោះ:</b> ${employeeName}\n<b>ID:</b> ${employeeId}\n<b>Request ID:</b> ${checkInData.requestId}\n<b>រូបថតចូលវិញ:</b> <a href="${photoUrl}">ចុចមើល</a>\n<b>ទីតាំងចូលវិញ:</b> <a href="${locationLink}">ចុចមើល</a>`
      );

      return { status: "success", message: "Check-in successful." };
    }
    return { status: "error", message: "Request not found for check-in." };
  } catch (e) {
    Logger.log(`submitCheckIn Error: ${e.stack}`);
    return { status: "error", message: `Check-in failed: ${e.message}` };
  }
}

function adminSubmitCheckIn(requestId, adminRole) {
  if (!requestId || !adminRole) {
    return {
      status: "error",
      message: "Request ID and Admin Role are required.",
    };
  }
  try {
    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const found = findRequestRow(ss, requestId);
    if (found) {
      const employeeName = found.sheet
        .getRange(found.row, EMPLOYEE_NAME_COL)
        .getValue();
      const employeeId = found.sheet
        .getRange(found.row, EMPLOYEE_ID_COL)
        .getValue();

      found.sheet
        .getRange(found.row, CHECKIN_TIMESTAMP_COL)
        .setValue(new Date());
      found.sheet
        .getRange(found.row, ADMIN_CHECKIN_NOTE_COL)
        .setValue(`Checked in by: ${adminRole}`);
      found.sheet
        .getRange(found.row, CHECKIN_LOCATION_LINK_COL)
        .setValue("N/A (Admin)");
      sendTelegramNotification(
        `<b>✅ បុគ្គលិកបានត្រលប់ចូលវិញ (Admin)</b>\n------------------------------------\n<b>ឈ្មោះ:</b> ${employeeName}\n<b>ID:</b> ${employeeId}\n<b>Request ID:</b> ${requestId}\n<b>បញ្ជាក់ដោយ:</b> ${adminRole}`
      );

      return { status: "success", message: "Admin check-in successful." };
    }
    return {
      status: "error",
      message: "Request not found for admin check-in.",
    };
  } catch (e) {
    Logger.log(`adminSubmitCheckIn Error: ${e.stack}`);
    return { status: "error", message: `Admin check-in failed: ${e.message}` };
  }
}

function checkAdminPassword(password) {
  const role = ADMIN_ROLES[password];
  if (role) {
    const adminUrl = `${ScriptApp.getService().getUrl()}?page=admin&role=${encodeURIComponent(
      role
    )}`;
    return { status: "success", url: adminUrl };
  } else {
    return { status: "error", message: "ពាក្យសម្ងាត់មិនត្រឹមត្រូវ" };
  }
}

function saveMultipleImagesToDrive(imageDataJson, baseFileName, folderId) {
  let images = [];
  try {
    images = JSON.parse(imageDataJson);
  } catch (e) {
    if (isBase64Image(imageDataJson)) {
      images.push(imageDataJson);
    }
  }

  if (!Array.isArray(images) || images.length === 0) {
    return "";
  }

  const urls = images
    .map((base64Data, index) => {
      if (!isBase64Image(base64Data)) {
        if (typeof base64Data === "string" && base64Data.startsWith("http")) {
          return base64Data;
        }
        return null;
      }
      const fileName = `${baseFileName}_${index + 1}`;
      return saveImageToDrive(base64Data, fileName, folderId);
    })
    .filter((url) => url !== null);

  return JSON.stringify(urls);
}

function checkForOverdueCheckIns() {
  try {
    // Step 1: Get employee data map from cache
    let employeeMap = getFromCache(CACHE_EMPLOYEE_KEY);
    if (!employeeMap) {
      checkEmployeeId({ employeeId: "trigger_cache_build" }); // Rebuild cache if empty
      employeeMap = getFromCache(CACHE_EMPLOYEE_KEY) || {};
    }

    const ss = SpreadsheetApp.openById(LEAVE_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(PERMISSION_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log("No permission requests to check.");
      return;
    }

    const dataRange = sheet.getRange(
      2,
      1,
      sheet.getLastRow() - 1,
      NOTIFICATION_SENT_COL
    );
    const values = dataRange.getValues();
    const now = new Date(); // Current time

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const status = row[STATUS_COL - 1];
      const checkInTimestamp = row[CHECKIN_TIMESTAMP_COL - 1];
      const notificationSent = row[NOTIFICATION_SENT_COL - 1]; // Skip if already checked-in or not approved
      if (checkInTimestamp || status !== "Approved") {
        continue;
      }

      const employeeName = row[EMPLOYEE_NAME_COL - 1];
      const employeeId = row[EMPLOYEE_ID_COL - 1]
        ? row[EMPLOYEE_ID_COL - 1].toString().trim()
        : "";
      const dayValue = row[DAYS_COL - 1];
      const numericDays = getNumericDayValue(dayValue);
      const reason = row[REASON_COL - 1];
      const startDate = new Date(row[START_DATE_COL - 1]);
      startDate.setHours(0, 0, 0, 0); // Normalize start date
      const approvalTimestamp = row[APPROVAL_TIMESTAMP_COL - 1]
        ? new Date(row[APPROVAL_TIMESTAMP_COL - 1])
        : null; // Step 2: Get employee photo URL

      const employeeInfo = employeeMap[employeeId] || {};
      const photoUrl =
        employeeInfo.photoUrl ||
        "https://placehold.co/100x100/EFEFEF/AAAAAA&text=No+Img";
      const formattedApprovalTime = approvalTimestamp
        ? approvalTimestamp.toLocaleString("en-GB", {
            dateStyle: "short",
            timeStyle: "short",
          })
        : "N/A";

      let messageTitle = "";
      let messageContext = "";
      let newNotificationStatus = null; // --- Half-Day Logic ---

      if (numericDays == 0.5) {
        const morningApprovalDeadline = new Date(startDate);
        morningApprovalDeadline.setHours(8, 31, 0, 0);

        const checkinDeadline1130 = new Date(startDate);
        checkinDeadline1130.setHours(11, 30, 0, 0);

        const checkinDeadline1430 = new Date(startDate);
        checkinDeadline1430.setHours(14, 30, 0, 0); // 11:30 + 3 hours

        if (approvalTimestamp && approvalTimestamp < morningApprovalDeadline) {
          if (
            now > checkinDeadline1430 &&
            notificationSent === "OverdueHalfDay_1"
          ) {
            messageTitle = `🚨 បុគ្គលិកមិនទាន់ចូលមកវិញ (លើកទី២)`;
            messageContext = `បុគ្គលិកដែលសុំច្បាប់កន្លះថ្ងៃព្រឹក មិនទាន់បានបញ្ជាក់ការចូលមកវិញទេ ក្រោយម៉ោង 2:30 រសៀល។`;
            newNotificationStatus = "OverdueHalfDay_2";
          } else if (now > checkinDeadline1130 && !notificationSent) {
            messageTitle = `⏰ ដល់ម៉ោងចូលមកវិញ`;
            messageContext = `បុគ្គលិកដែលសុំច្បាប់កន្លះថ្ងៃព្រឹក ត្រូវដល់ម៉ោងចូលមកវិញហើយ (11:30 AM)។`;
            newNotificationStatus = "OverdueHalfDay_1";
          }
        }
      } // --- Full-Day Logic ---
      else {
        const overdueDayDeadline = new Date(startDate);
        overdueDayDeadline.setDate(overdueDayDeadline.getDate() + 1);
        overdueDayDeadline.setHours(5, 0, 0, 0); // 5:00 AM the next day

        const overdueTimeDeadline = new Date(startDate);
        overdueTimeDeadline.setHours(18, 0, 0, 0); // 6:00 PM on the same day

        if (now > overdueDayDeadline && notificationSent !== "OverdueDay") {
          messageTitle = `🚨 លើសថ្ងៃចូល`;
          messageContext = `បុគ្គលិកដែលបានសុំច្បាប់ចេញក្រៅកាលពីថ្ងៃទី ${startDate.toLocaleDateString(
            "en-GB"
          )} មិនទាន់បានបញ្ជាក់ការចូលមកវិញទេ។`;
          newNotificationStatus = "OverdueDay";
        } else if (
          now > overdueTimeDeadline &&
          now < overdueDayDeadline &&
          !notificationSent
        ) {
          messageTitle = `⏰ លើសម៉ោងចូល`;
          messageContext = `បុគ្គិកដែលបានសុំច្បាប់ចេញក្រៅថ្ងៃនេះ មិនទាន់បានបញ្ជាក់ការចូលមកវិញទេ ក្រោយម៉ោង 6:00 ល្ងាច។`;
          newNotificationStatus = "OverdueTime";
        }
      }
      if (messageTitle && newNotificationStatus) {
        // Step 3: Construct the detailed message
        let daysDisplay = dayValue;
        if (!isNaN(parseFloat(daysDisplay))) {
          daysDisplay += " ថ្ងៃ";
        }
        const message = `<b>${messageTitle}</b>
------------------------------------
<b>ឈ្មោះ:</b> ${employeeName}
<b>ID:</b> ${employeeId}
<b>មូលហេតុ:</b> ${reason}
<b>ចំនួន:</b> ${daysDisplay}
<b>ម៉ោងអនុម័ត:</b> ${formattedApprovalTime}
<b>រូបថត:</b> <a href="${photoUrl}">ចុចមើលរូបថត</a>
------------------------------------
${messageContext}`;

        sendTelegramNotification(message);
        sheet
          .getRange(i + 2, NOTIFICATION_SENT_COL)
          .setValue(newNotificationStatus);
      }
    }
  } catch (e) {
    Logger.log(`Error in checkForOverdueCheckIns: ${e.stack}`);
    sendTelegramNotification(
      `<b>⚠️ មានបញ្ហាក្នុងការពិនិត្យអ្នកលើសម៉ោង:</b> ${e.message}`
    );
  }
}
