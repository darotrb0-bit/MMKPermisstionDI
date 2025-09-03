// --- Dependencies ---
const express = require("express");
const bodyParser = require("body-parser");
const path = require("path");
const { google } = require("googleapis");
const fs = require("fs");
const axios = require("axios");
const NodeCache = require("node-cache");
require("dotenv").config();

// --- Express App Initialization ---
const app = express();
const port = process.env.PORT || 3000;
app.use(bodyParser.json({ limit: "50mb" }));
app.use(bodyParser.urlencoded({ extended: true, limit: "50mb" }));
app.use(express.static(path.join(__dirname, "public")));

// --- In-Memory Cache Initialization ---
const scriptCache = new NodeCache({ stdTTL: 3600 }); // Cache for 1 hour

// --- Google API Authentication ---
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
  ],
});
const sheets = google.sheets({ version: "v4", auth });
const drive = google.drive({ version: "v3", auth });

// --- CONFIGURATION (from .env) ---
const ADMIN_ROLES = {
  "110011DR": "ជំនួយការពិសេសលោកគ្រូដារ៉ូ",
  mmk110011: "គណៈគ្រប់គ្រង លោកគ្រូ ពៅ ដារ៉ូ",
};
const DEFAULT_ADMIN_NAME = "Admin";

const EMPLOYEE_DATA_SHEET_ID = process.env.EMPLOYEE_DATA_SHEET_ID;
const LEAVE_SPREADSHEET_ID = process.env.LEAVE_SPREADSHEET_ID;
const SELFIE_DRIVE_FOLDER_ID = process.env.SELFIE_DRIVE_FOLDER_ID;
const DOCUMENT_DRIVE_FOLDER_ID = process.env.DOCUMENT_DRIVE_FOLDER_ID;
const PAYMENT_RECEIPT_DRIVE_FOLDER_ID =
  process.env.PAYMENT_RECEIPT_DRIVE_FOLDER_ID;

const TELEGRAM_BOT_TOKENS = process.env.TELEGRAM_BOT_TOKENS.split(",");
const TELEGRAM_CHAT_IDS = process.env.TELEGRAM_CHAT_IDS.split(",");
const ACTION_BOT_TOKEN = process.env.ACTION_BOT_TOKEN;

// --- SPREADSHEET NAMES & CONSTANTS ---
const EMPLOYEE_SHEET_NAME = "បញ្ជឺឈ្មោះរួម";
const PERMISSION_SHEET_NAME = "ច្បាប់ចេញក្រៅ";
const LEAVE_SHEET_NAME = "ច្បាប់ឈប់សម្រាក";
const HOME_LEAVE_SHEET_NAME = "ច្បាប់ទៅផ្ទះ";
const ALL_LEAVE_SHEETS = [
  PERMISSION_SHEET_NAME,
  LEAVE_SHEET_NAME,
  HOME_LEAVE_SHEET_NAME,
];

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
  ADMIN_CHECKIN_NOTE_COL = 21;
const EMPLOYEE_DATA_START_ROW = 9;
const CACHE_EMPLOYEE_KEY = "employee_data_map";

// --- HELPER FUNCTIONS ---
const dayValueMap = { មួយព្រឹក: 0.5, មួយរសៀល: 0.5, ពេលយប់: 0.5 };
const getNumericDayValue = (dayValue) => {
  if (dayValueMap[dayValue]) return dayValueMap[dayValue];
  const numericValue = parseFloat(dayValue);
  return isNaN(numericValue) ? 0 : numericValue;
};
const isBase64Image = (str) =>
  typeof str === "string" && str.startsWith("data:image");

async function saveImageToDrive(base64Data, fileName, folderId) {
  try {
    const mimeType = base64Data.match(/data:(.*);base64,/)[1];
    const fileExtension = mimeType.split("/")[1] || "jpg";
    const buffer = Buffer.from(base64Data.split(",")[1], "base64");

    const fileMetadata = {
      name: `${fileName}.${fileExtension}`,
      parents: [folderId],
    };
    const media = {
      mimeType: mimeType,
      body: require("stream").Readable.from(buffer),
    };

    const file = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: "id, webViewLink",
    });

    // Make file publicly viewable
    await drive.permissions.create({
      fileId: file.data.id,
      requestBody: { role: "reader", type: "anyone" },
    });

    return file.data.webViewLink.replace(
      "view?usp=drivesdk",
      "view?usp=sharing"
    );
  } catch (e) {
    console.error(`saveImageToDrive Error: ${e.stack}`);
    return null;
  }
}

async function findRequestRow(spreadsheetId, requestId) {
  for (const sheetName of ALL_LEAVE_SHEETS) {
    try {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${sheetName}!B2:B`, // Only need to check the request ID column
      });
      const requestIds = res.data.values;
      if (requestIds && requestIds.length > 0) {
        const rowIndex = requestIds.findIndex((row) => row[0] === requestId);
        if (rowIndex !== -1) {
          return { sheet: sheetName, row: rowIndex + 2 };
        }
      }
    } catch (e) {
      if (e.code !== 400) {
        // Ignore "Sheet not found" errors
        console.error(
          `Error finding request in sheet ${sheetName}: ${e.message}`
        );
      }
    }
  }
  return null;
}

// --- TELEGRAM FUNCTIONS ---
async function sendTelegramNotification(message, keyboard = null) {
  if (!TELEGRAM_BOT_TOKENS || TELEGRAM_BOT_TOKENS.length === 0) {
    console.log("Telegram Bot Tokens are not set. Skipping notification.");
    return;
  }

  for (const [index, token] of TELEGRAM_BOT_TOKENS.entries()) {
    const chatId = TELEGRAM_CHAT_IDS[index];
    if (token && chatId) {
      const url = `https://api.telegram.org/bot${token}/sendMessage`;
      const payloadOptions = {
        text: message,
        parse_mode: "HTML",
        chat_id: chatId,
      };

      if (token === ACTION_BOT_TOKEN && keyboard) {
        payloadOptions.reply_markup = keyboard;
      }

      try {
        const response = await axios.post(url, payloadOptions, {
          headers: { "Content-Type": "application/json" },
        });
        console.log(`Successfully sent message to Chat ID ${chatId}`);
      } catch (e) {
        console.error(
          `CRITICAL error sending to Chat ID ${chatId}: ${
            e.response ? e.response.data.description : e.message
          }`
        );
      }
    }
  }
}

async function editTelegramMessage(chatId, messageId, text) {
  const url = `https://api.telegram.org/bot${ACTION_BOT_TOKEN}/editMessageText`;
  const payload = {
    chat_id: String(chatId),
    message_id: messageId,
    text: text,
    parse_mode: "HTML",
    reply_markup: { inline_keyboard: [] },
  };
  try {
    await axios.post(url, payload);
  } catch (e) {
    console.error(`Could not edit Telegram message: ${e.message}`);
  }
}

// --- CORE LOGIC FUNCTIONS (Adapted for Node.js) ---

async function checkEmployeeId(employeeId) {
  if (!employeeId || !employeeId.trim()) {
    return { status: "error", message: "សូម​បញ្ចូល​អត្តលេខ។" };
  }
  const trimmedEmployeeId = employeeId.trim();

  try {
    let employeeMap = scriptCache.get(CACHE_EMPLOYEE_KEY);
    if (!employeeMap) {
      employeeMap = {};
      const res = await sheets.spreadsheets.get({
        spreadsheetId: EMPLOYEE_DATA_SHEET_ID,
        ranges: [`'${EMPLOYEE_SHEET_NAME}'!A${EMPLOYEE_DATA_START_ROW}:P`],
        includeGridData: true,
      });

      const rows = res.data.sheets[0].data[0].rowData;
      if (rows) {
        for (const row of rows) {
          if (row.values && row.values[4] && row.values[4].formattedValue) {
            const id = row.values[4].formattedValue.toString().trim();
            const name = row.values[11]
              ? row.values[11].formattedValue || ""
              : "";
            let photoUrl = "";
            if (row.values[15] && row.values[15].hyperlink) {
              photoUrl = row.values[15].hyperlink;
            } else if (
              row.values[15] &&
              row.values[15].userEnteredValue &&
              row.values[15].userEnteredValue.formulaValue
            ) {
              const formula = row.values[15].userEnteredValue.formulaValue;
              const match = formula.match(/["'](https?:\/\/[^"']+)["']/);
              if (match && match[1]) photoUrl = match[1];
            }
            employeeMap[id] = { name, photoUrl };
          }
        }
      }
      scriptCache.set(CACHE_EMPLOYEE_KEY, employeeMap);
    }

    const employeeInfo = employeeMap[trimmedEmployeeId];
    return employeeInfo
      ? {
          status: "success",
          name: employeeInfo.name,
          photoUrl: employeeInfo.photoUrl,
        }
      : { status: "error", message: "អត្តលេខមិនត្រឹមត្រូវ។" };
  } catch (e) {
    console.error(`checkEmployeeId Error: ${e.stack}`);
    return { status: "error", message: `System Error: ${e.message}` };
  }
}

async function getUserStatus(employeeId) {
  // This is a simplified version. A full implementation would require fetching and iterating through all leave sheets.
  // For performance in a stateless environment, it's better to query sheets directly for the latest record.
  // However, to maintain logic parity, we'll adapt the original loop structure.
  try {
    let latestRequest = { status: "Clear", timestamp: 0 };
    for (const sheetName of ALL_LEAVE_SHEETS) {
      try {
        const result = await sheets.spreadsheets.values.get({
          spreadsheetId: LEAVE_SPREADSHEET_ID,
          range: `${sheetName}!A2:U`,
        });
        const rows = result.data.values;
        if (!rows) continue;

        for (const row of rows.slice().reverse()) {
          // Iterate from the end
          const reqEmployeeId = row[EMPLOYEE_ID_COL - 1]
            ? row[EMPLOYEE_ID_COL - 1].toString().trim()
            : "";
          if (reqEmployeeId === employeeId) {
            const reqStatus = row[STATUS_COL - 1];
            // Prioritize Pending and Approved-Unchecked-In
            if (
              reqStatus === "Pending" ||
              (reqStatus === "Approved" && !row[CHECKIN_TIMESTAMP_COL - 1])
            ) {
              return { status: reqStatus, requestId: row[REQUEST_ID_COL - 1] };
            }
            // If we find any other status, it's the latest, but keep searching for priority ones.
            if (
              new Date(row[TIMESTAMP_COL - 1]).getTime() >
              latestRequest.timestamp
            ) {
              latestRequest = {
                status: reqStatus,
                requestId: row[REQUEST_ID_COL - 1],
                timestamp: new Date(row[TIMESTAMP_COL - 1]).getTime(),
                reason: row[ADMIN_CHECKIN_NOTE_COL - 1] || "",
              };
            }
          }
        }
      } catch (e) {
        /* Ignore sheet not found */
      }
    }
    return latestRequest;
  } catch (e) {
    console.error(`getUserStatus Error: ${e.stack}`);
    return { status: "Error", message: e.message };
  }
}

async function getMonthlyLeaveStats(employeeId) {
  // This function remains largely for notification purposes.
  // It is computationally expensive. For a high-performance system, consider a summary sheet updated by triggers.
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  let stats = {
    totalRequests: 0,
    totalDays: 0,
    permissionCount: 0,
    leaveCount: 0,
  };

  for (const sheetName of ALL_LEAVE_SHEETS) {
    try {
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: LEAVE_SPREADSHEET_ID,
        range: `${sheetName}!C2:J`,
      });
      const rows = result.data.values;
      if (!rows) continue;

      rows.forEach((row) => {
        const rowEmployeeId = row[0] ? row[0].toString().trim() : "";
        const status = row[STATUS_COL - EMPLOYEE_ID_COL];
        const requestDate = new Date(row[START_DATE_COL - EMPLOYEE_ID_COL]);

        if (
          rowEmployeeId === employeeId &&
          status === "Approved" &&
          requestDate.getMonth() === currentMonth &&
          requestDate.getFullYear() === currentYear
        ) {
          stats.totalRequests++;
          const days = getNumericDayValue(row[DAYS_COL - EMPLOYEE_ID_COL]);
          stats.totalDays += days;
          if (sheetName === PERMISSION_SHEET_NAME) stats.permissionCount++;
          else if (sheetName === LEAVE_SHEET_NAME) stats.leaveCount++;
        }
      });
    } catch (e) {
      /* Ignore sheet not found */
    }
  }
  return stats;
}

// --- API Endpoints ---
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.post("/verifyEmployeeAndGetStatus", async (req, res) => {
  const { employeeId } = req.body;
  if (!employeeId) {
    return res.json({
      verificationStatus: "error",
      message: "សូម​បញ្ចូល​អត្តលេខ។",
    });
  }
  const employeeCheck = await checkEmployeeId(employeeId);
  if (employeeCheck.status === "error") {
    return res.json({
      verificationStatus: "error",
      message: employeeCheck.message,
    });
  }
  const leaveStatus = await getUserStatus(employeeId);
  if (leaveStatus.status === "Error") {
    return res.json({
      verificationStatus: "error",
      message: leaveStatus.message,
    });
  }
  res.json({
    verificationStatus: "success",
    employeeInfo: {
      name: employeeCheck.name,
      photoUrl: employeeCheck.photoUrl,
    },
    leaveStatus: leaveStatus,
  });
});

app.post("/checkForDuplicateRequests", async (req, res) => {
  const {
    employeeId,
    startDate,
    leaveType,
    requestId: currentRequestId,
  } = req.body;
  // Simplified logic for Node.js - checking for duplicates can be resource-intensive.
  // A robust solution might involve a database query. Here we replicate the sheet scan.
  try {
    const result = await sheets.spreadsheets.values.get({
      spreadsheetId: LEAVE_SPREADSHEET_ID,
      range: `${leaveType}!B2:F`,
    });
    const rows = result.data.values;
    if (!rows) return res.json({ isDuplicate: false });

    const requestDate = new Date(startDate);
    requestDate.setHours(0, 0, 0, 0);

    for (const row of rows) {
      const existingRequestId = row[0];
      const existingEmployeeId = row[EMPLOYEE_ID_COL - REQUEST_ID_COL]
        ? row[EMPLOYEE_ID_COL - REQUEST_ID_COL].toString().trim()
        : "";
      if (existingEmployeeId === employeeId) {
        const existingDate = new Date(row[START_DATE_COL - REQUEST_ID_COL]);
        existingDate.setHours(0, 0, 0, 0);
        if (existingDate.getTime() === requestDate.getTime()) {
          if (currentRequestId && existingRequestId === currentRequestId)
            continue;
          return res.json({
            isDuplicate: true,
            message: `អត្តលេខ ${employeeId} បានស្នើសុំ '${leaveType}' សម្រាប់ថ្ងៃនេះរួចហើយ។`,
          });
        }
      }
    }
    res.json({ isDuplicate: false });
  } catch (e) {
    console.error(`Error in checkForDuplicateRequests: ${e.stack}`);
    res.json({ isDuplicate: false, error: e.message });
  }
});

app.post("/submitLeaveRequest", async (req, res) => {
  const leaveDetails = req.body;
  try {
    const timestamp = new Date();
    const requestId = `REQ-${timestamp.getTime()}`;
    let selfieUrl = "";
    let locationLink = "";

    // MODIFIED: Skip selfie and location for 'ច្បាប់ចេញក្រៅ'
    if (leaveDetails.leaveType !== PERMISSION_SHEET_NAME) {
      if (
        leaveDetails.selfieImageData &&
        isBase64Image(leaveDetails.selfieImageData)
      ) {
        selfieUrl = await saveImageToDrive(
          leaveDetails.selfieImageData,
          `Selfie_${leaveDetails.employeeId}_${requestId}`,
          SELFIE_DRIVE_FOLDER_ID
        );
      }
      if (leaveDetails.latitude && leaveDetails.longitude) {
        locationLink = `http://maps.google.com/maps?q=${leaveDetails.latitude},${leaveDetails.longitude}`;
      }
    }

    let documentUrlsJson = "";
    // Document images are only for "ច្បាប់ឈប់សម្រាក", but we check for data just in case.
    if (
      leaveDetails.documentImageData &&
      leaveDetails.documentImageData.startsWith("[")
    ) {
      const images = JSON.parse(leaveDetails.documentImageData);
      const urls = await Promise.all(
        images.map((imgData, i) =>
          saveImageToDrive(
            imgData,
            `Document_${leaveDetails.employeeId}_${requestId}_${i + 1}`,
            DOCUMENT_DRIVE_FOLDER_ID
          )
        )
      );
      documentUrlsJson = JSON.stringify(urls.filter(Boolean));
    }

    let paymentReceiptUrl = "";
    if (leaveDetails.paymentReceiptImageData) {
      paymentReceiptUrl = await saveImageToDrive(
        leaveDetails.paymentReceiptImageData,
        `Payment_${leaveDetails.employeeId}_${requestId}`,
        PAYMENT_RECEIPT_DRIVE_FOLDER_ID
      );
    }

    const monthlyStats = await getMonthlyLeaveStats(leaveDetails.employeeId);

    const newRow = [
      timestamp.toISOString(),
      requestId,
      leaveDetails.employeeId,
      leaveDetails.employeeName,
      leaveDetails.leaveType,
      leaveDetails.startDate,
      leaveDetails.endDate,
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
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: LEAVE_SPREADSHEET_ID,
      range: `${leaveDetails.leaveType}!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [newRow] },
    });

    // Build and send notification
    let daysDisplay = leaveDetails.numberOfDays;
    const numericDays = getNumericDayValue(daysDisplay);
    if (numericDays > 0 && !isNaN(parseFloat(daysDisplay))) {
      daysDisplay += " ថ្ងៃ";
    }

    let notificationMessage = `<b>📢 សំណើសុំច្បាប់ថ្មី</b>\n`;
    notificationMessage += `------------------------------------\n`;
    notificationMessage += `<b>ឈ្មោះ:</b> ${leaveDetails.employeeName} (ID: ${leaveDetails.employeeId})\n`;
    notificationMessage += `<b>ប្រភេទច្បាប់:</b> ${leaveDetails.leaveType}\n`;
    notificationMessage += `<b>ពីថ្ងៃ:</b> ${leaveDetails.startDate} <b>ដល់</b> ${leaveDetails.endDate}\n`;
    notificationMessage += `<b>ចំនួន:</b> ${daysDisplay}\n`;
    notificationMessage += `<b>មូលហេតុ:</b> ${leaveDetails.reason}`;
    if (selfieUrl)
      notificationMessage += `\n<b>រូបថត:</b> <a href="${selfieUrl}">មើលរូបថត</a>`;
    if (locationLink)
      notificationMessage += `\n<b>📍 ទីតាំង:</b> <a href="${locationLink}">ចុចមើលទីតាំង</a>`;
    if (paymentReceiptUrl)
      notificationMessage += `\n<b>វិក័យបត្រ:</b> <a href="${paymentReceiptUrl}">មើលវិក័យបត្រ</a>`;

    const khmerMonthYear = new Date().toLocaleString("km-KH", {
      month: "long",
      year: "numeric",
    });
    notificationMessage += `\n\n<b>📊 ប្រវត្តិសុំច្បាប់ (បានអនុម័ត) ${khmerMonthYear}</b>\n`;
    notificationMessage += `------------------------------------\n`;
    notificationMessage += `<b>- ចំនួនដងសរុប:</b> ${monthlyStats.totalRequests} ដង\n`;
    notificationMessage += `<b>- ចំនួនថ្ងៃសរុប:</b> ${monthlyStats.totalDays} ថ្ងៃ\n`;
    notificationMessage += `<b>- ច្បាប់ចេញក្រៅ:</b> ${monthlyStats.permissionCount} ដង\n`;
    notificationMessage += `<b>- ច្បាប់ឈប់សម្រាក:</b> ${monthlyStats.leaveCount} ដង`;
    notificationMessage += `\n------------------------------------\nសូមធ្វើការសម្រេចចិត្តខាងក្រោម 👇`;

    const adminKey = "mmk110011";
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

    await sendTelegramNotification(notificationMessage, keyboard);

    res.json({ status: "success", requestId: requestId });
  } catch (e) {
    console.error(`submitLeaveRequest Error: ${e.stack}`);
    res.json({
      status: "error",
      message: `ការដាក់ស្នើបានបរាជ័យ: ${e.message}`,
    });
  }
});

// Other endpoints like getRequestStatus, submitCheckIn, etc. would be converted similarly...
// Below are a few key ones for functionality.

app.post("/getRequestStatus", async (req, res) => {
  const { requestId } = req.body;
  if (!requestId)
    return res.json({ status: "Error", message: "No Request ID provided." });
  try {
    const found = await findRequestRow(LEAVE_SPREADSHEET_ID, requestId);
    if (found) {
      const result = await sheets.spreadsheets.values.get({
        spreadsheetId: LEAVE_SPREADSHEET_ID,
        range: `${found.sheet}!A${found.row}:U${found.row}`,
      });
      const rowData = result.data.values[0];
      const statusValue = rowData[STATUS_COL - 1];
      const leaveType = rowData[LEAVE_TYPE_COL - 1];
      const checkInTimestamp = rowData[CHECKIN_TIMESTAMP_COL - 1];
      const adminCheckinNote = rowData[ADMIN_CHECKIN_NOTE_COL - 1];

      if (statusValue === "Approved" && checkInTimestamp && adminCheckinNote) {
        return res.json({ status: "AdminCheckedIn", leaveType });
      }
      if (statusValue === "Rejected") {
        return res.json({
          status: "Rejected",
          leaveType,
          reason: adminCheckinNote || "",
        });
      }
      return res.json({ status: statusValue, leaveType });
    }
    res.json({ status: "Not Found" });
  } catch (e) {
    console.error(`getRequestStatus Error: ${e.stack}`);
    res.json({ status: "Error", message: e.message });
  }
});

app.post("/getLeaveRequestDetails", async (req, res) => {
  // A simplified conversion. This function would mirror getRequestStatus but return more columns.
  // Full implementation would be similar to getRequestStatus.
  // ...
});

app.post("/webhook", async (req, res) => {
  // This is the new doPost
  try {
    const callbackQuery = req.body.callback_query;
    if (callbackQuery) {
      const data = callbackQuery.data;
      const message = callbackQuery.message;
      const chatId = message.chat.id;
      const messageId = message.message_id;

      const dataParts = data.split("_");
      const action = dataParts[0];
      const requestId = dataParts[1];
      const adminKey = dataParts[2];
      const approverName = ADMIN_ROLES[adminKey] || DEFAULT_ADMIN_NAME;

      let result;
      if (action === "approve") {
        result = await updateRequestStatus(requestId, "Approved", approverName);
      } else if (action === "reject") {
        result = await updateRequestStatus(
          requestId,
          "Rejected",
          approverName,
          "Rejected via Telegram"
        );
      }

      if (result && result.status === "success") {
        const newText =
          message.text +
          `\n\n------------------------------------\n<b>${
            action === "approve" ? "✅ Approved" : "❌ Rejected"
          } by: ${approverName}</b>`;
        await editTelegramMessage(chatId, messageId, newText);
      } else {
        const errorText = `⚠️ Action Failed!\n${
          result ? result.message : "Unknown error."
        }`;
        await editTelegramMessage(
          chatId,
          messageId,
          message.text + "\n\n" + errorText
        );
      }
    }
  } catch (err) {
    console.error(`Webhook Error: ${err.stack}`);
  }
  res.status(200).send({ status: "ok" });
});

async function updateRequestStatus(
  requestId,
  newStatus,
  approverRole,
  rejectionReason = ""
) {
  // This function is called internally by the webhook
  if (!requestId || !newStatus)
    return {
      status: "error",
      message: "Request ID and new status are required.",
    };
  try {
    const found = await findRequestRow(LEAVE_SPREADSHEET_ID, requestId);
    if (found) {
      // Check current status first to prevent race conditions
      const currentStatusRes = await sheets.spreadsheets.values.get({
        spreadsheetId: LEAVE_SPREADSHEET_ID,
        range: `${found.sheet}!J${found.row}`,
      });
      if (
        currentStatusRes.data.values &&
        currentStatusRes.data.values[0][0] !== "Pending"
      ) {
        return {
          status: "error",
          message: "This request has already been processed.",
        };
      }

      const values = [[newStatus, approverRole, new Date().toISOString()]];
      if (newStatus === "Rejected") values[0][3] = rejectionReason; // A bit hacky, assumes column order

      await sheets.spreadsheets.values.update({
        spreadsheetId: LEAVE_SPREADSHEET_ID,
        range: `${found.sheet}!J${found.row}`, // Update Status, Approver, Timestamp
        valueInputOption: "USER_ENTERED",
        requestBody: { values },
      });

      if (newStatus === "Rejected" && rejectionReason) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: LEAVE_SPREADSHEET_ID,
          range: `${found.sheet}!U${found.row}`, // Update rejection reason
          valueInputOption: "USER_ENTERED",
          requestBody: { values: [[rejectionReason]] },
        });
      }

      // In a Node.js environment, clearing a server cache is more direct.
      scriptCache.del(CACHE_EMPLOYEE_KEY);

      // Fetch details for notification
      const detailsRes = await sheets.spreadsheets.values.get({
        spreadsheetId: LEAVE_SPREADSHEET_ID,
        range: `${found.sheet}!C${found.row}:D${found.row}`,
      });
      const [employeeId, employeeName] = detailsRes.data.values[0];

      const statusEmoji = newStatus === "Approved" ? "✅" : "❌";
      let notificationMessage = `<b>${statusEmoji} សំណើច្បាប់ត្រូវបានសម្រេច</b>\n------------------------------------\n<b>ឈ្មោះ:</b> ${employeeName} (ID: ${employeeId})\n<b>Request ID:</b> ${requestId}\n<b>ស្ថានភាពថ្មី:</b> ${newStatus}\n<b>សម្រេចដោយ:</b> ${approverRole}`;
      if (newStatus === "Rejected" && rejectionReason) {
        notificationMessage += `\n<b>មូលហេតុ:</b> ${rejectionReason}`;
      }
      notificationMessage += `\n------------------------------------`;

      await sendTelegramNotification(notificationMessage, null);
      return { status: "success" };
    }
    return { status: "error", message: "Request ID not found." };
  } catch (e) {
    console.error(`updateRequestStatus Error: ${e.stack}`);
    return { status: "error", message: e.message };
  }
}

// --- Server Start ---
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
