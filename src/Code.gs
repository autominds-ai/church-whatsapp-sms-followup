/**
 * CHURCH WHATSAPP FOLLOW-UP AUTOMATION (v2 - Same-Day + Saturday Reminder)
 * ==========================================================================
 * Built for: a church (client name private)
 *
 * HOW TO INSTALL:
 * 1. Open your Google Sheet -> Extensions -> Apps Script
 * 2. Delete any starter code, paste this whole file in
 * 3. Fill in the CONFIG SHEET (see keys below) with your real values
 * 4. Run setupTriggers() once (Run button > select setupTriggers) to install
 *    the automatic Same-Day trigger, Saturday Reminder trigger, and
 *    form-submit watcher
 * 5. Authorize the script when Google prompts you
 *
 * REQUIRED SHEET TABS:
 * - "Members"        (main database)
 * - "Config"         (settings - key/value pairs)
 * - "Form Responses" (auto-created when you link your Google Form)
 *
 * MEMBERS TAB COLUMNS (row 1 = headers):
 * A: Timestamp | B: Name | C: Phone | D: Category | E: Conversion Date |
 * F: Last Attendance Date | G: Status This Week | H: Last Message Sent |
 * I: Last Message Date | J: Notes | K: Channel
 *
 * Category values: "First Timer", "New Convert", "Regular Member"
 * Status This Week values: "Attended", "Absent"
 * Channel values: "WhatsApp" or "SMS"
 *
 * CONFIG TAB (2 columns: Key | Value) - add these rows:
 * PHONE_NUMBER_ID                 | (production Phone Number ID from Meta)
 * ACCESS_TOKEN                    | (your permanent System User access token)
 * COUNTRY_CODE                    | 234
 * NEXT_SERVICE_TIME               | e.g. "9:00 AM"
 * TEMPLATE_FIRST_TIMER_SAMEDAY    | first_timer_sameday
 * TEMPLATE_NEW_CONVERT_SAMEDAY    | new_convert_sameday
 * TEMPLATE_ABSENTEE_SAMEDAY       | absentee_sameday
 * TEMPLATE_GENERAL_SAMEDAY        | general_sameday_thankyou
 * TEMPLATE_FIRST_TIMER_SATURDAY   | first_timer_saturday_reminder
 * TEMPLATE_NEW_CONVERT_SATURDAY   | new_convert_saturday_reminder
 * TEMPLATE_SATURDAY_REMINDER      | absentee_saturday_reminder  (shared: absentees + regular members)
 * NEW_CONVERT_WEEKS               | 8
 * BULKSMS_USERNAME                | (your account username at portal.nigeriabulksms.com)
 * BULKSMS_PASSWORD                | (your account password at portal.nigeriabulksms.com)
 * BULKSMS_SENDER_ID               | (your approved sender name, e.g. "YourChurch")
 */

const MEMBERS_SHEET = "Members";
const CONFIG_SHEET = "Config";
const FORM_RESPONSES_SHEET = "Form Responses";

// ---------- CONFIG HELPER ----------
function getConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      let value = data[i][1];
      if (Object.prototype.toString.call(value) === "[object Date]") {
        value = Utilities.formatDate(value, Session.getScriptTimeZone() || "Africa/Lagos", "h:mm a");
      }
      config[data[i][0].toString().trim()] = value;
    }
  }
  return config;
}

// ---------- PHONE NUMBER NORMALIZER ----------
function normalizePhone(raw, countryCode) {
  let phone = raw.toString().replace(/[^0-9]/g, "");
  if (phone.startsWith("0")) phone = countryCode + phone.substring(1);
  if (!phone.startsWith(countryCode)) phone = countryCode + phone;
  return phone;
}

// ---------- SEND A WHATSAPP TEMPLATE MESSAGE ----------
function sendWhatsAppTemplate(phone, templateName, bodyParams, config, languageCode) {
  const url = `https://graph.facebook.com/v20.0/${config.PHONE_NUMBER_ID}/messages`;

  const payload = {
    messaging_product: "whatsapp",
    to: phone,
    type: "template",
    template: {
      name: templateName,
      language: { code: languageCode || "en_US" },
      components: [
        {
          type: "body",
          parameters: bodyParams.map(p => ({ type: "text", text: p }))
        }
      ]
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + config.ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  return { success: code === 200, response: response.getContentText() };
}

// ---------- LANGUAGE CODE OVERRIDE ----------
function getTemplateLanguage(templateName, config) {
  return "en";
}

// ---------- EXTRACT WHATSAPP MESSAGE ID FROM A SEND RESPONSE ----------
function getMessageIdFromResponse(responseText) {
  try {
    const parsed = JSON.parse(responseText);
    if (parsed.messages && parsed.messages[0] && parsed.messages[0].id) {
      return parsed.messages[0].id;
    }
  } catch (e) {
    // not parseable JSON - no message id available
  }
  return "";
}

// ---------- SEND SMS VIA NIGERIA BULK SMS ----------
function sendSMS(phone, message, config) {
  const baseUrl = "https://portal.nigeriabulksms.com/api/";
  const params = {
    username: config.BULKSMS_USERNAME,
    password: config.BULKSMS_PASSWORD,
    message: message,
    sender: config.BULKSMS_SENDER_ID,
    mobiles: phone
  };
  const query = Object.keys(params)
    .map(k => encodeURIComponent(k) + "=" + encodeURIComponent(params[k]))
    .join("&");

  const response = UrlFetchApp.fetch(baseUrl + "?" + query, { muteHttpExceptions: true });
  let result;
  try {
    result = JSON.parse(response.getContentText());
  } catch (e) {
    return { success: false, response: "Could not parse response: " + response.getContentText() };
  }
  return { success: result.status === "OK", response: JSON.stringify(result) };
}

// ---------- PLAIN TEXT VERSIONS (for SMS members) ----------
function getSameDayText(category, name) {
  if (category === "First Timer") {
    return `Dear ${name}, thank you for worshipping with us today at [Church Name]. We're glad to have you and trust you were blessed. You are always welcome. We hope to see you in our subsequent services. God bless you.`;
  } else if (category === "New Convert") {
    return `Dear ${name}, congratulations on your decision to follow Christ today. Welcome to the family of God. [Church Name] is here to support you as you grow in faith. If you need further counseling, kindly reach out to the church office. God bless you.`;
  } else if (category === "Absentee") {
    return `Dear ${name}, we missed you at church today at [Church Name]. We trust you're doing well. Please know you are loved and remembered. God bless you.`;
  } else {
    return `Dear ${name}, thank you for worshipping with us today at [Church Name]. Your presence blessed us, and we trust you were edified. God bless you.`;
  }
}

function getSaturdayText(category, name, serviceTime) {
  if (category === "First Timer") {
    return `Dear ${name}, you are warmly invited to our Church Celebration Service tomorrow at [Church Name] by ${serviceTime}. We look forward to worshipping with you. God bless you.`;
  } else if (category === "New Convert") {
    return `Dear ${name}, we're excited to invite you to our Church Celebration Service tomorrow at [Church Name] by ${serviceTime}. We look forward to seeing you. God bless you.`;
  } else {
    return `Dear ${name}, you are warmly invited to our Church Celebration Service tomorrow at [Church Name] by ${serviceTime}. We would love to have you with us. God bless you.`;
  }
}

// ---------- HELPER: determine effective category (handles New Convert graduation) ----------
function getEffectiveCategory(category, conversionDate, statusThisWeek, config) {
  let newCategory = category;
  if (category === "New Convert" && conversionDate) {
    const weeksSince = (new Date() - new Date(conversionDate)) / (1000 * 60 * 60 * 24 * 7);
    if (weeksSince > (config.NEW_CONVERT_WEEKS || 8)) {
      newCategory = "Regular Member";
    }
  }
  const attended = statusThisWeek === "Attended";
  if (category === "First Timer") return { effective: "First Timer", newCategory };
  if (newCategory === "New Convert") return { effective: "New Convert", newCategory };
  if (!attended) return { effective: "Absentee", newCategory };
  return { effective: "General", newCategory };
}

// ---------- SAME-DAY FOLLOW-UP (run Sunday evening, after attendance is marked) ----------
function sendSameDayFollowUps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getDataRange().getValues();

  for (let row = 1; row < data.length; row++) {
    const [timestamp, name, rawPhone, category, conversionDate, lastAttendance,
           statusThisWeek, lastMsgSent, lastMsgDate, notes, channel] = data[row];

    if (!rawPhone || !name) continue;

    const phone = normalizePhone(rawPhone, config.COUNTRY_CODE || "234");
    const { effective, newCategory } = getEffectiveCategory(category, conversionDate, statusThisWeek, config);

    let templateName;
    if (effective === "First Timer") templateName = config.TEMPLATE_FIRST_TIMER_SAMEDAY;
    else if (effective === "New Convert") templateName = config.TEMPLATE_NEW_CONVERT_SAMEDAY;
    else if (effective === "Absentee") templateName = config.TEMPLATE_ABSENTEE_SAMEDAY;
    else templateName = config.TEMPLATE_GENERAL_SAMEDAY;

    const bodyParams = [name];
    let result, sentLabel, messageId = "";

    if (channel === "SMS") {
      const messageText = getSameDayText(effective, name);
      result = sendSMS(phone, messageText, config);
      sentLabel = "SMS SameDay: " + effective;
    } else {
      result = sendWhatsAppTemplate(phone, templateName, bodyParams, config, getTemplateLanguage(templateName, config));
      sentLabel = templateName;
      if (result.success) messageId = getMessageIdFromResponse(result.response);
    }

    sheet.getRange(row + 1, 4).setValue(newCategory);
    sheet.getRange(row + 1, 8).setValue(sentLabel);
    sheet.getRange(row + 1, 9).setValue(new Date());
    sheet.getRange(row + 1, 10).setValue(result.success ? "Sent OK (Same-Day)" : "FAILED: " + result.response);
    if (messageId) {
      sheet.getRange(row + 1, 12).setValue(messageId);
      sheet.getRange(row + 1, 13).setValue("Pending");
    }

    Utilities.sleep(300);
  }
}

// ---------- SATURDAY REMINDER (run Saturday morning, day before next service) ----------
function sendSaturdayReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getDataRange().getValues();
  const serviceTime = config.NEXT_SERVICE_TIME || "9:00 AM";

  for (let row = 1; row < data.length; row++) {
    const [timestamp, name, rawPhone, category, conversionDate, lastAttendance,
           statusThisWeek, lastMsgSent, lastMsgDate, notes, channel] = data[row];

    if (!rawPhone || !name) continue;

    const phone = normalizePhone(rawPhone, config.COUNTRY_CODE || "234");
    const { effective } = getEffectiveCategory(category, conversionDate, statusThisWeek, config);

    let templateName;
    if (effective === "First Timer") templateName = config.TEMPLATE_FIRST_TIMER_SATURDAY;
    else if (effective === "New Convert") templateName = config.TEMPLATE_NEW_CONVERT_SATURDAY;
    else templateName = config.TEMPLATE_SATURDAY_REMINDER;

    const bodyParams = [name, serviceTime];
    let result, sentLabel, messageId = "";

    if (channel === "SMS") {
      const messageText = getSaturdayText(effective, name, serviceTime);
      result = sendSMS(phone, messageText, config);
      sentLabel = "SMS Saturday: " + effective;
    } else {
      result = sendWhatsAppTemplate(phone, templateName, bodyParams, config, getTemplateLanguage(templateName, config));
      sentLabel = templateName;
      if (result.success) messageId = getMessageIdFromResponse(result.response);
    }

    sheet.getRange(row + 1, 8).setValue(sentLabel);
    sheet.getRange(row + 1, 9).setValue(new Date());
    sheet.getRange(row + 1, 10).setValue(result.success ? "Sent OK (Saturday)" : "FAILED: " + result.response);
    if (messageId) {
      sheet.getRange(row + 1, 12).setValue(messageId);
      sheet.getRange(row + 1, 13).setValue("Pending");
    }

    Utilities.sleep(300);
  }
}

// ---------- SAFE TEST: SEND SAME-DAY MESSAGE TO JUST ONE ROW ----------
function sendTestSameDayToRow(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getRange(rowNumber, 1, 1, 11).getValues()[0];
  const [timestamp, name, rawPhone, category, conversionDate, lastAttendance,
         statusThisWeek, lastMsgSent, lastMsgDate, notes, channel] = data;

  if (!rawPhone || !name) {
    Logger.log("Row " + rowNumber + " is missing Name or Phone - nothing sent.");
    return;
  }

  const phone = normalizePhone(rawPhone, config.COUNTRY_CODE || "234");
  const { effective } = getEffectiveCategory(category, conversionDate, statusThisWeek, config);

  let templateName;
  if (effective === "First Timer") templateName = config.TEMPLATE_FIRST_TIMER_SAMEDAY;
  else if (effective === "New Convert") templateName = config.TEMPLATE_NEW_CONVERT_SAMEDAY;
  else if (effective === "Absentee") templateName = config.TEMPLATE_ABSENTEE_SAMEDAY;
  else templateName = config.TEMPLATE_GENERAL_SAMEDAY;

  let result;
  if (channel === "SMS") {
    result = sendSMS(phone, getSameDayText(effective, name), config);
  } else {
    result = sendWhatsAppTemplate(phone, templateName, [name], config, getTemplateLanguage(templateName, config));
  }

  Logger.log("Result: " + JSON.stringify(result));
  sheet.getRange(rowNumber, 10).setValue(result.success ? "TEST Same-Day: Sent OK" : "TEST FAILED: " + result.response);
}

// ---------- SAFE TEST: SEND SATURDAY REMINDER TO JUST ONE ROW ----------
function sendTestSaturdayToRow(rowNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getRange(rowNumber, 1, 1, 11).getValues()[0];
  const [timestamp, name, rawPhone, category, conversionDate, lastAttendance,
         statusThisWeek, lastMsgSent, lastMsgDate, notes, channel] = data;

  if (!rawPhone || !name) {
    Logger.log("Row " + rowNumber + " is missing Name or Phone - nothing sent.");
    return;
  }

  const phone = normalizePhone(rawPhone, config.COUNTRY_CODE || "234");
  const { effective } = getEffectiveCategory(category, conversionDate, statusThisWeek, config);
  const serviceTime = config.NEXT_SERVICE_TIME || "9:00 AM";

  let templateName;
  if (effective === "First Timer") templateName = config.TEMPLATE_FIRST_TIMER_SATURDAY;
  else if (effective === "New Convert") templateName = config.TEMPLATE_NEW_CONVERT_SATURDAY;
  else templateName = config.TEMPLATE_SATURDAY_REMINDER;

  let result;
  if (channel === "SMS") {
    result = sendSMS(phone, getSaturdayText(effective, name, serviceTime), config);
  } else {
    result = sendWhatsAppTemplate(phone, templateName, [name, serviceTime], config, getTemplateLanguage(templateName, config));
  }

  Logger.log("Result: " + JSON.stringify(result));
  sheet.getRange(rowNumber, 10).setValue(result.success ? "TEST Saturday: Sent OK" : "TEST FAILED: " + result.response);
}

// ---------- TEST A TRUSTED GROUP OF SPECIFIC ROWS (Same-Day + Saturday) ----------
function sendToTrustedGroup(rowNumbers) {
  rowNumbers.forEach(function(rowNumber) {
    Logger.log("Sending Same-Day to row " + rowNumber);
    sendTestSameDayToRow(rowNumber);
    Utilities.sleep(500);
    Logger.log("Sending Saturday reminder to row " + rowNumber);
    sendTestSaturdayToRow(rowNumber);
    Utilities.sleep(500);
  });
}

function runTrustedGroupTest() {
  sendToTrustedGroup([12, 25, 47, 63, 88]);
}

function findDuplicateMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
  const data = sheet.getDataRange().getValues();

  const phoneMap = {};
  const nameMap = {};

  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];
    const rawPhone = data[i][2];
    if (!name && !rawPhone) continue;

    const rowNum = i + 1;

    if (rawPhone) {
      const cleanPhone = rawPhone.toString().replace(/[^0-9]/g, "");
      if (!phoneMap[cleanPhone]) phoneMap[cleanPhone] = [];
      phoneMap[cleanPhone].push(rowNum);
    }

    if (name) {
      const cleanName = name.toString().trim().toLowerCase().replace(/\s+/g, " ");
      if (!nameMap[cleanName]) nameMap[cleanName] = [];
      nameMap[cleanName].push({ row: rowNum, phone: rawPhone });
    }
  }

  let output = "=== SAME PHONE NUMBER (likely true duplicates) ===\n";
  let phoneDupeCount = 0;
  for (const phone in phoneMap) {
    if (phoneMap[phone].length > 1) {
      output += "Phone " + phone + " appears in rows: " + phoneMap[phone].join(", ") + "\n";
      phoneDupeCount++;
    }
  }
  if (phoneDupeCount === 0) output += "None found.\n";

  output += "\n=== SAME NAME, DIFFERENT PHONE (review manually) ===\n";
  let nameDupeCount = 0;
  for (const name in nameMap) {
    if (nameMap[name].length > 1) {
      const details = nameMap[name].map(e => "row " + e.row + " (" + e.phone + ")").join(", ");
      output += "\"" + name + "\": " + details + "\n";
      nameDupeCount++;
    }
  }
  if (nameDupeCount === 0) output += "None found.\n";

  Logger.log(output);
  Logger.log("\nSummary: " + phoneDupeCount + " phone-based duplicate group(s), " + nameDupeCount + " name-based match(es) to review.");
}

// ---------- FIND LAST ROW WITH AN ACTUAL MEMBER (ignores stray blank/leftover rows) ----------
function getLastMemberRow(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let row = data.length - 1; row >= 1; row--) {
    const name = data[row][1];
    if (name && name.toString().trim() !== "") {
      return row + 1;
    }
  }
  return 1;
}

// ---------- SAFE INSERT: ALWAYS ADDS RIGHT AFTER THE LAST REAL MEMBER ----------
function addMemberRow(sheet, rowData) {
  const targetRow = getLastMemberRow(sheet) + 1;
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
}

// ---------- ONE-TIME CLEANUP: REMOVE BLANK "GHOST" ROWS ----------
function cleanupBlankRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
  const data = sheet.getDataRange().getValues();

  let deletedCount = 0;
  for (let row = data.length - 1; row >= 1; row--) {
    const name = data[row][1];
    const phone = data[row][2];
    const isBlank = (!name || name.toString().trim() === "") &&
                     (!phone || phone.toString().trim() === "");
    if (isBlank) {
      sheet.deleteRow(row + 1);
      deletedCount++;
    }
  }
  Logger.log("Cleanup complete: " + deletedCount + " blank rows deleted.");
  SpreadsheetApp.getUi().alert("Cleanup complete: " + deletedCount + " blank rows removed.");
}

// ---------- REVIEW RESULTS: SAME-DAY SEND ----------
function reviewLastSendResults() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
  const data = sheet.getDataRange().getValues();

  let sentOK = 0;
  const failures = {};

  for (let row = 1; row < data.length; row++) {
    const name = data[row][1];
    const phone = data[row][2];
    const notes = data[row][9];
    if (!name || !phone || !notes) continue;

    const notesStr = notes.toString();
    if (notesStr.startsWith("Sent OK")) {
      sentOK++;
    } else if (notesStr.startsWith("FAILED")) {
      if (!failures[notesStr]) failures[notesStr] = [];
      failures[notesStr].push("row " + (row + 1) + ": " + name + " (" + phone + ")");
    }
  }

  let output = "Sent OK: " + sentOK + "\n\n=== FAILURES BY TYPE ===\n";
  for (const errMsg in failures) {
    output += "\n" + errMsg + "\n  -> " + failures[errMsg].length + " people:\n  " +
              failures[errMsg].slice(0, 10).join("\n  ") +
              (failures[errMsg].length > 10 ? "\n  ...and " + (failures[errMsg].length - 10) + " more" : "") + "\n";
  }
  Logger.log(output);
}

// ---------- REVIEW RESULTS: BY CATEGORY ----------
function reviewResultsByCategory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getDataRange().getValues();

  const summary = {};

  for (let row = 1; row < data.length; row++) {
    const [timestamp, name, rawPhone, category, conversionDate, lastAttendance,
           statusThisWeek, lastMsgSent, lastMsgDate, notes, channel] = data[row];
    if (!name || !rawPhone || !notes) continue;

    const { effective } = getEffectiveCategory(category, conversionDate, statusThisWeek, config);
    if (!summary[effective]) summary[effective] = { ok: 0, failed: 0 };

    if (notes.toString().startsWith("Sent OK")) summary[effective].ok++;
    else if (notes.toString().startsWith("FAILED")) summary[effective].failed++;
  }

  let output = "=== RESULTS BY CATEGORY ===\n";
  for (const cat in summary) {
    output += cat + ": " + summary[cat].ok + " sent OK, " + summary[cat].failed + " failed\n";
  }
  Logger.log(output);
}

// ---------- SPECIAL PROGRAM ANNOUNCEMENT ----------
function sendSpecialProgramInvite(programDetails, dateTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getDataRange().getValues();

  let sent = 0, failed = 0;
  for (let row = 1; row < data.length; row++) {
    const name = data[row][1];
    const rawPhone = data[row][2];
    const channel = data[row][10];
    if (!rawPhone || !name) continue;

    const phone = normalizePhone(rawPhone, config.COUNTRY_CODE || "234");
    let result;

    if (channel === "SMS") {
      const message = `Dear ${name}, join us for ${programDetails} on ${dateTime}. We look forward to seeing you there. God bless you.`;
      result = sendSMS(phone, message, config);
    } else {
      const templateName = "special_program_invite";
      result = sendWhatsAppTemplate(phone, templateName, [name, programDetails, dateTime], config, getTemplateLanguage(templateName, config));
    }

    if (result.success) sent++; else failed++;
    Utilities.sleep(300);
  }
  Logger.log("Special program announcement: " + sent + " sent, " + failed + " failed.");
}

function runSpecialProgramAnnouncement() {
  sendSpecialProgramInvite("Youth Conference", "Sunday, August 24th at 10 AM");
}

// ---------- SMS FALLBACK: SAME-DAY MESSAGES TO EVERYONE ----------
function sendSameDayFollowUpsSMSFallback() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getDataRange().getValues();

  let sent = 0, failed = 0;
  for (let row = 1; row < data.length; row++) {
    const [timestamp, name, rawPhone, category, conversionDate, lastAttendance,
           statusThisWeek] = data[row];

    if (!rawPhone || !name) continue;

    const phone = normalizePhone(rawPhone, config.COUNTRY_CODE || "234");
    const { effective } = getEffectiveCategory(category, conversionDate, statusThisWeek, config);
    const messageText = getSameDayText(effective, name);
    const result = sendSMS(phone, messageText, config);

    if (result.success) sent++; else failed++;
    sheet.getRange(row + 1, 10).setValue(
      (result.success ? "SMS FALLBACK Same-Day: Sent OK" : "SMS FALLBACK FAILED: " + result.response)
    );
    Utilities.sleep(300);
  }
  Logger.log("Same-Day SMS fallback: " + sent + " sent, " + failed + " failed.");
}

// ---------- SMS FALLBACK: SATURDAY REMINDER TO EVERYONE ----------
function sendSaturdayRemindersSMSFallback() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  const config = getConfig();
  const data = sheet.getDataRange().getValues();
  const serviceTime = config.NEXT_SERVICE_TIME || "9:00 AM";

  let sent = 0, failed = 0;
  for (let row = 1; row < data.length; row++) {
    const [timestamp, name, rawPhone, category, conversionDate, lastAttendance,
           statusThisWeek] = data[row];

    if (!rawPhone || !name) continue;

    const phone = normalizePhone(rawPhone, config.COUNTRY_CODE || "234");
    const { effective } = getEffectiveCategory(category, conversionDate, statusThisWeek, config);
    const messageText = getSaturdayText(effective, name, serviceTime);
    const result = sendSMS(phone, messageText, config);

    if (result.success) sent++; else failed++;
    sheet.getRange(row + 1, 10).setValue(
      (result.success ? "SMS FALLBACK Saturday: Sent OK" : "SMS FALLBACK FAILED: " + result.response)
    );
    Utilities.sleep(300);
  }
  Logger.log("Saturday SMS fallback: " + sent + " sent, " + failed + " failed.");
}

// ---------- HEALTH CHECK: CONFIRM ALL TRIGGER FUNCTIONS ACTUALLY EXIST ----------
function healthCheckBeforeTriggers() {
  const requiredFunctions = [
    "syncChecklistToDatabase",
    "refreshChecklist",
    "sendSameDayFollowUps",
    "sendSaturdayReminders",
    "sendSaturdayRemindersSMSFallback",
    "sendSameDayFollowUpsSMSFallback",
    "onFirstTimerFormSubmit"
  ];

  let missing = [];
  requiredFunctions.forEach(function(fnName) {
    if (typeof this[fnName] !== "function") missing.push(fnName);
  }, this);

  if (missing.length === 0) {
    Logger.log("✅ All " + requiredFunctions.length + " required functions are present. Safe to run setupTriggers().");
  } else {
    Logger.log("⚠️ MISSING FUNCTIONS - do NOT run setupTriggers() yet:\n" + missing.join("\n"));
  }
}

// ---------- MANUAL TEST SHORTCUT ----------
function runMyTest() {
  sendTestSameDayToRow(91); // change 91 to your actual test row number
}

// ---------- USHER CHECKLIST SYNC ----------
// A separate, simple Google Sheet ("Attendance Checklist") is
// shared with ushers instead of the real Members database. It has three tabs:
// "This Week" (Name, Phone, Status dropdown), "New First Timers"
// (Name, Phone, Uses WhatsApp?), and "New Converts" (Name, Phone, Uses
// WhatsApp?) - for anyone saved that Sunday, whether they're already a
// member or brand new. These functions bridge it with Members.
// Add CHECKLIST_SHEET_ID to your Config tab (the ID from the checklist
// sheet's URL) before using these.

function refreshChecklist() {
  const config = getConfig();
  const checklistSS = SpreadsheetApp.openById(config.CHECKLIST_SHEET_ID);
  const checklistSheet = checklistSS.getSheetByName("This Week");
  const membersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
  const data = membersSheet.getDataRange().getValues();

  const lastRow = checklistSheet.getLastRow();
  if (lastRow > 1) checklistSheet.getRange(2, 1, lastRow - 1, 3).clearContent();

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const name = data[i][1];
    const phone = data[i][2];
    if (name && phone) rows.push([name, phone, ""]);
  }
  if (rows.length > 0) {
    checklistSheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  Logger.log("Checklist refreshed with " + rows.length + " members.");
}

function syncChecklistToDatabase() {
  const config = getConfig();
  const checklistSS = SpreadsheetApp.openById(config.CHECKLIST_SHEET_ID);
  const checklistSheet = checklistSS.getSheetByName("This Week");
  const firstTimerSheet = checklistSS.getSheetByName("New First Timers");
  const membersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);

  const checklistData = checklistSheet.getDataRange().getValues();
  const membersData = membersSheet.getDataRange().getValues();

  const phoneToRow = {};
  for (let i = 1; i < membersData.length; i++) {
    const phone = membersData[i][2] ? membersData[i][2].toString().replace(/[^0-9]/g, "") : "";
    if (phone) phoneToRow[phone] = i + 1;
  }

  const today = new Date();
  let updated = 0, unmatched = 0;
  for (let i = 1; i < checklistData.length; i++) {
    const [name, phone, status] = checklistData[i];
    if (!phone || !status) continue;
    const cleanPhone = phone.toString().replace(/[^0-9]/g, "");
    const rowNum = phoneToRow[cleanPhone];
    if (!rowNum) { unmatched++; continue; }

    const statusValue = status.toString().trim().toLowerCase() === "attended" ? "Attended" : "Absent";
    membersSheet.getRange(rowNum, 7).setValue(statusValue);
    if (statusValue === "Attended") membersSheet.getRange(rowNum, 6).setValue(today);
    updated++;
  }

  let added = 0;
  if (firstTimerSheet) {
    const ftData = firstTimerSheet.getDataRange().getValues();
    for (let i = 1; i < ftData.length; i++) {
      const [name, phone, usesWhatsApp] = ftData[i];
      if (!name || !phone) continue;
      const channel = (usesWhatsApp && usesWhatsApp.toString().toLowerCase().startsWith("n")) ? "SMS" : "WhatsApp";
      addMemberRow(membersSheet, [
        new Date(), name, phone, "First Timer", "", new Date(), "Attended",
        "", "", "Added via usher checklist", channel
      ]);
      added++;
    }
    const lastRow = firstTimerSheet.getLastRow();
    if (lastRow > 1) firstTimerSheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  let convertsUpdated = 0, convertsAdded = 0;
  const newConvertSheet = checklistSS.getSheetByName("New Converts");
  if (newConvertSheet) {
    const ncData = newConvertSheet.getDataRange().getValues();
    for (let i = 1; i < ncData.length; i++) {
      const [name, phone, usesWhatsApp] = ncData[i];
      if (!name || !phone) continue;
      const cleanPhone = phone.toString().replace(/[^0-9]/g, "");
      const rowNum = phoneToRow[cleanPhone];

      if (rowNum) {
        membersSheet.getRange(rowNum, 4).setValue("New Convert");
        membersSheet.getRange(rowNum, 5).setValue(today);
        membersSheet.getRange(rowNum, 6).setValue(today);
        membersSheet.getRange(rowNum, 7).setValue("Attended");
        convertsUpdated++;
      } else {
        const channel = (usesWhatsApp && usesWhatsApp.toString().toLowerCase().startsWith("n")) ? "SMS" : "WhatsApp";
        addMemberRow(membersSheet, [
          new Date(), name, phone, "New Convert", today, today, "Attended",
          "", "", "Added via usher checklist (New Convert)", channel
        ]);
        convertsAdded++;
      }
    }
    const lastRow = newConvertSheet.getLastRow();
    if (lastRow > 1) newConvertSheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  Logger.log("Synced: " + updated + " updated, " + unmatched + " unmatched (check phone numbers), " +
             added + " new first-timers added, " + convertsUpdated + " existing members marked as New Convert, " +
             convertsAdded + " brand new New Converts added.");
}

// ---------- HANDLE NEW FIRST-TIMER FORM SUBMISSIONS ----------
function onFirstTimerFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = ss.getSheetByName(MEMBERS_SHEET);

  const values = e.values;
  const timestamp = values[0];
  const name = values[1];
  const phone = values[2];
  const usesWhatsApp = values[6];

  const channel = (usesWhatsApp && usesWhatsApp.toString().toLowerCase().startsWith("n")) ? "SMS" : "WhatsApp";

  addMemberRow(members, [
    timestamp, name, phone, "First Timer", "", new Date(), "Attended",
    "", "", "Auto-added from first-timer form", channel
  ]);
}

// ---------- ONE-TIME SETUP: RUN THIS ONCE ----------
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Sync usher checklist into Members: Sunday 4:00 PM
  ScriptApp.newTrigger("syncChecklistToDatabase")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(16)
    .create();

  // Second checklist sync pass: Sunday 6:50 PM (catches anyone marked
  // between the 4 PM sync and the 7 PM Same-Day send - e.g. if service
  // ran long or someone was added late)
  ScriptApp.newTrigger("syncChecklistToDatabase")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(18)
    .nearMinute(50)
    .create();

  // Refresh checklist with current member list: Saturday 6 PM (ready for Sunday)
  ScriptApp.newTrigger("refreshChecklist")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(18)
    .create();

  // Same-Day follow-up: every Sunday evening at 7 PM
  ScriptApp.newTrigger("sendSameDayFollowUps")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(19)
    .create();

  // Saturday reminder: every Saturday morning at 9 AM
  ScriptApp.newTrigger("sendSaturdayReminders")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(9)
    .create();

  // SMS fallback (Saturday): runs right after, same time window
  ScriptApp.newTrigger("sendSaturdayRemindersSMSFallback")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(9)
    .nearMinute(10)
    .create();

  // SMS fallback (Same-Day): runs a few minutes after the 7 PM WhatsApp
  // trigger, well before Nigeria Bulk SMS's 7:40 PM cutoff
  ScriptApp.newTrigger("sendSameDayFollowUpsSMSFallback")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(19)
    .nearMinute(5)
    .create();

  // Form submit watcher
  ScriptApp.newTrigger("onFirstTimerFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  Logger.log("Triggers installed: Checklist refresh (Sat 6pm), Checklist sync (Sun 4pm + 6:50pm), Same-Day WhatsApp (Sun 7pm), Same-Day SMS (Sun 7:05pm), Saturday Reminder WhatsApp (Sat 9am), Saturday Reminder SMS (Sat 9:10am), form-submit watcher.");
}

// ---------- WHATSAPP DELIVERY STATUS WEBHOOK ----------
// SETUP (one-time):
// 1. In Config tab, add a row: WEBHOOK_VERIFY_TOKEN | (any random string, e.g. "any-random-string")
// 2. In Apps Script: Deploy -> New deployment -> type "Web app"
//      Execute as: Me
//      Who has access: Anyone
//    Click Deploy, copy the Web app URL it gives you.
// 3. In Meta Business Suite -> WhatsApp -> Configuration -> Webhook:
//      Callback URL: paste the Web app URL from step 2
//      Verify token: the SAME string you put in WEBHOOK_VERIFY_TOKEN
//    Click "Verify and save", then subscribe to the "messages" field.

function doGet(e) {
  const config = getConfig();
  const verifyToken = config.WEBHOOK_VERIFY_TOKEN;

  if (e.parameter["hub.mode"] === "subscribe" &&
      e.parameter["hub.verify_token"] === verifyToken) {
    return ContentService.createTextOutput(e.parameter["hub.challenge"]);
  }
  return ContentService.createTextOutput("Verification failed");
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const entry = payload.entry && payload.entry[0];
    const change = entry && entry.changes && entry.changes[0];
    const statuses = change && change.value && change.value.statuses;

    if (statuses && statuses.length > 0) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
      const data = sheet.getDataRange().getValues();

      const idToRow = {};
      for (let i = 1; i < data.length; i++) {
        const msgId = data[i][11];
        if (msgId) idToRow[msgId] = i + 1;
      }

      statuses.forEach(function(s) {
        const rowNum = idToRow[s.id];
        if (rowNum) {
          const when = new Date(parseInt(s.timestamp, 10) * 1000);
          const statusText = s.status + " (" + when.toLocaleString() + ")";
          sheet.getRange(rowNum, 13).setValue(statusText);
        }
      });
    }
  } catch (err) {
    Logger.log("Webhook error: " + err);
  }

  return ContentService.createTextOutput(JSON.stringify({ received: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- CHECK DELIVERY STATUS SUMMARY ----------
function reviewDeliveryStatus() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEMBERS_SHEET);
  const data = sheet.getDataRange().getValues();

  const summary = {};
  let noStatusYet = 0;

  for (let row = 1; row < data.length; row++) {
    const msgId = data[row][11];
    const status = data[row][12];
    if (!msgId) continue;

    if (!status || status === "Pending") {
      noStatusYet++;
      continue;
    }
    const statusKey = status.toString().split(" (")[0];
    summary[statusKey] = (summary[statusKey] || 0) + 1;
  }

  let output = "=== DELIVERY STATUS SUMMARY ===\n";
  for (const key in summary) {
    output += key + ": " + summary[key] + "\n";
  }
  output += "Still pending (no webhook update yet): " + noStatusYet + "\n";
  Logger.log(output);
}