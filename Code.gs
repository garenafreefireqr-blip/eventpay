// ============================================================
// EVENT PAYMENT MANAGEMENT SYSTEM — Google Apps Script Backend
// ============================================================
// SETUP INSTRUCTIONS (read carefully):
// 1. Go to script.google.com → New Project → paste this entire file
// 2. In Google Sheets, create 4 sheets named exactly:
//    Payments | Complaints | Settings | Admins
// 3. In Settings sheet, Row 2 add these values in columns A-F:
//    A2: EventName   B2: UPI_ID   C2: SessionTimeoutMinutes   D2: RazorpayKeyID   E2: OrganizerEmail  F2: MaxAmount
//    A3: Wedding of X & Y   B3: yourname@upi   C3: 30   D3:   E3: your@gmail.com  F3: 10000
// 4. In Admins sheet, Row 1 headers: Username | Password | Role | AccessLevel
//    Row 2: admin | admin123 | superadmin | full
//    (Add more rows for more admins — only you control this sheet)
// 5. In Payments sheet, Row 1 headers:
//    RefID | Date | Time | Name | Village | Phone | Amount | UTR | Status | VerifiedBy | VerifiedAt | Notes
// 6. In Complaints sheet, Row 1 headers:
//    Date | Time | Name | Village | Phone | Email | Complaint | Attachment | Status | AdminReply | RepliedAt
// 7. Deploy → New Deployment → Web App → Execute as Me → Anyone can access
// 8. Copy the deployment URL and paste in config.js as SCRIPT_URL
// ============================================================

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  const action = e.parameter.action;
  const result = handleAction(action, e.parameter, null);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const params = e.parameter;
  const action = params.action;
  const result = handleAction(action, params, e.postData);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAction(action, params, postData) {
  try {
    if (action === "getSettings")    return getSettings();
    if (action === "loginAdmin")     return loginAdmin(params);
    if (action === "insertPayment")  return insertPayment(params);
    if (action === "checkStatus")    return checkStatus(params);
    if (action === "getPayments")    return getPayments(params);
    if (action === "updatePayments") return updatePayments(params);
    if (action === "getComplaints")  return getComplaints(params);
    if (action === "insertComplaint")return insertComplaint(params);
    if (action === "updateComplaint")return updateComplaint(params);
    if (action === "logActivity")    return logActivity(params);
    if (action === "getActivity")    return getActivity(params);
    if (action === "checkAccess")    return checkAccess(params);
    return { error: "Unknown action" };
  } catch(err) {
    return { error: err.message };
  }
}

// ---- SETTINGS ----
function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Settings");
  const headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values  = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => { if(h) obj[h] = values[i]; });
  return obj;
}

// ---- ADMIN LOGIN ----
function loginAdmin(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Admins");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === params.username &&
        String(data[i][1]).trim() === params.password) {
      
      const settings = getSettings();
      const timeout = parseInt(settings.SessionTimeoutMinutes) || 30;
      const expiry = new Date(Date.now() + timeout * 60 * 1000).toISOString();
      const token = Utilities.getUuid();
      
      logActivity({ adminUser: params.username, action: "Login", detail: "Successful login" });
      
      return {
        success: true,
        role: data[i][2] || "admin",
        accessLevel: data[i][3] || "full",
        token: token,
        expiry: expiry
      };
    }
  }
  return { success: false };
}

// ---- CHECK ACCESS (for sheet-controlled access) ----
function checkAccess(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Admins");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === params.username) {
      return { allowed: true, role: data[i][2], accessLevel: data[i][3] };
    }
  }
  return { allowed: false };
}

// ---- INSERT PAYMENT ----
function insertPayment(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Payments");
  const data = sheet.getDataRange().getValues();
  
  // Duplicate phone check
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][5]).trim() === String(params.phone).trim()) {
      return { result: "DuplicatePhone" };
    }
  }
  
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");
  
  sheet.appendRow([
    params.refid,
    dateStr,
    timeStr,
    params.name,
    params.village,
    params.phone,
    Number(params.amount),
    params.utr,
    "Pending",
    "",  // VerifiedBy
    "",  // VerifiedAt
    ""   // Notes
  ]);
  
  // Send confirmation email (free via Gmail)
  try {
    const settings = getSettings();
    if (settings.OrganizerEmail) {
      MailApp.sendEmail({
        to: settings.OrganizerEmail,
        subject: "New Payment: " + params.name + " ₹" + params.amount,
        body: "New payment received:\n\nName: " + params.name +
              "\nVillage: " + params.village +
              "\nPhone: " + params.phone +
              "\nAmount: ₹" + params.amount +
              "\nUTR: " + params.utr +
              "\nRef ID: " + params.refid +
              "\nTime: " + dateStr + " " + timeStr
      });
    }
  } catch(e) {}
  
  return { result: "Inserted" };
}

// ---- CHECK STATUS ----
function checkStatus(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Payments");
  const data = sheet.getDataRange().getValues();
  const last5 = String(params.refid).trim();
  
  for (let i = 1; i < data.length; i++) {
    const refid = String(data[i][0]).trim();
    if (refid.slice(-5) === last5) {
      return {
        found: true,
        refid:   data[i][0],
        date:    data[i][1],
        time:    data[i][2],
        name:    data[i][3],
        village: data[i][4],
        phone:   data[i][5],
        amount:  data[i][6],
        utr:     data[i][7],
        status:  data[i][8]
      };
    }
  }
  return { found: false };
}

// ---- GET PAYMENTS (admin) ----
function getPayments(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Payments");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    row._row = i + 1;
    rows.push(row);
  }
  return { payments: rows };
}

// ---- UPDATE PAYMENTS (bulk save) ----
function updatePayments(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Payments");
  
  const updates = JSON.parse(params.updates);
  updates.forEach(u => {
    // Status is column 9 (index 8, 1-based = col I)
    sheet.getRange(u.row, 9).setValue(u.status);
    sheet.getRange(u.row, 10).setValue(params.adminUser || "admin");
    sheet.getRange(u.row, 11).setValue(new Date().toISOString());
  });
  
  logActivity({ adminUser: params.adminUser, action: "VerifyPayments", detail: updates.length + " records updated" });
  return { result: "Saved" };
}

// ---- INSERT COMPLAINT ----
function insertComplaint(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Complaints");
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");
  
  sheet.appendRow([
    dateStr,
    timeStr,
    params.name,
    params.village,
    params.phone,
    params.email,
    params.complaint,
    params.filedata ? "Attached" : "None",
    "Open",
    "",
    ""
  ]);
  
  // Email notification to organizer
  try {
    const settings = getSettings();
    if (settings.OrganizerEmail) {
      MailApp.sendEmail({
        to: settings.OrganizerEmail,
        subject: "New Complaint from " + params.name,
        body: "Complaint Details:\n\nName: " + params.name +
              "\nVillage: " + params.village +
              "\nPhone: " + params.phone +
              "\nEmail: " + params.email +
              "\nComplaint: " + params.complaint
      });
    }
  } catch(e) {}
  
  return { result: "Inserted" };
}

// ---- GET COMPLAINTS (admin) ----
function getComplaints(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Complaints");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    row._row = i + 1;
    rows.push(row);
  }
  return { complaints: rows };
}

// ---- UPDATE COMPLAINT ----
function updateComplaint(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Complaints");
  const row = parseInt(params.row);
  
  sheet.getRange(row, 9).setValue(params.status);   // Status
  sheet.getRange(row, 10).setValue(params.reply);   // AdminReply
  sheet.getRange(row, 11).setValue(new Date().toISOString()); // RepliedAt
  
  // Send reply email to complainant
  try {
    if (params.email && params.reply) {
      const settings = getSettings();
      MailApp.sendEmail({
        to: params.email,
        subject: "Reply to your complaint — " + (settings.EventName || "Event"),
        body: "Dear " + params.name + ",\n\n" +
              "Your complaint has been reviewed.\n\n" +
              "Admin Reply:\n" + params.reply + "\n\n" +
              "Status: " + params.status + "\n\n" +
              "Thank you,\nEvent Team"
      });
    }
  } catch(e) {}
  
  logActivity({ adminUser: params.adminUser, action: "ReplyComplaint", detail: "Replied to " + params.name });
  return { result: "Updated" };
}

// ---- ACTIVITY LOG ----
function logActivity(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("ActivityLog");
  if (!sheet) {
    sheet = ss.insertSheet("ActivityLog");
    sheet.appendRow(["Date", "Time", "AdminUser", "Action", "Detail"]);
  }
  const now = new Date();
  sheet.appendRow([
    Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MMM-yyyy"),
    Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a"),
    params.adminUser || "",
    params.action || "",
    params.detail || ""
  ]);
  return { result: "Logged" };
}

function getActivity(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("ActivityLog");
  if (!sheet) return { activities: [] };
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = Math.max(1, data.length - 20); i < data.length; i++) {
    rows.push({ date: data[i][0], time: data[i][1], user: data[i][2], action: data[i][3], detail: data[i][4] });
  }
  return { activities: rows.reverse() };
}

// ---- HELPER: verify admin token (basic check) ----
function verifyAdmin(params) {
  if (!params.adminToken) throw new Error("Unauthorized");
  // Token expiry check
  if (params.adminExpiry && new Date() > new Date(params.adminExpiry)) {
    throw new Error("Session expired");
  }
}
