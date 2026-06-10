// ============================================================
// CODE.GS v5 — EventPay — Exact Sheet Column Match
// ============================================================
// SHEETS: Payments | AuditLog | Villages | ActivityLog |
//         UTRBlacklist | Complaints | Settings | Admins
// ============================================================
const SPREADSHEET_ID = "1TsSOerv8tI1oqxrlhdJts5hEyTbY5sfu8m3AD3XxZjM";

function doGet(e) {
  const r = handleAction(e.parameter.action, e.parameter, null);
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  const r = handleAction(e.parameter.action, e.parameter, e.postData);
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function handleAction(action, p, pd) {
  try {
    // Public
    if (action === "getSettings")          return getSettings();
    if (action === "getPublicVisibility")  return getPublicVisibility();
    if (action === "getPublicStats")       return getPublicStats();
    if (action === "getPublicPayments")    return getPublicPayments();
    if (action === "checkStatus")          return checkStatus(p);
    if (action === "insertPayment")        return insertPayment(p);
    if (action === "insertComplaint")      return insertComplaint(p);
    if (action === "getVillageSuggestions")return getVillageSuggestions();
    if (action === "validateUTR")          return validateUTR(p);
    if (action === "getGalleryImages")     return getGalleryImages();
    // Admin
    if (action === "loginAdmin")           return loginAdmin(p);
    if (action === "getPayments")          return getPayments(p);
    if (action === "updatePayments")       return updatePayments(p);
    if (action === "getComplaints")        return getComplaints(p);
    if (action === "updateComplaint")      return updateComplaint(p);
    if (action === "logActivity")          return logActivity(p);
    if (action === "getActivity")          return getActivity(p);
    if (action === "updatePublicDisplay")  return updatePublicDisplay(p);
    if (action === "addVillageSuggestion") return addVillageSuggestion(p);
    if (action === "addUTRBlacklist")      return addUTRBlacklist(p);
    if (action === "getUTRBlacklist")      return getUTRBlacklist(p);
    // Super Admin
    if (action === "updateSettings")       return updateSettings(p);
    if (action === "getAuditLog")          return getAuditLog(p);
    if (action === "getSheetsList")        return getSheetsList(p);
    if (action === "getSheetData")         return getSheetData(p);
    if (action === "updateSheetCell")      return updateSheetCell(p);
    if (action === "addSheetRow")          return addSheetRow(p);
    if (action === "deleteSheetRow")       return deleteSheetRow(p);
    return { error: "Unknown action: " + action };
  } catch(err) { return { error: err.message }; }
}

// ============================================================
// HELPERS
// ============================================================
function serializeVal(val, key) {
  if (!(val instanceof Date)) return val;
  const tz = Session.getScriptTimeZone(), k = String(key||'').toLowerCase().trim();
  if (val.getFullYear() <= 1900) return Utilities.formatDate(val, tz, "hh:mm a");
  if (k === 'date')              return Utilities.formatDate(val, tz, "dd-MMM-yyyy");
  if (k === 'time')              return Utilities.formatDate(val, tz, "hh:mm a");
  return Utilities.formatDate(val, tz, "dd-MMM-yyyy hh:mm a");
}
function getColMap(headers) {
  const m = {};
  headers.forEach((h,i) => { if(h) m[String(h).trim().toLowerCase()] = i; });
  return m;
}
function extractFolderID(v) {
  if (!v) return null;
  const s = String(v).trim();
  const m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  const f = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (f) return f[1];
  return s;
}
function levenshtein(a, b) {
  const m=a.length, n=b.length, dp=[];
  for(let i=0;i<=m;i++){dp[i]=[i];for(let j=1;j<=n;j++)dp[i][j]=0;}
  for(let j=0;j<=n;j++)dp[0][j]=j;
  for(let i=1;i<=m;i++)for(let j=1;j<=n;j++)
    dp[i][j]=a[i-1]===b[j-1]?dp[i-1][j-1]:1+Math.min(dp[i-1][j],dp[i][j-1],dp[i-1][j-1]);
  return dp[m][n];
}
function verifyAdmin(params) {
  if (!params.adminToken) throw new Error("Unauthorized: no token");
  if (params.adminExpiry && new Date() > new Date(params.adminExpiry)) throw new Error("Session expired");
}
function verifySuperAdmin(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Admins");
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).trim()===params.adminUser){
      const role = String(data[i][2]).trim().toLowerCase();
      if(role !== "superadmin") throw new Error("Super Admin access required");
      return;
    }
  }
  throw new Error("User not found");
}
function nowFormatted() {
  const tz = Session.getScriptTimeZone(), now = new Date();
  return {
    date: Utilities.formatDate(now, tz, "dd-MMM-yyyy"),
    time: Utilities.formatDate(now, tz, "hh:mm a"),
    full: Utilities.formatDate(now, tz, "dd-MMM-yyyy hh:mm:ss")
  };
}

// ============================================================
// SETTINGS — Vertical format: Col A = key, Col B = value
// ============================================================
function getSettings() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const obj = {};
  data.forEach(r => { if (r[0]) obj[String(r[0]).trim()] = r[1]; });
  return obj;
}

function getPublicVisibility() {
  const s = getSettings();
  const isActive = (key) => String(s[key]||"ACTIVE").toUpperCase().trim() === "ACTIVE";
  return {
    showDonorList:       isActive("SHOW_DONOR_LIST"),
    showStatistics:      isActive("SHOW_STATISTICS"),
    showHomepageStats:   isActive("SHOW_HOMEPAGE_STATS"),
    showHomepageDonors:  isActive("SHOW_HOMEPAGE_DONORS"),
    showGallery:         isActive("SHOW_GALLERY"),
    showInviteCard:      isActive("SHOW_INVITE_CARD"),
    showPendingPayments: isActive("SHOW_PENDING_PAYMENTS"),
    showVerifiedPayments:isActive("SHOW_VERIFIED_PAYMENTS"),
    showRecentPayments:  isActive("SHOW_RECENT_PAYMENTS")
  };
}

function updateSettings(params) {
  verifySuperAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Settings");
  if (!sheet) throw new Error("Settings sheet not found");
  const data = sheet.getDataRange().getValues();
  const updates = JSON.parse(params.updates || '{}');
  Object.keys(updates).forEach(key => {
    let found = false;
    for(let i=0;i<data.length;i++){
      if(String(data[i][0]).trim()===key){
        const oldVal = data[i][1];
        sheet.getRange(i+1, 2).setValue(updates[key]);
        logAudit({adminUser:params.adminUser, module:"Settings", action:"Update",
          field:key, oldValue:oldVal, newValue:updates[key], reason:params.reason||""});
        found=true; break;
      }
    }
    if(!found) sheet.appendRow([key, updates[key]]);
  });
  return { result:"Saved" };
}

// ============================================================
// ADMIN LOGIN
// Admins sheet: Username(A) | Password(B) | Role(C) | AcessLevel(D) | Status(E) | Email(F) | CreatedAt(G) | LastLogin(H)
// ============================================================
function loginAdmin(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Admins");
  if (!sheet) return { success:false, error:"Admins sheet not found" };
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    const u=String(data[i][0]).trim(), p=String(data[i][1]).trim();
    const status = String(data[i][4]||"Active").trim();
    if(u===params.username && p===params.password){
      if(status.toLowerCase()==="inactive") return { success:false, error:"Account inactive" };
      const s = getSettings();
      const timeout = parseInt(s.SessionTimeoutMinutes)||30;
      const expiry = new Date(Date.now()+timeout*60*1000).toISOString();
      const token  = Utilities.getUuid();
      // Update LastLogin
      try{ sheet.getRange(i+1, 8).setValue(nowFormatted().full); }catch(e){}
      logActivity({adminUser:params.username, action:"Login", detail:"Successful login"});
      logAudit({adminUser:params.username, module:"Auth", action:"Login",
        field:"session", oldValue:"", newValue:"active", reason:"Login"});
      return {
        success:true,
        role:        data[i][2]||"admin",
        accessLevel: data[i][3]||"full",   // AcessLevel column
        email:       data[i][5]||"",
        token, expiry
      };
    }
  }
  logAudit({adminUser:params.username||"unknown", module:"Auth", action:"FailedLogin",
    field:"", oldValue:"", newValue:"", reason:"Wrong credentials"});
  return { success:false };
}

// ============================================================
// UTR VALIDATION & FRAUD DETECTION
// UTRBlacklist sheet: UTR(A) | AddedAt(B) | Reason(C)
// ============================================================
function isUTRBlacklisted(utr) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("UTRBlacklist");
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    for(let i=1;i<data.length;i++){
      if(String(data[i][0]).trim()===utr) return true;
    }
  } catch(e){}
  return false;
}

function addUTRBlacklist(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("UTRBlacklist");
  if (!sheet) throw new Error("UTRBlacklist sheet not found");
  const n = nowFormatted();
  sheet.appendRow([params.utr, n.full, params.reason||"Manually blacklisted by "+params.adminUser]);
  return { result:"Blacklisted" };
}

function getUTRBlacklist(params) {
  verifyAdmin(params);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("UTRBlacklist");
  if (!sheet) return { list:[] };
  const data = sheet.getDataRange().getValues();
  const list = [];
  for(let i=1;i<data.length;i++) if(data[i][0]) list.push({utr:data[i][0],addedAt:serializeVal(data[i][1],'date'),reason:data[i][2]});
  return { list };
}

function validateUTR(params) {
  const utr = String(params.utr||'').trim();
  if (!utr) return { valid:false, risk:"HIGH", score:100, flags:["Empty UTR"], block:true };

  const s = getSettings();
  const highT = parseInt(s.FRAUD_THRESHOLD_HIGH)||70;
  const medT  = parseInt(s.FRAUD_THRESHOLD_MEDIUM)||40;

  let score=0; const flags=[];

  // Check blacklist first
  if(isUTRBlacklisted(utr)) return { valid:false, risk:"HIGH", score:100, flags:["UTR is blacklisted"], block:true };

  // Format checks
  if(!/^\d+$/.test(utr))    { score+=35; flags.push("Non-numeric characters"); }
  if(utr.length<10)          { score+=30; flags.push("Too short (min 10 digits)"); }
  if(utr.length>22)          { score+=15; flags.push("Too long (max 22 digits)"); }

  // Pattern analysis
  if(/^(.)\1+$/.test(utr))  { score+=45; flags.push("All identical digits"); }
  const testVals=["123456789012","000000000000","111111111111","999999999999","123123123123"];
  if(testVals.includes(utr)) { score+=50; flags.push("Known test/fake value"); }
  let isSeq=true;
  for(let i=1;i<Math.min(utr.length,8);i++) if(parseInt(utr[i])-parseInt(utr[i-1])!==1){isSeq=false;break;}
  if(isSeq&&utr.length>=6)  { score+=25; flags.push("Sequential digits"); }

  // Check Payments sheet for duplicates & similarity
  try {
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet=ss.getSheetByName("Payments");
    const data=sheet.getDataRange().getValues();
    const col=getColMap(data[0]);
    const utrC=col["utr"]!==undefined?col["utr"]:7;
    const phoneC=col["phone number"]!==undefined?col["phone number"]:(col["phone"]!==undefined?col["phone"]:5);
    const recentUTRs=[];
    for(let i=Math.max(1,data.length-300);i<data.length;i++){
      const eu=String(data[i][utrC]||'').trim();
      if(!eu) continue;
      if(eu===utr) return { valid:false, risk:"HIGH", score:100, flags:["Exact duplicate UTR"], block:true };
      if(eu.length>=10&&utr.length>=10){
        const dist=levenshtein(utr,eu);
        if(dist<=1){ score+=50; flags.push("Nearly identical to existing UTR"); }
        else if(dist<=2){ score+=25; flags.push("Very similar to existing UTR"); }
      }
      // Same phone rapid submission
      if(params.phone&&String(data[i][phoneC]||'').trim()===String(params.phone).trim()) score+=20;
    }
  } catch(e){}

  score = Math.min(score, 100);
  const risk  = score>=highT?"HIGH":score>=medT?"MEDIUM":"LOW";
  const block = score>=highT;
  return { valid:!block, risk, score, flags, block };
}

// ============================================================
// PAYMENTS
// Payments sheet: RefID(A) | Date(B) | Time(C) | Full Name(D) |
//   Village(E) | Phone number(F) | Amount(G) | UTR(H) | Status(I) |
//   FraudScore(J) | RiskLevel(K) | ReviewFlag(L) | ShowPublic(M) |
//   Verified By(N) | VerifiedAt(O) | Notes(P)
// ============================================================
function insertPayment(params) {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Payments");
  const data=sheet.getDataRange().getValues();
  const col=getColMap(data[0]);
  // Phone column: "phone number" in your sheet
  const phoneC=col["phone number"]!==undefined?col["phone number"]:(col["phone"]!==undefined?col["phone"]:5);
  const utrC=col["utr"]!==undefined?col["utr"]:7;

  // Duplicate phone check
  for(let i=1;i<data.length;i++){
    if(String(data[i][phoneC]).trim()===String(params.phone).trim())
      return { result:"DuplicatePhone" };
  }

  // UTR validation
  const utrCheck=validateUTR({utr:params.utr, phone:params.phone});
  if(utrCheck.block) return { result:"DuplicateUTR", message:utrCheck.flags.join(", "), risk:utrCheck.risk };

  const n=nowFormatted();
  const reviewFlag=utrCheck.risk==="MEDIUM"?"Review":utrCheck.risk==="HIGH"?"HighRisk":"";
  const status=utrCheck.risk==="MEDIUM"?"Pending (Review)":"Pending";

  sheet.appendRow([
    params.refid,         // A: RefID
    n.date,               // B: Date
    n.time,               // C: Time
    params.name,          // D: Full Name
    params.village,       // E: Village
    params.phone,         // F: Phone number
    Number(params.amount),// G: Amount
    params.utr,           // H: UTR
    status,               // I: Status
    utrCheck.score,       // J: FraudScore
    utrCheck.risk,        // K: RiskLevel
    reviewFlag,           // L: ReviewFlag
    "Yes",                // M: ShowPublic
    "", "", ""            // N:VerifiedBy O:VerifiedAt P:Notes
  ]);

  try {
    const s=getSettings();
    if(s.OrganizerEmail){
      MailApp.sendEmail({
        to:String(s.OrganizerEmail),
        subject:(utrCheck.risk==="MEDIUM"?"⚠️[Review] ":"💰")+"New: "+params.name+" ₹"+params.amount,
        body:"Name: "+params.name+"\nPhone: "+params.phone+"\nAmount: ₹"+params.amount+
             "\nUTR: "+params.utr+"\nRisk: "+utrCheck.risk+
             (utrCheck.flags.length?"\nFlags: "+utrCheck.flags.join(", "):"")+
             "\nRef: "+params.refid+"\n"+n.date+" "+n.time
      });
    }
  } catch(e){}
  return { result:"Inserted", riskLevel:utrCheck.risk };
}

function getPublicStats() {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Payments");
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return {total:0,count:0,pending:0};
  const col=getColMap(data[0]);
  const aC=col["amount"]!==undefined?col["amount"]:6;
  const sC=col["status"]!==undefined?col["status"]:8;
  let total=0,count=0,pending=0;
  for(let i=1;i<data.length;i++){
    const st=String(data[i][sC]).trim(), amt=Number(data[i][aC])||0;
    if(st==="Verified"){total+=amt;count++;}
    if(st.startsWith("Pending")) pending++;
  }
  return {total,count,pending};
}

function checkStatus(params) {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Payments");
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return {found:false};
  const col=getColMap(data[0]);
  // Map to your actual column names
  const C={
    refid:   col["refid"]!==undefined?col["refid"]:0,
    date:    col["date"]!==undefined?col["date"]:1,
    time:    col["time"]!==undefined?col["time"]:2,
    name:    col["full name"]!==undefined?col["full name"]:(col["name"]!==undefined?col["name"]:3),
    village: col["village"]!==undefined?col["village"]:4,
    phone:   col["phone number"]!==undefined?col["phone number"]:(col["phone"]!==undefined?col["phone"]:5),
    amount:  col["amount"]!==undefined?col["amount"]:6,
    utr:     col["utr"]!==undefined?col["utr"]:7,
    status:  col["status"]!==undefined?col["status"]:8,
    risk:    col["risklevel"]!==undefined?col["risklevel"]:10
  };
  const type=params.searchType||'refid', val=String(params.searchVal||params.refid||'').trim();
  for(let i=1;i<data.length;i++){
    let match=false;
    if(type==='phone') match=String(data[i][C.phone]).trim()===val;
    else if(type==='utr') match=String(data[i][C.utr]).trim()===val;
    else match=String(data[i][C.refid]).trim().slice(-5)===val;
    if(match) return {
      found:true,
      refid:data[i][C.refid],
      date:serializeVal(data[i][C.date],'date'),
      time:serializeVal(data[i][C.time],'time'),
      name:data[i][C.name],
      village:data[i][C.village],
      phone:data[i][C.phone],
      amount:data[i][C.amount],
      utr:data[i][C.utr],
      status:data[i][C.status],
      riskLevel:data[i][C.risk]||""
    };
  }
  return {found:false};
}

function getPayments(params) {
  verifyAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Payments");
  const data=sheet.getDataRange().getValues();
  const headers=data[0]; const rows=[];
  for(let i=1;i<data.length;i++){
    const row={_row:i+1};
    headers.forEach((h,j)=>{ if(h) row[String(h).trim()]=serializeVal(data[i][j],h); });
    // Normalize for frontend
    row.Name     = row["Full Name"]    || row["Name"]    || "";
    row.Phone    = row["Phone number"] || row["Phone"]   || "";
    row.RefID    = row["RefID"]        || "";
    row.VerifiedBy = row["Verified By"]|| row["VerifiedBy"] || "";
    rows.push(row);
  }
  return {payments:rows};
}

function updatePayments(params) {
  verifyAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Payments");
  const data=sheet.getDataRange().getValues();
  const col=getColMap(data[0]);
  const n=nowFormatted();
  // Column indices (1-based for setRange)
  const stC  =(col["status"]!==undefined?col["status"]:8)+1;
  const vbC  =(col["verified by"]!==undefined?col["verified by"]:(col["verifiedby"]!==undefined?col["verifiedby"]:13))+1;
  const vaC  =(col["verifiedat"]!==undefined?col["verifiedat"]:14)+1;
  const updates=JSON.parse(params.updates);
  updates.forEach(u=>{
    const oldSt=sheet.getRange(u.row,stC).getValue();
    sheet.getRange(u.row,stC).setValue(u.status);
    sheet.getRange(u.row,vbC).setValue(params.adminUser||"admin");
    sheet.getRange(u.row,vaC).setValue(n.full);
    logAudit({adminUser:params.adminUser, module:"Payments", action:"StatusChange",
      field:"Status", oldValue:oldSt, newValue:u.status, reason:params.reason||""});
  });
  logActivity({adminUser:params.adminUser, action:"VerifyPayments", detail:updates.length+" records updated"});
  return {result:"Saved"};
}

function updatePublicDisplay(params) {
  verifyAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Payments");
  const data=sheet.getDataRange().getValues();
  const col=getColMap(data[0]);
  const spC=(col["showpublic"]!==undefined?col["showpublic"]:12)+1;
  sheet.getRange(parseInt(params.row),spC).setValue(params.showPublic);
  return {result:"Updated"};
}

function getPublicPayments() {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Payments");
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return {donors:[]};
  const col=getColMap(data[0]);
  const nC=col["full name"]!==undefined?col["full name"]:(col["name"]!==undefined?col["name"]:3);
  const aC=col["amount"]!==undefined?col["amount"]:6;
  const sC=col["status"]!==undefined?col["status"]:8;
  const spC=col["showpublic"]!==undefined?col["showpublic"]:12;
  const dC=col["date"]!==undefined?col["date"]:1;
  const donors=[];
  for(let i=1;i<data.length;i++){
    if(String(data[i][sC]).trim()==="Verified"&&String(data[i][spC]).trim()!=="No")
      donors.push({name:data[i][nC], amount:Number(data[i][aC])||0, date:serializeVal(data[i][dC],'date')});
  }
  donors.sort((a,b)=>b.amount-a.amount);
  return {donors};
}

// ============================================================
// COMPLAINTS
// Complaints sheet: ComplaintID(A) | Date(B) | Time(C) | Name(D) |
//   Village(E) | Phone(F) | Email(G) | Complaint(H) |
//   Attachment(I) | AttachmentURL(J) | AttachmentName(K) |
//   Status(L) | ReplyBy(M) | AdminReply(N) | RepliedAt(O) | Priority(P)
// ============================================================
function insertComplaint(params) {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Complaints");
  const n=nowFormatted();
  let fileUrl="", fileStatus="None";

  if(params.filedata && params.filename){
    try {
      const s=getSettings();
      const folderID=extractFolderID(s.COMPLAINT_UPLOAD_FOLDER_ID)||"1nMx6KmUbp0CZCmK5FDrXzvyiuqiOhdKH";
      // DriveApp requires drive scope — ensure script has 'https://www.googleapis.com/auth/drive' in appsscript.json
      const folder=DriveApp.getFolderById(folderID);
      const decoded=Utilities.base64Decode(params.filedata);
      const blob=Utilities.newBlob(decoded, params.filetype||"application/octet-stream", params.filename);
      const file=folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl="https://drive.google.com/file/d/"+file.getId()+"/view";
      fileStatus="Attached";
    } catch(e){
      fileUrl="Error: "+e.message;
      fileStatus="Error";
      // Don't block complaint submission on file error
    }
  }

  // Generate ComplaintID
  const cID="CP"+Date.now().toString().slice(-8);

  sheet.appendRow([
    cID,            // A: ComplaintID
    n.date,         // B: Date
    n.time,         // C: Time
    params.name,    // D: Name
    params.village, // E: Village
    params.phone,   // F: Phone
    params.email,   // G: Email
    params.complaint,// H: Complaint
    fileStatus,     // I: Attachment
    fileUrl,        // J: AttachmentURL
    params.filename||"",// K: AttachmentName
    "Open",         // L: Status
    "", "", "", ""  // M:ReplyBy N:AdminReply O:RepliedAt P:Priority
  ]);

  try {
    const s=getSettings();
    if(s.OrganizerEmail) MailApp.sendEmail({
      to:String(s.OrganizerEmail),
      subject:"📋 Complaint: "+params.name,
      body:"ID: "+cID+"\nName: "+params.name+"\nVillage: "+params.village+
           "\nPhone: "+params.phone+"\nComplaint:\n"+params.complaint+
           (fileUrl?"\nAttachment: "+fileUrl:"")
    });
  } catch(e){}
  return {result:"Inserted", complaintID:cID};
}

function getComplaints(params) {
  verifyAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Complaints");
  const data=sheet.getDataRange().getValues();
  const headers=data[0]; const rows=[];
  for(let i=1;i<data.length;i++){
    const row={_row:i+1};
    headers.forEach((h,j)=>{ if(h) row[String(h).trim()]=serializeVal(data[i][j],h); });
    rows.push(row);
  }
  return {complaints:rows};
}

function updateComplaint(params) {
  verifyAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Complaints");
  const data=sheet.getDataRange().getValues();
  const col=getColMap(data[0]);
  const n=nowFormatted();
  // Complaint columns (1-based)
  const stC  =(col["status"]!==undefined?col["status"]:11)+1;
  const rbC  =(col["replyby"]!==undefined?col["replyby"]:12)+1;
  const arC  =(col["adminreply"]!==undefined?col["adminreply"]:13)+1;
  const raC  =(col["repliedat"]!==undefined?col["repliedat"]:14)+1;
  const oldSt=sheet.getRange(parseInt(params.row),stC).getValue();
  sheet.getRange(parseInt(params.row),stC).setValue(params.status);
  sheet.getRange(parseInt(params.row),rbC).setValue(params.adminUser||"admin");
  sheet.getRange(parseInt(params.row),arC).setValue(params.reply);
  sheet.getRange(parseInt(params.row),raC).setValue(n.full);
  try {
    if(params.email&&params.reply){
      const s=getSettings();
      MailApp.sendEmail({to:params.email,
        subject:"Reply — "+(s.EventName||"Event"),
        body:"Dear "+params.name+",\n\nReply:\n"+params.reply+"\n\nStatus: "+params.status+"\n\nEvent Team"});
    }
  } catch(e){}
  logAudit({adminUser:params.adminUser, module:"Complaints", action:"Reply",
    field:"Status", oldValue:oldSt, newValue:params.status, reason:"Complaint reply"});
  logActivity({adminUser:params.adminUser, action:"ReplyComplaint", detail:"Replied to "+params.name});
  return {result:"Updated"};
}

// ============================================================
// VILLAGES
// Villages sheet: Village(A) | NormalizedName(B) | Count(C) | Status(D)
// ============================================================
function getVillageSuggestions() {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Villages");
  if (!sheet) return {villages:[]};
  const data=sheet.getDataRange().getValues();
  const villages=[];
  for(let i=1;i<data.length;i++){
    const v=String(data[i][0]||'').trim();
    const status=String(data[i][3]||'Active').trim();
    if(v&&status.toLowerCase()!=='inactive') villages.push(v);
  }
  return {villages:[...new Set(villages)].sort()};
}

function addVillageSuggestion(params) {
  verifyAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("Villages");
  if(!sheet) throw new Error("Villages sheet not found");
  const data=sheet.getDataRange().getValues();
  const n=nowFormatted();
  // Check if already exists, update Count
  for(let i=1;i<data.length;i++){
    if(String(data[i][0]).trim().toLowerCase()===params.village.toLowerCase()){
      const count=parseInt(data[i][2]||0)+1;
      sheet.getRange(i+1,3).setValue(count);
      return {result:"Updated"};
    }
  }
  sheet.appendRow([params.village, params.village.toLowerCase(), 1, "Active"]);
  return {result:"Added"};
}

// ============================================================
// GALLERY — from Drive folder (EVENT_GALLERY_FOLDER_ID setting)
// ============================================================
function getGalleryImages() {
  try {
    const s=getSettings();
    const folderURL=s.EVENT_GALLERY_FOLDER_ID||"";
    const folderID=extractFolderID(folderURL);
    if(!folderID) return {images:[], error:"Gallery folder not configured"};
    const folder=DriveApp.getFolderById(folderID);
    const files=folder.getFiles();
    const images=[];
    while(files.hasNext()){
      const f=files.next();
      if(f.getMimeType().startsWith("image/")){
        try{ f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW); }catch(e){}
images.push({
 id:f.getId(),
 name:f.getName(),
 url:"https://drive.google.com/uc?id="+f.getId(),
 thumb:"https://drive.google.com/thumbnail?id="+f.getId()+"&sz=w400",
 download:"https://drive.google.com/uc?export=download&id="+f.getId()
});
      }
    }
    return {images};
  } catch(e){ return {images:[],error:e.message}; }
}

// ============================================================
// ACTIVITY LOG
// ActivityLog sheet: RecordID(A) | Date(B) | Time(C) | AdminUser(D) |
//   Action(E) | Detail(F) | OldValue(G) | NewValue(H) |
//   Details(I) | Duration(J) | Browser(K) | Device(L) | LogoutType(M)
// ============================================================
function logActivity(params) {
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet=ss.getSheetByName("ActivityLog");
  if(!sheet){
    sheet=ss.insertSheet("ActivityLog");
    sheet.appendRow(["RecordID","Date","Time","AdminUser","Action","Detail",
      "OldValue","NewValue","Details","Duration","Browser","Device","LogoutType"]);
  }
  const n=nowFormatted();
  const recID="AL"+Date.now().toString().slice(-8);
  sheet.appendRow([
    recID, n.date, n.time,
    params.adminUser||"", params.action||"", params.detail||"",
    params.oldValue||"", params.newValue||"",
    params.details||"", "", "", "", ""
  ]);
  return {result:"Logged"};
}

function getActivity(params) {
  verifyAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("ActivityLog");
  if(!sheet) return {activities:[]};
  const data=sheet.getDataRange().getValues();
  const rows=[];
  for(let i=Math.max(1,data.length-50);i<data.length;i++){
    rows.push({
      date:serializeVal(data[i][1],'date'), time:serializeVal(data[i][2],'time'),
      user:data[i][3], action:data[i][4], detail:data[i][5]
    });
  }
  return {activities:rows.reverse()};
}

// ============================================================
// AUDIT LOG
// AuditLog sheet: Timestamp(A) | AdminUser(B) | Module(C) |
//   Action(D) | Field(E) | OldValue(F) | NewValue(G) | Reason(H)
// ============================================================
function logAudit(params) {
  try {
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet=ss.getSheetByName("AuditLog");
    if(!sheet){
      sheet=ss.insertSheet("AuditLog");
      sheet.appendRow(["Timestamp","AdminUser","Module","Action","Field","OldValue","NewValue","Reason"]);
    }
    const n=nowFormatted();
    sheet.appendRow([n.full, params.adminUser||"", params.module||"", params.action||"",
      params.field||"", params.oldValue||"", params.newValue||"", params.reason||""]);
  } catch(e){}
}

function getAuditLog(params) {
  verifySuperAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName("AuditLog");
  if(!sheet) return {logs:[]};
  const data=sheet.getDataRange().getValues();
  const logs=[];
  const limit=parseInt(params.limit)||100;
  for(let i=Math.max(1,data.length-limit);i<data.length;i++){
    logs.push({timestamp:String(data[i][0]),user:data[i][1],module:data[i][2],
      action:data[i][3],field:data[i][4],oldValue:data[i][5],newValue:data[i][6],reason:data[i][7]});
  }
  return {logs:logs.reverse()};
}

// ============================================================
// SHEET EDITOR (Super Admin Only)
// ============================================================
function getSheetsList(params) {
  verifySuperAdmin(params);
  return {sheets:SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().map(s=>s.getName())};
}
function getSheetData(params) {
  verifySuperAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName(params.sheetName);
  if(!sheet) return {error:"Sheet not found"};
  const data=sheet.getDataRange().getValues();
  return {data, rows:data.length, cols:data[0]?data[0].length:0, sheetName:params.sheetName};
}
function updateSheetCell(params) {
  verifySuperAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName(params.sheetName);
  if(!sheet) throw new Error("Sheet not found");
  const row=parseInt(params.row), col=parseInt(params.col);
  if(row<2) throw new Error("Cannot edit header row");
  const oldVal=sheet.getRange(row,col).getValue();
  sheet.getRange(row,col).setValue(params.value);
  logAudit({adminUser:params.adminUser, module:"Sheet:"+params.sheetName, action:"CellEdit",
    field:"R"+row+"C"+col, oldValue:oldVal, newValue:params.value, reason:params.reason||"Direct edit"});
  return {result:"Updated"};
}
function addSheetRow(params) {
  verifySuperAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName(params.sheetName);
  if(!sheet) throw new Error("Sheet not found");
  const rowData=JSON.parse(params.rowData||"[]");
  sheet.appendRow(rowData);
  logAudit({adminUser:params.adminUser, module:"Sheet:"+params.sheetName, action:"AddRow",
    field:"row", oldValue:"", newValue:JSON.stringify(rowData), reason:params.reason||"New row"});
  return {result:"Added", row:sheet.getLastRow()};
}
function deleteSheetRow(params) {
  verifySuperAdmin(params);
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=ss.getSheetByName(params.sheetName);
  if(!sheet) throw new Error("Sheet not found");
  const row=parseInt(params.row);
  if(row<2) throw new Error("Cannot delete header row");
  const oldData=sheet.getRange(row,1,1,sheet.getLastColumn()).getValues()[0];
  sheet.deleteRow(row);
  logAudit({adminUser:params.adminUser, module:"Sheet:"+params.sheetName, action:"DeleteRow",
    field:"row "+row, oldValue:JSON.stringify(oldData), newValue:"", reason:params.reason||"Deleted"});
  return {result:"Deleted"};
}
