// ============================================================
// CONFIG.JS — Edit ONLY this file with your settings
// ============================================================

const APP_CONFIG = {

  // STEP 1: Paste your Google Apps Script URL here after deploying
  SCRIPT_URL: "https://script.google.com/macros/s/YOUR_SCRIPT_ID_HERE/exec",

  // STEP 2: Your UPI ID (e.g. "yourname@paytm" or "9876543210@upi")
  UPI_ID: "yourname@upi",

  // STEP 3: Event title shown on all pages (also editable in Google Sheets Settings)
  EVENT_NAME: "Event Contribution",

  // STEP 4: Your organization/event name for receipts
  ORG_NAME: "Event Organisation",

  // STEP 5: Currency (INR for India)
  CURRENCY: "INR",

  // STEP 6: Minimum contribution amount
  MIN_AMOUNT: 50,

  // STEP 7: Maximum contribution amount (0 = no limit)
  MAX_AMOUNT: 0,

  // STEP 8: Theme color (hex) — change to match your event
  THEME_COLOR: "#6366f1",

  // App version
  VERSION: "2.0"

};

// DO NOT EDIT BELOW THIS LINE
window.APP_CONFIG = APP_CONFIG;
