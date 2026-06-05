
cat > /home/claude/EventPay/config.js << 'ENDOFFILE'
const APP_CONFIG = {
  SCRIPT_URL: "https://script.google.com/macros/s/AKfycbzvgfN2UJ8v1sTFXmBMgFCxf8NzEOu9-VLueR9TjcbOwkg7Ghb2wYz4zHhAbw6GMV7D8g/exec",
  UPI_ID: "9014844036@superyes",
  EVENT_NAME: "Wedding of ram & Sita",
  ORG_NAME: "Event Organisation",
  CURRENCY: "INR",
  MIN_AMOUNT: 50,
  MAX_AMOUNT: 0,
  THEME_COLOR: "#6366f1",
  VERSION: "3"
};
window.APP_CONFIG = APP_CONFIG;
ENDOFFILE
echo "config.js written"
Output

config.js written
