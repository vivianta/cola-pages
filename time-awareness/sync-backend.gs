// ============================================================
// 时间知觉 — Google Sheets 同步后端
// ============================================================
// 部署步骤：
// 1. Google Sheets → 扩展程序 → Apps Script
// 2. 粘贴全部代码，保存
// 3. 部署 → 新部署 → 网页应用
// 4. 执行身份：我，访问权限：任何人
// 5. 复制新 URL，粘贴到设置页
// ============================================================

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange('A1').getValue();
  return ContentService.createTextOutput(data || '{"sessions":[]}')
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // Accept text/plain to avoid CORS preflight
  const body = e.postData.contents;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  try {
    JSON.parse(body); // validate
    sheet.getRange('A1').setValue(body);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Invalid JSON' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
