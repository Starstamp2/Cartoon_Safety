// ============================================================
//  Safety Quiz — Google Apps Script Backend
//  วิธีใช้:
//  1. เปิด script.google.com → New Project
//  2. วางโค้ดนี้ทั้งหมด → บันทึก
//  3. Deploy → New deployment → Web app
//     - Execute as: Me
//     - Who has access: Anyone
//  4. คัดลอก Web App URL ไปใส่ใน safety.html
// ============================================================

const SHEET_NAME = 'QuizResults';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ลำดับ', 'ชื่อ-นามสกุล', 'แผนก', 'ภาษา', 'คะแนน', 'เปอร์เซ็นต์', 'สถานะ', 'วันที่', 'Timestamp']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#0f4c81').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'save') {
      const sheet = getSheet();
      const lastRow = sheet.getLastRow();
      const rowNum = lastRow; // หัวตาราง = แถว 1 ดังนั้น rowNum = lastRow (ไม่นับหัว)
      sheet.appendRow([
        rowNum,
        data.name,
        data.dept,
        data.lang,
        data.score,
        data.pct + '%',
        data.passed ? 'ผ่าน' : 'ไม่ผ่าน',
        data.date,
        data.timestamp
      ]);
      return jsonResponse({ status: 'ok', message: 'บันทึกสำเร็จ' });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'getAll') {
      const sheet = getSheet();
      const rows = sheet.getDataRange().getValues();
      if (rows.length <= 1) return jsonResponse({ status: 'ok', data: [] });

      const headers = rows[0];
      const data = rows.slice(1).map(row => ({
        name:      row[1],
        dept:      row[2],
        lang:      row[3],
        score:     Number(row[4]),
        pct:       Number(String(row[5]).replace('%', '')),
        passed:    row[6] === 'ผ่าน',
        date:      row[7],
        timestamp: row[8]
      }));
      return jsonResponse({ status: 'ok', data });
    }

    if (action === 'ping') {
      return jsonResponse({ status: 'ok', message: 'pong' });
    }

    return jsonResponse({ status: 'error', message: 'Unknown action' });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
