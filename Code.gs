const SHEET_ID = 'YOUR_MAIN_SHEET_ID_HERE'; 
const SHEET_NAME = 'DisputeData';
const SOURCE_SHEET_ID = 'YOUR_SOURCE_SHEET_ID_HERE'; 
const SOURCE_SHEET_NAME = 'Source - CS'; // เปลี่ยนจากชื่อเฉพาะของบริษัท

function setSheetFormatting(sheet) {
  sheet.getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm:ss");
  sheet.getRange("B:B").setNumberFormat("@");
  sheet.getRange("C:C").setNumberFormat("dd/mm/yyyy hh:mm:ss");
}

function isFormOpen() {
  const now = new Date();
  const day = now.getDay();
  const hours = now.getHours();
  // ปิดรับฟอร์มวันอังคาร ตั้งแต่ 02:00 น.
  if (day === 2 && hours >= 2) return false;
  // เปิดรับวัน อาทิตย์(0), จันทร์(1), อังคาร(2), ศุกร์(5), เสาร์(6)
  return [0, 1, 2, 5, 6].includes(day);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('QA Dispute System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function submitData(formData) {
  try {
    if (!isFormOpen()) {
      throw new Error('❌ ไม่สามารถกรอกฟอร์มได้ในขณะนี้ อยู่นอกเหนือเวลาการกรอกฟอร์ม ❌');
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
   
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = ["เวลาที่ Submit ฟอร์ม", "Email Address", "วันที่ Dispute ฟอร์ม", "QA Week", "📄 รายละเอียดที่ต้องการ Dispute *", "🤔 เหตุผลที่ QA หักคะแนน * ", "🔗 แนบลิงก์ Ticket *", "🔗 ลิงก์ Ref Slack/Knowledge Base", "📝 หมายเหตุเพิ่มเติม ", "สถานะ 📃", "Sup/ Senior Remark",  "QA Remark"];
      sheet.appendRow(headers);
      setSheetFormatting(sheet);
    }
    const columnAValues = sheet.getRange("A:A").getValues();
    let lastRowWithContent = 0;
    for (let i = columnAValues.length - 1; i >= 0; i--) {
        if (columnAValues[i][0] !== "") {
            lastRowWithContent = i + 1;
            break;
        }
    }
    if (lastRowWithContent === 0) {
        lastRowWithContent = sheet.getLastRow();
    }
    const targetRow = lastRowWithContent + 1;

    const userEmail = Session.getActiveUser().getEmail();
    const disputeDateObject = new Date(formData.disputeDate);
    const newRowData = [
      new Date(), userEmail, disputeDateObject,
      formData.qaWeek || '-', formData.details, formData.qaReason,
      formData.ticketLink, formData.refLink, formData.notes,
      '',
      '-'
    ];
    sheet.getRange(targetRow, 1, 1, newRowData.length).setValues([newRowData]);
    setSheetFormatting(sheet);
    
    return 'บันทึกข้อมูลสำเร็จ!';
  } catch (error) {
    throw new Error(error.message);
  }
}

function updateDisputeStatus() {
  const sourceSheet = SpreadsheetApp.openById(SOURCE_SHEET_ID).getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) return;
  
  const sourceData = sourceSheet.getRange("A2:O").getValues();
  const statusMap = new Map();
  for (const row of sourceData) {
    const ticket = row[0];
    const status = row[14];
    if (ticket) statusMap.set(ticket.toString().trim(), status);
  }
  
  const mainSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const mainData = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 12).getValues();
  const newStatuses = [];
  
  for (const r of mainData) {
    const ticketLink = r[6];
    const supRemark = r[10];
    const currentStatus = r[9];
    let newStatus = currentStatus;
    
    if (typeof supRemark === 'string' && supRemark.includes('ไม่') && supRemark.includes('Dispute')) {
      newStatus = 'Sup/Senior ไม่ Dispute';
    } else {
      const match = ticketLink?.match(/\/(\d+)$/);
      if (match) {
        const tk = 'TK' + match[1];
        const statusFromCS = statusMap.get(tk);
        if (statusFromCS === 'Yes') newStatus = 'Dispute สำเร็จ';
        else if (statusFromCS === 'No') newStatus = 'Dispute ไม่สำเร็จ';
      }
    }
    newStatuses.push([newStatus]);
  }
  mainSheet.getRange(2, 10, newStatuses.length, 1).setValues(newStatuses);
  Logger.log('อัปเดตสถานะเรียบร้อย');
}

function getStatusData(filters) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];

    let weekFilter = filters && filters.week ? filters.week.trim() : '';
    let emailFilter = filters && filters.email ? filters.email.trim() : '';
   
    let query = "SELECT * WHERE A IS NOT NULL";
    if (weekFilter) {
      query += ` AND lower(D) = lower('${weekFilter.replace(/'/g, "\\'")}')`;
    }
    if (emailFilter) {
      query += ` AND B matches '(?i)${emailFilter.replace(/'/g, "\\'")}'`;
    }
   
    if (!weekFilter && !emailFilter) {
      query += " ORDER BY A DESC";
    } else {
      query += " ORDER BY A DESC LIMIT 100";
    }

    const queryEncoded = encodeURIComponent(query);
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?sheet=${SHEET_NAME}&tq=${queryEncoded}`;
   
    const options = { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } };
    const response = UrlFetchApp.fetch(url, options);
    const text = response.getContentText();
   
    const jsonString = text.substring(text.indexOf('(') + 1, text.lastIndexOf(')'));
    const json = JSON.parse(jsonString);

    if (json.table.rows.length === 0) return [];

    const data = json.table.rows.map(r => {
      const row = r.c;
      return {
        submitTimestamp: row[0] ? row[0].f : '',
        email: row[1] ? row[1].v : '',
        disputeDate: row[2] ? row[2].f : '',
        qaWeek: row[3] ? row[3].v : '',
        details: row[4] ? row[4].v : '',
        qaReason: row[5] ? row[5].v : '',
        ticketLink: row[6] ? row[6].v : '',
        refLink: row[7] ? row[7].v : '',
        notes: row[8] ? row[8].v : '',
        status: row[9] ? row[9].v : '',
        remark: row[10] ? row[10].v : '',
        columnL: row[11] ? row[11].v : '' 
      };
    });
    return data;
  } catch (error) {
    throw new Error('เกิดข้อผิดพลาดในการดึงข้อมูล: ' + error.message);
  }
}

function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}
