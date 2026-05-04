/**
 * =========================================================================
 * REPORT SYSTEM BACKEND (Ultimate Merge: Fix Date + Base64 Image Support)
 * อัปเกรดระบบรองรับการดึงรูปภาพ และรองรับโหมด HK (Housekeeping)
 * =========================================================================
 */

function getReportData(params) {
  try {
    // 🎯 โหมด Routine และ Routine_HK จะดึงข้อมูลจากชีต Routine_Log เหมือนกัน
    const sheetName = params.reportType === 'Safety' ? 'Safety_Log' : 'Routine_Log';
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    
    if (!sheet) throw new Error("ไม่พบชีตชื่อ: " + sheetName);

    SpreadsheetApp.flush();

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const formulas = dataRange.getFormulas(); 
    
    if (data.length <= 1) throw new Error("ไม่มีข้อมูลในชีต");

    const headers = data.shift(); 
    formulas.shift(); 

    const safeHeaders = headers.map(h => h.toString().trim());
    
    const allRecords = data.map((row, rowIndex) => {
      const obj = {};
      safeHeaders.forEach((header, i) => {
        let val = row[i];
        if (val instanceof Date) {
          val = val.toISOString(); 
        }
        obj[header] = val;
      });
      
      obj['_formulas'] = formulas[rowIndex]; 
      return obj;
    });

    let filteredRecords = allRecords;

    if (params.sessionId) {
      const targetSession = params.sessionId.toString().trim();
      filteredRecords = filteredRecords.filter(row => {
        const rowSession = (row['Session ID'] || row['Session'] || '').toString().trim();
        return rowSession === targetSession;
      });
    } 
    else if (params.startDate && params.endDate) {
      const start = new Date(params.startDate).setHours(0,0,0,0);
      const end = new Date(params.endDate).setHours(23,59,59,999);
      
      filteredRecords = filteredRecords.filter(row => {
        const rowDateRaw = row['Timestamp'] || row['Date'] || row['วันที่']; 
        const rowDateMillis = parseCustomDate(rowDateRaw);
        if (!rowDateMillis) return false; 
        return rowDateMillis >= start && rowDateMillis <= end;
      });
    }

    let summary = {};
    // 🎯 เพิ่มทางแยกให้ระบบไปเรียกใช้ฟังก์ชันจัดการสถิติให้ถูกโหมด
    if (params.reportType === 'Routine') {
      summary = processRoutineData(filteredRecords);
    } else if (params.reportType === 'Routine_HK') {
      summary = processHKData(filteredRecords);
    } else if (params.reportType === 'Safety') {
      summary = processSafetyData(filteredRecords);
    }

    return {
      status: 'success',
      recordsFound: filteredRecords.length,
      records: filteredRecords,
      summary: summary
    };

  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

/**
 * -------------------------------------------------------------------------
 * ฟังก์ชันผู้ช่วย: ประมวลผลสถิติ Routine_Log (สำหรับ รปภ. / ช่าง)
 * -------------------------------------------------------------------------
 */
function processRoutineData(records) {
  let total = 0;
  let statusCount = { 'Normal': 0, 'Warning': 0, 'Critical': 0 };
  let issueList = []; 

  records.forEach(row => {
    const id = (row['Log ID'] || row['ID'] || row['Session ID'] || '-').toString().trim();
    const sessionId = (row['Session ID'] || '').toString().trim();
    
    // 🎯 ข้ามรายการของแม่บ้าน (HK) ทันที เพื่อไม่ให้ไปปนในรายงานของ รปภ./ช่าง
    if (id.startsWith('HK') || sessionId.startsWith('HK')) return;

    total++; // นับยอดตรวจเฉพาะของ Patrol

    const status = (row['Status'] || row['Final Status'] || '').toString().trim();
    const floor = (row['Location / Floor'] || row['Floor'] || '').toString().trim();
    const zone = (row['Location / Zone'] || row['Zone'] || '').toString().trim();
    const fullLocation = floor && zone ? `${floor} - ${zone}` : (floor || zone || '-');
    const remarks = row['Remarks'] || row['Topic'] || '-'; 
    const gmrl = row['Ref. GMRL'] || row['GMRL'] || '-';

    if (status.includes('🟢') || (status.includes('ปกติ') && !status.includes('ผิด'))) {
      statusCount['Normal']++;
    } 
    else if (status.includes('🟠') || status.includes('ซ่อม') || status.includes('แก้ไข')) {
      statusCount['Warning']++;
    } 
    else if (status.includes('🔴') || status.includes('ผิดปกติ') || status.includes('ชำรุด')) {
      statusCount['Critical']++;
    }

    // ดึงเฉพาะรายการที่มีปัญหา
    if (status.includes('🟠') || status.includes('🔴') || status.includes('ซ่อม') || status.includes('ผิดปกติ')) {
      const imgFormula = row['_formulas'] ? row['_formulas'][8] : ''; 
      const base64Image = getBase64FromHyperlinkFormula(imgFormula);

      // ในฟังก์ชัน processRoutineData() ให้แก้บล็อกการ push เข้า array เป็นแบบนี้:
      issueList.push({ 
        id: id, 
        inspector: row['Inspector'] || row['ผู้ตรวจ'] || '-', // 🎯 เพิ่มบรรทัดนี้
        location: fullLocation, 
        details: remarks, 
        gmrl: gmrl, 
        status: status,
        imageUrl: base64Image
      });
    }
  });

  return { totalInspected: total, statusCount: statusCount, issueList: issueList };
}

/**
 * -------------------------------------------------------------------------
 * ฟังก์ชันผู้ช่วย: ประมวลผลสถิติ Routine_Log (สำหรับ แม่บ้าน HK) - ✨ NEW!
 * -------------------------------------------------------------------------
 */
function processHKData(records) {
  let total = 0;
  let statusCount = { 'Normal': 0, 'Warning': 0, 'Critical': 0 };
  let issueList = []; 

  records.forEach(row => {
    const id = (row['Log ID'] || row['ID'] || row['Session ID'] || '-').toString().trim();
    const sessionId = (row['Session ID'] || '').toString().trim();
    
    // 🎯 คัดกรองเอาเฉพาะงานของแม่บ้าน (HK) เท่านั้น
    if (!id.startsWith('HK') && !sessionId.startsWith('HK')) return;

    total++;

    const status = (row['Status'] || row['Final Status'] || '').toString().trim();
    const floor = (row['Location / Floor'] || row['Floor'] || '').toString().trim();
    const zone = (row['Location / Zone'] || row['Zone'] || '').toString().trim();
    const fullLocation = floor && zone ? `${floor} - ${zone}` : (floor || zone || '-');
    const remarks = row['Remarks'] || row['Topic'] || '-'; 
    const gmrl = row['Ref. GMRL'] || row['GMRL'] || '-';

    // นับสถิติโดยอิงคำที่แม่บ้านน่าจะใช้
    if (status.includes('🟢') || status.includes('ปกติ') || status.includes('เสร็จ')) {
      statusCount['Normal']++;
    } 
    else if (status.includes('🟠') || status.includes('รอดำเนินการ') || status.includes('แจ้ง')) {
      statusCount['Warning']++;
    } 
    else if (status.includes('🔴') || status.includes('ผิดปกติ') || status.includes('อุปสรรค')) {
      statusCount['Critical']++;
    }

    // 🎯 งานแม่บ้าน: ดึง "ทุกรายการ" ลงตาราง KPI (เอาเงื่อนไข If คัดสถานะออกไปเลย)
    const imgFormula = row['_formulas'] ? row['_formulas'][8] : ''; 
    const base64Image = getBase64FromHyperlinkFormula(imgFormula);

    // ทำเหมือนกันในฟังก์ชัน processHKData() ให้แก้บล็อกการ push เป็นแบบนี้:
    issueList.push({ 
      id: id, 
      inspector: row['Inspector'] || row['ผู้ตรวจ'] || '-', // 🎯 เพิ่มบรรทัดนี้
      location: fullLocation, 
      details: remarks, 
      gmrl: gmrl, 
      status: status,
      imageUrl: base64Image
    });
  });

  return { totalInspected: total, statusCount: statusCount, issueList: issueList };
}

/**
 * -------------------------------------------------------------------------
 * ฟังก์ชันผู้ช่วย: ประมวลผลสถิติ Safety_Log
 * -------------------------------------------------------------------------
 */
function processSafetyData(records) {
  let total = records.length;
  let statusCount = { 'Normal': 0, 'Warning': 0, 'Critical': 0 };
  let failureCauses = {}; 
  let issueList = [];

  records.forEach(row => {
    const id = (row['Audit ID'] || row['ID'] || row['Session ID'] || '-').toString().trim();

    const status = (row['Final Status'] || row['Status'] || '').toString().trim();
    const auditDetails = (row['Audit Details'] || '').toString();
    const category = row['Asset Category'] || '-';
    const floor = (row['Location / Floor'] || row['Floor'] || '').toString().trim();
    const zone = (row['Location / Zone'] || row['Zone'] || '').toString().trim();
    const fullLocation = floor && zone ? `${floor} - ${zone}` : (floor || zone || '-');
    const image = row['Evidence Image'] || '-';

    if (status.includes('🟢') || (status.includes('ปกติ') && !status.includes('ผิด'))) {
      statusCount['Normal']++;
    } 
    else if (status.includes('🟡') || status.includes('🟠') || status.includes('ซ่อม') || status.includes('เฝ้าระวัง')) {
      statusCount['Warning']++;
    } 
    else if (status.includes('🔴') || status.includes('ผิดปกติ') || status.includes('ชำรุด')) {
      statusCount['Critical']++;
    }

    if (auditDetails.includes('❌') || auditDetails.includes('Fail')) {
      const lines = auditDetails.split('\n');
      lines.forEach(line => {
        if (line.includes('❌') || line.includes('Fail')) {
          const cause = line.replace('❌', '').replace(/Fail/gi, '').replace('|', '').trim();
          if(cause) {
            failureCauses[cause] = (failureCauses[cause] || 0) + 1;
          }
        }
      });
    }

    if (status.includes('🔴') || status.includes('ผิดปกติ') || status.includes('ชำรุด')) {
       const imgFormula = row['_formulas'] ? row['_formulas'][10] : '';
       const base64Image = getBase64FromHyperlinkFormula(imgFormula);

       issueList.push({
        id: id, 
        category: category,
        location: fullLocation,
        image: image,
        details: auditDetails,
        imageUrl: base64Image
      });
    }
  });

  return { totalAudited: total, statusCount: statusCount, failureCauses: failureCauses, issueList: issueList };
}

/**
 * -------------------------------------------------------------------------
 * ฟังก์ชันผู้ช่วย: แปลงวันที่ให้รองรับทุกรูปแบบ (รองรับ ISO String ด้วย)
 * -------------------------------------------------------------------------
 */
function parseCustomDate(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date) return dateValue.getTime();

  let fallbackDate = new Date(dateValue).getTime();
  if (!isNaN(fallbackDate)) return fallbackDate;

  let str = dateValue.toString().trim();
  let parts = str.split(' '); 
  let datePart = parts[0]; 
  let timePart = parts[1] || '00:00:00'; 
  let dateElements = datePart.split('/'); 
  
  if (dateElements.length === 3) {
    let day = parseInt(dateElements[0], 10);
    let month = parseInt(dateElements[1], 10) - 1; 
    let year = parseInt(dateElements[2], 10);

    let timeElements = timePart.split(':'); 
    let hours = parseInt(timeElements[0], 10) || 0;
    let minutes = parseInt(timeElements[1], 10) || 0;
    let seconds = parseInt(timeElements[2], 10) || 0;

    return new Date(year, month, day, hours, minutes, seconds).getTime();
  }
  
  return null;
}

/**
 * =========================================================================
 * 🎯 ฟังก์ชันช่วยเหลือ: สกัด ID รูปภาพจากสูตร HYPERLINK และแปลงเป็น Base64
 * =========================================================================
 */
function getBase64FromHyperlinkFormula(formulaString) {
  if (!formulaString || formulaString === "") return null;
  var idMatch = formulaString.match(/https:\/\/drive\.google\.com\/uc\?id=([a-zA-Z0-9-_]+)/);
  if (idMatch && idMatch[1]) {
    var fileId = idMatch[1]; 
    try {
      var file = DriveApp.getFileById(fileId);
      var blob = file.getBlob();
      var base64 = Utilities.base64Encode(blob.getBytes());
      var mimeType = blob.getContentType();
      return "data:" + mimeType + ";base64," + base64;
    } catch (e) {
      Logger.log("ดึงรูปภาพไม่ได้ (อาจไม่มีสิทธิ์ หรือไฟล์ใหญ่เกิน): " + e.message);
      return null; 
    }
  }
  return null; 
}
