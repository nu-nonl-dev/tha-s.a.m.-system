// =========================================================================
// 🛠️ MODULE 3: MAINTENANCE LOG (DATABASE) & SMART PM API (FIXED DATE FORMAT)
// =========================================================================

function saveMaintenanceData(formObj) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Maintenance_Log');
    if (!sheet) throw new Error("ไม่พบชีต Maintenance_Log");

    const techVal = formObj.mtTechnicianSelect === 'other' ? formObj.newMtTechnician : formObj.mtTechnicianSelect;
    
    // จัดการอัปโหลดไฟล์
    let mtImgUrl = "ไม่ได้แนบไฟล์"; 
    if (formObj.mtImage) mtImgUrl = processAndUploadFile(formObj.mtImage, IMAGE_FOLDER_ID, "MTIMG");

    // รันเลข ID อัตโนมัติ (MT-XXXX)
    let nextId = "MT-0001";
    if (sheet.getLastRow() > 1) {
      const lastIdStr = sheet.getRange(sheet.getLastRow(), 1).getValue().toString();
      const match = lastIdStr.match(/MT-(\d+)/);
      if (match) { nextId = "MT-" + ("0000" + (parseInt(match[1], 10) + 1)).slice(-4); }
    }

    // ⚡ จุดแก้ไข: แปลงรูปแบบวันที่จาก YYYY-MM-DD เป็น DD/MM/YYYY ก่อนบันทึก
    const formattedDate = Utilities.formatDate(new Date(formObj.date), "GMT+7", "dd/MM/yyyy");
    
    // แปลงวัน PM ถัดไป (ถ้ามี)
    let formattedNextPm = "";
    if (formObj.nextPmDate) {
      formattedNextPm = Utilities.formatDate(new Date(formObj.nextPmDate), "GMT+7", "dd/MM/yyyy");
    }

    const newRow = [
      nextId, 
      formattedDate, // คอลัมน์ B: Date (DD/MM/YYYY)
      formObj.assetId, 
      formObj.type, 
      formObj.details,
      formObj.cost, 
      formattedNextPm, // วัน PM ถัดไป (DD/MM/YYYY)
      techVal, 
      mtImgUrl 
    ];
    
    sheet.appendRow(newRow);
    
    // ใส่สูตร IMAGE ถ้ามีรูปภาพ
    if (mtImgUrl !== "ไม่ได้แนบไฟล์" && !mtImgUrl.startsWith("Error")) {
       sheet.getRange(sheet.getLastRow(), 9).setFormula(`=IMAGE("${mtImgUrl}")`); 
    }

    return { success: true, message: "บันทึกประวัติซ่อมบำรุงเรียบร้อย รหัส: " + nextId };
  } catch (error) { return { success: false, message: error.message }; }
}

// 🌟 NEW API: ฟังก์ชันนักสืบ หาประวัติ PM เพื่อส่งไปโชว์ที่ Smart Widget
function getAssetPMInfo(assetId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assetSheet = ss.getSheetByName('Asset_Master');
    const mtSheet = ss.getSheetByName('Maintenance_Log');
    
    if (!assetSheet) throw new Error("ไม่พบชีต Asset_Master");
    
    // 1. ค้นหาระยะรอบเวลา PM จากคอลัมน์ AC (Index 28)
    const assetData = assetSheet.getDataRange().getValues();
    let pmInterval = null;
    let foundAsset = false;
    
    for (let i = 1; i < assetData.length; i++) {
      if (assetData[i][0] === assetId) {
        foundAsset = true;
        pmInterval = assetData[i][28]; 
        break;
      }
    }
    
    if (!foundAsset) return { status: 'NOT_FOUND' };
    
    // 2. ค้นหาวันที่ทำ PM ล่าสุดจาก Maintenance_Log
    let lastPMDate = null;
    if (mtSheet && mtSheet.getLastRow() > 1) {
      const mtData = mtSheet.getDataRange().getValues();
      for (let i = mtData.length - 1; i >= 1; i--) {
        if (mtData[i][2] === assetId && mtData[i][3].toString().includes("PM")) { 
          lastPMDate = mtData[i][1]; 
          break;
        }
      }
    }
    
    // 3. คำนวณวัน PM ถัดไป และจัดรูปแบบการแสดงผล
    let nextPMStr = "";
    let lastPMStr = "ยังไม่มีประวัติ";
    let status = lastPMDate ? "OK" : "NEW";

    // ⚡ จุดแก้ไข: ปรับการแสดงผลวันที่ให้เป็น DD/MM/YYYY สำหรับ Widget
    if (pmInterval && !isNaN(pmInterval) && pmInterval > 0) {
        let baseDate = lastPMDate ? new Date(lastPMDate) : new Date();
        baseDate.setMonth(baseDate.getMonth() + parseInt(pmInterval, 10));
        nextPMStr = Utilities.formatDate(baseDate, "GMT+7", "dd/MM/yyyy");
    }
    
    if(lastPMDate) {
       lastPMStr = Utilities.formatDate(new Date(lastPMDate), "GMT+7", "dd/MM/yyyy");
    }

    return {
      status: status,
      interval: pmInterval,
      lastPM: lastPMStr,
      nextPM: nextPMStr
    };

  } catch (e) { return { status: 'ERROR', message: e.message }; }
}
