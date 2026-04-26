// =========================================================
// 🚶‍♂️ GS: ROUTINE PATROL (เซฟลงชีต + อัปโหลดรูประบบบีบอัด + ตั้งชื่อไฟล์อัจฉริยะ)
// =========================================================
function processRoutineUpload(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Routine_Log'); 
    
    if (!sheet) {
      sheet = ss.insertSheet('Routine_Log');
      sheet.appendRow(["Log ID", "Session ID", "Timestamp", "Inspector", "Floor", "Zone", "Topic", "Status", "Image Link", "Ref. GMRL"]);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#d9ead3");
    }

    const ROUTINE_FOLDER_ID = '1hxh5bhWR7WNeLOlELf1rHbo6GspT8sAr';
    
    // 🌟 1. ดัดแปลงชื่อไฟล์ก่อนส่งให้ Upload_Sys.gs
    // โดยเอา ชื่อชั้น (Floor) มาตั้งเป็นชื่อไฟล์แทนชื่อเดิมของรูป
    // ผมแถมโซน (Zone) ต่อท้ายให้ด้วยนิดนึง เพื่อให้ระบุตำแหน่งได้แม่นยำขึ้นครับ
    const floorName = payload.floor ? payload.floor : "UnknowFloor";
    const zoneName = payload.zone ? payload.zone : "";
    const customFileName = `${floorName}_${zoneName}.jpg`;
    
    const fileDataObj = {
      base64: payload.fileData,
      mimeType: payload.mimeType,
      name: customFileName // หลอกระบบ Upload_Sys ว่าไฟล์นี้ชื่อตามจุดที่เช็ค
    };
    
    let imgUrl = "ไม่ได้แนบไฟล์";
    if (payload.fileData) {
      // 🌟 2. กำหนด Prefix เป็น "RTP" ส่งเข้า Upload_Sys
      // Upload_Sys จะผสมคำออกมาเป็น: RTP_20260426_160849_M_Flr_Zone_A.jpg ให้เองเป๊ะๆ ครับ!
      imgUrl = processAndUploadFile(fileDataObj, ROUTINE_FOLDER_ID, "RTP");
      
      if (imgUrl.includes("Error")) {
        throw new Error("การอัปโหลดรูปภาพล้มเหลว: " + imgUrl);
      }
    }

    // รันรหัส Log ID (RT-XXXX)
    let logId = "RT-0001";
    if (sheet.getLastRow() > 1) {
      const lastIdStr = sheet.getRange(sheet.getLastRow(), 1).getValue().toString();
      const match = lastIdStr.match(/RT-(\d+)/);
      if (match) { logId = "RT-" + ("0000" + (parseInt(match[1], 10) + 1)).slice(-4); }
    }

    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    
    // บันทึกลงชีต
    sheet.appendRow([
      logId,
      payload.sessionId,
      timestamp,
      payload.inspector,
      payload.floor,
      payload.zone,
      payload.topic,
      payload.status,
      imgUrl,
      ""
    ]);
    
    // ทำลิงก์ HYPERLINK สวยๆ
    const targetRow = sheet.getLastRow();
    if (imgUrl !== "ไม่ได้แนบไฟล์" && !imgUrl.startsWith("Error")) {
      sheet.getRange(targetRow, 9).setFormula(`=HYPERLINK("${imgUrl}", "ดูรูปภาพจุดตรวจ")`);
    }
    
    return { success: true, logId: logId };
    
  } catch (error) {
    throw new Error(error.toString());
  }
}
