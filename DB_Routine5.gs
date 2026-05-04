// =========================================================
// 🚶‍♂️ GS: ROUTINE & HK PATROL (ฉบับรองรับ Multi-Mode + Perfect Upload)
// =========================================================

// 📦 1. MODULE: ระบบอัปโหลดความเร็วสูง (ใช้ระบบ Thumbnail ลดขนาดภาพ)
function processAndUploadFile(fileData, gdriveId, prefix) {
  try {
    if (!fileData || !fileData.base64) return "Error: ไม่ได้รับข้อมูลไฟล์";
    const decodedFile = Utilities.base64Decode(fileData.base64);
    let blob = Utilities.newBlob(decodedFile, fileData.mimeType, "temp_file");
    
    let lastDot = fileData.name ? fileData.name.lastIndexOf('.') : -1;
    let origName = lastDot !== -1 ? fileData.name.substring(0, lastDot) : (fileData.name || "File");
    origName = origName.replace(/[^\w\u0E00-\u0E7F-]/g, "_").substring(0, 20); 
    let finalExt = lastDot !== -1 ? fileData.name.substring(lastDot) : (fileData.mimeType.includes("pdf") ? ".pdf" : ".jpg");

    // ถ้าเป็นรูปภาพ ใช้ระบบบีบอัดผ่าน Thumbnail เพื่อความรวดเร็ว
    if (fileData.mimeType.includes("image")) {
      const tempFile = DriveApp.getFolderById(gdriveId).createFile(blob);
      try {
        let file = Drive.Files.get(tempFile.getId(), { fields: "thumbnailLink" });
        if (!file || !file.thumbnailLink) { 
          Utilities.sleep(1000); 
          file = Drive.Files.get(tempFile.getId(), { fields: "thumbnailLink" }); 
        }
        if (file && file.thumbnailLink) {
          const thumbUrl = file.thumbnailLink.replace(/=s\d+/, "=s800"); // บีบอัดเหลือ 800px
          blob = UrlFetchApp.fetch(thumbUrl).getBlob().setContentType("image/jpeg");
          finalExt = ".jpg"; 
        }
      } catch (e) {
        console.log("Thumbnail Error: " + e.message);
      } finally {
        tempFile.setTrashed(true); // ลบไฟล์ชั่วคราวทิ้งทันที
      }
    }

    const fileName = `${prefix}_${Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss")}_${origName}${finalExt}`;
    const finalFile = DriveApp.getFolderById(gdriveId).createFile(Utilities.newBlob(blob.getBytes(), fileData.mimeType.includes("image") ? "image/jpeg" : fileData.mimeType, fileName));
    try { finalFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    
    return 'https://drive.google.com/uc?id=' + finalFile.getId();
  } catch (error) { 
    return "Error: " + error.message; 
  }
}

// 🚀 2. ฟังก์ชันหลัก: บันทึกข้อมูลทีละข้อ และส่งไม้ต่อให้ GMRL
function processRoutineUpload(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Routine_Log'); 
    
    // ตรวจสอบ/สร้างหัวตาราง (12 คอลัมน์ A-L)
    if (!sheet) {
      sheet = ss.insertSheet('Routine_Log');
      sheet.appendRow(["Log ID", "Session ID", "Timestamp", "Inspector", "Floor", "Zone", "Topic", "Status", "Image Link", "Ref. GMRL", "Remarks", "Final Status"]);
      sheet.getRange("A1:L1").setFontWeight("bold").setBackground("#d9ead3");
    }

    const ROUTINE_FOLDER_ID = '1hxh5bhWR7WNeLOlELf1rHbo6GspT8sAr';
    
    // 🎯 [ระบบ Multi-Mode] ตรวจสอบว่าเป็นงาน Patrol หรือ HK จาก Session ID
    const isHK = payload.sessionId && String(payload.sessionId).startsWith('HK-');
    const idPrefix = isHK ? "HK" : "RT";
    const filePrefix = isHK ? "HKP" : "RTP";
    const sourceName = isHK ? "Housekeeping (HK)" : "Routine Patrol";
    
    // จัดการชื่อไฟล์รูปภาพให้สอดคล้องกับโหมด
    const floorName = payload.floor ? String(payload.floor).replace(/[^\w\sก-๙]/gi, '_') : "UnknownFloor";
    const zoneName = payload.zone ? String(payload.zone).replace(/[^\w\sก-๙]/gi, '_') : "";
    const customFileName = filePrefix + "_" + floorName + "_" + zoneName + ".jpg"; 
    
    let imgUrl = "ไม่ได้แนบไฟล์";
    
    // 🌟 ระบบอัปโหลดรูปความเร็วสูง
    if (payload.fileData) {
      const fileDataObj = {
        base64: payload.fileData,
        mimeType: payload.mimeType,
        name: customFileName
      };
      
      // เรียกใช้ MODULE 1
      imgUrl = processAndUploadFile(fileDataObj, ROUTINE_FOLDER_ID, filePrefix);
      
      if (imgUrl.includes("Error")) {
        throw new Error(imgUrl);
      }
    }

    // 🎯 [ระบบรันเลขแยกอิสระ] รันรหัส RT-XXXX หรือ HK-XXXX โดยไม่ให้ตีกัน
    let logId = idPrefix + "-0001";
    if (sheet.getLastRow() > 1) {
      // ดึงข้อมูล ID ทั้งหมดมาตรวจสอบหาเลขล่าสุดของ Prefix นั้นๆ
      const allIds = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat().filter(String);
      const idRegex = new RegExp(idPrefix + "-(\\d+)");
      let maxId = 0;
      
      // ค้นหาย้อนกลับเพื่อหา ID ล่าสุดที่ตรงกับหมวดหมู่
      for (let i = allIds.length - 1; i >= 0; i--) {
        const match = allIds[i].match(idRegex);
        if (match) {
          maxId = parseInt(match[1], 10);
          break; // เจอตัวล่าสุดแล้วหยุดค้นหา
        }
      }
      
      if (maxId > 0) {
        logId = idPrefix + "-" + ("0000" + (maxId + 1)).slice(-4);
      }
    }

    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    const dateOnly = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    
    let gmrlIdRef = "";
    
    // 🚀 [GMRL HANDOFF] สร้างใบแจ้งซ่อม GMRL จริงๆ
    if (payload.createGmrl && (payload.status === "🟠 แจ้งซ่อม/แก้ไข" || payload.status === "🔴 ผิดปกติ")) {
      let gmrlSheet = ss.getSheetByName('Gmrl_Log');
      if (gmrlSheet) {
        let gmrlId = "Gmrl-0001";
        if (gmrlSheet.getLastRow() > 1) {
          const lastGmrlStr = gmrlSheet.getRange(gmrlSheet.getLastRow(), 1).getValue().toString();
          const matchGmrl = lastGmrlStr.match(/Gmrl-(\d+)/);
          if (matchGmrl) { gmrlId = "Gmrl-" + ("0000" + (parseInt(matchGmrl[1], 10) + 1)).slice(-4); }
        }

        // บันทึกลง Gmrl_Log โดยปรับเปลี่ยน sourceName ให้ตรงกับผู้ส่ง
        const gmrlRow = [
          gmrlId, dateOnly, "", "งานอาคารสถานที่ (Facility)", 
          sourceName, String(payload.floor) + " " + String(payload.zone), 
          "แจ้งจากระบบ " + sourceName + " (" + logId + ")\nรายละเอียด: " + payload.topic, 
          "", "", payload.inspector, payload.remarks || "", imgUrl, "ไม่ได้แนบไฟล์"
        ];
        gmrlSheet.appendRow(gmrlRow);
        
        // ทำลิงก์รูปในหน้า GMRL (คอลัมน์ L = 12)
        if (imgUrl !== "ไม่ได้แนบไฟล์") {
          gmrlSheet.getRange(gmrlSheet.getLastRow(), 12).setFormula('=HYPERLINK("' + imgUrl + '", "ดูรูปภาพ")');
        }
        gmrlIdRef = gmrlId; 
      }
    }

    // บันทึกลง Routine_Log
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
      gmrlIdRef, 
      payload.remarks || "", 
      "⌛ กำลังดำเนินการ"
    ]);
    
    // ทำลิงก์รูปในหน้า Routine (คอลัมน์ I = 9)
    if (imgUrl !== "ไม่ได้แนบไฟล์") {
      sheet.getRange(sheet.getLastRow(), 9).setFormula('=HYPERLINK("' + imgUrl + '", "ดูรูปภาพจุดตรวจ")');
    }
    
    return { success: true, logId: logId, gmrlId: gmrlIdRef };
    
  } catch (error) {
    throw new Error(error.toString());
  }
}

// ✅ 3. ฟังก์ชันพิเศษ: สแตมป์ปิดจบทั้ง Session ในคอลัมน์ L (ปรับคำให้เป็นกลางใช้ได้ทุกแผนก)
function finishRoutineSessionInSheet(sessionId) {
  if (!sessionId) return false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Routine_Log');
  if (!sheet) return false;
  
  const data = sheet.getDataRange().getValues();
  let updated = false;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(sessionId)) { 
      sheet.getRange(i + 1, 12).setValue("🟢 ปิดจบรอบการทำงานสำเร็จ"); 
      updated = true;
    }
  }
  return updated;
}

// =========================================================
// 🧠 ฟังก์ชันดูดคลังคำศัพท์จากชีต (ทำ Auto-suggest)
// =========================================================
function getRoutineTopics() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Routine_Log');
    if (!sheet) return [];
    
    // ดึงข้อมูลคอลัมน์ G (Topic) ตั้งแต่แถวที่ 2 ลงมา
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    const data = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
    
    // กรองเอาเฉพาะข้อความที่ไม่ว่างเปล่า และตัดคำที่ซ้ำกันออก (Unique)
    const uniqueTopics = [...new Set(data.map(row => row[0]).filter(topic => topic.toString().trim() !== ""))];
    
    return uniqueTopics;
  } catch (e) {
    return [];
  }
}

// =========================================================
// 🚀 เช็คว่ามีงานค้าง "⌛ กำลังดำเนินการ" แยกตามโหมดที่กำลังเปิด
// =========================================================
function checkPendingRoutineSession(frontendMode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Routine_Log');
  if (!sheet) return { hasPending: false };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { hasPending: false };
  
  // 🎯 กำหนดคำนำหน้า (Prefix) เพื่อใช้ค้นหาตามโหมด
  // ถ้าเป็นโหมด HK ให้หาคำว่า "HK-" ถ้าเป็นโหมดอื่น (Patrol) ให้หา "RT-"
  let prefix = (frontendMode === 'HK') ? 'HK-' : 'RT-';
  
  // ค้นหาจากแถวล่าสุดย้อนขึ้นไป
  for (let i = data.length - 1; i >= 1; i--) {
    let status = data[i][11]; // คอลัมน์ L (Index 11) สถานะ
    let sessionId = String(data[i][1]).trim(); // คอลัมน์ B (Index 1) รหัสอ้างอิง
    
    // 🎯 เช็ค 2 เงื่อนไข: สถานะต้องกำลังดำเนินการ "และ" รหัสต้องตรงกับโหมดที่เปิดอยู่
    if (status === "⌛ กำลังดำเนินการ" && sessionId.startsWith(prefix)) {
      return { hasPending: true, sessionId: sessionId, mode: frontendMode };
    }
  }
  
  // ถ้าหาจนจบแล้วไม่เจอเลย
  return { hasPending: false };
}

// บังคับเปลี่ยนสถานะงานที่ค้างให้เป็นปิดจบ (ใช้ตอนที่ผู้ใช้กด ล้างข้อมูล)
function forceClosePendingSessions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Routine_Log');
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  let updated = false;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][11] === "⌛ กำลังดำเนินการ") {
      sheet.getRange(i + 1, 12).setValue("🟢 ปิดจบรอบการทำงาน (อัตโนมัติ)");
      updated = true;
    }
  }
  return updated;
}

// =========================================================
// 🚀 ฟังก์ชันดึงรายการงานเก่าใน Session ที่ค้างอยู่มาแสดง (อัปเดตระบบสแกน)
// =========================================================
function getExistingRoutineItems(sessionId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Routine_Log');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let existingItems = [];
  
  // แปลง sessionId เป็นตัวอักษรเพื่อความชัวร์ (เช่น "HK-260503-1042")
  const targetSession = String(sessionId).trim();
  
  for (let i = 1; i < data.length; i++) {
    let rowId = String(data[i][1]).trim(); // คอลัมน์ B (Index 1) ที่เก็บรหัส
    
    // 🎯 แก้ไขจุดนี้: ใช้ .includes() เพื่อให้มันครอบคลุมรหัสที่มีเลขต่อท้าย (เช่น HK-...-1)
    // และนี่จะเป็นการกรองโหมด (RT/HK) ไปในตัวด้วยครับ!
    if (rowId.includes(targetSession)) {
      existingItems.push({
        timestamp: data[i][0] ? Utilities.formatDate(new Date(data[i][0]), "GMT+7", "dd/MM/yyyy - HH:mm:ss") : "",
        inspector: data[i][2] || "-",
        floor: data[i][4] || "-",
        zone: data[i][5] || "-",
        topic: data[i][6] || "-",
        status: data[i][7] || "🟢 ปกติ"
      });
    }
  }
  
  // สลับให้รายการล่าสุดอยู่ด้านบน
  return existingItems.reverse(); 
}
