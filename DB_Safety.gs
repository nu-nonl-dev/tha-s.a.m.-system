// =========================================================================
// 🛡️ MODULE 6: SAFETY AUDIT (DATABASE V.398 - RE-POSITIONED A-L)
// =========================================================================

function saveSafetyAuditData(formObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const safetySheet = ss.getSheetByName('Safety_Log');
    const gmrlSheet = ss.getSheetByName('Gmrl_Log');
    
    if (!safetySheet) throw new Error("ไม่พบชีต Safety_Log");

    // 1. สร้างรหัส Audit ID (SA-XXXX)
    let auditId = "SA-0001";
    if (safetySheet.getLastRow() > 1) {
      const lastIdStr = safetySheet.getRange(safetySheet.getLastRow(), 1).getValue().toString();
      const match = lastIdStr.match(/SA-(\d+)/);
      if (match) { 
        auditId = "SA-" + ("0000" + (parseInt(match[1], 10) + 1)).slice(-4); 
      }
    }

    // 2. วันที่และรูปภาพ (🌟 วิ่งผ่าน Upload_Sys.gs ของเจ้านายโดยตรง!)
    const formattedDate = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    let imgUrl = "ไม่ได้แนบไฟล์";
    const SAFETY_IMG_FOLDER_ID = '1GGwWcLTm7EQzbbcqSeK6l7aLWTz7aHVV'; 
    if (formObj.sfImage) {
      // เรียกใช้ฟังก์ชันจาก Upload_Sys.gs ที่มีอยู่แล้วในโปรเจกต์
      imgUrl = processAndUploadFile(formObj.sfImage, SAFETY_IMG_FOLDER_ID, "SAF." + auditId);
    }

    // 3. เตรียมสถานะและ Action
    const statusMap = { "1": "🟢 ปกติ / พร้อมใช้งาน", "2": "🟡 เฝ้าระวัง", "3": "🟠 ควรรีบแก้ไข", "4": "🔴 ชำรุดหนัก", "5": "🔵 จัดซื้อใหม่" };
    const finalStatusText = statusMap[formObj.severity] || statusMap["1"];
    let actionTaken = formObj.severity === "1" ? "🟢 ปกติ" : "รอการจัดการ";

    // 4. 🚀 [GMRL HANDOFF] ส่งไม้ต่อแจ้งซ่อม (ถ้าติ๊กเลือก + ส้ม/แดง)
    if (formObj.createGmrl && (formObj.severity === "3" || formObj.severity === "4")) {
      if (gmrlSheet) {
        let gmrlId = "Gmrl-0001";
        if (gmrlSheet.getLastRow() > 1) {
          const lastGmrlStr = gmrlSheet.getRange(gmrlSheet.getLastRow(), 1).getValue().toString();
          const matchGmrl = lastGmrlStr.match(/Gmrl-(\d+)/);
          if (matchGmrl) { gmrlId = "Gmrl-" + ("0000" + (parseInt(matchGmrl[1], 10) + 1)).slice(-4); }
        }

        // วางลงชีต Gmrl_Log (ตามโครงสร้างมาตรฐาน)
        const gmrlRow = [
          gmrlId, formattedDate, "", "งานความปลอดภัย (Safety)", 
          formObj.category, formObj.floor + " " + formObj.zone, 
          `แจ้งจากระบบ Safety (${auditId})\nอุปกรณ์: ${formObj.assetName}\nหมายเหตุ: ${formObj.remarks}`, 
          "", "", formObj.inspector, formObj.checklistData, imgUrl, "ไม่ได้แนบไฟล์"
        ];
        gmrlSheet.appendRow(gmrlRow);
        
        // ทำ Link รูปในหน้า GMRL (คอลัมน์ L = 12)
        if (imgUrl !== "ไม่ได้แนบไฟล์" && !imgUrl.startsWith("Error")) {
          gmrlSheet.getRange(gmrlSheet.getLastRow(), 12).setFormula(`=HYPERLINK("${imgUrl}", "ดูรูปภาพ")`);
        }
        actionTaken = `เปิดใบแจ้งซ่อมแล้ว (${gmrlId})`;
      }
    }

    // 5. 🌟 [POSITIONING A-L] จัดลำดับลง Safety_Log
    const newSafetyRow = [
      auditId,               // A: Audit ID
      formattedDate,         // B: Timestamp
      formObj.inspector,     // C: Inspector
      formObj.floor,         // D: Floor 
      formObj.zone,          // E: Zone 
      formObj.assetName,     // F: Asset ID & Name 
      formObj.category,      // G: Category 
      formObj.checklistData, // H: Audit Details 
      finalStatusText,       // I: Status
      formObj.remarks,       // J: Remarks
      imgUrl,                // K: Image (จะใส่เป็น URL ก่อน)
      actionTaken            // L: Action
    ];
    
    safetySheet.appendRow(newSafetyRow);
    const targetRow = safetySheet.getLastRow();

    // 6. 🌟 เปลี่ยน URL ดิบ ให้กลายเป็นปุ่ม HYPERLINK "ดูรูปภาพชำรุด" ในคอลัมน์ K (11)
    if (imgUrl !== "ไม่ได้แนบไฟล์" && !imgUrl.startsWith("Error")) {
      safetySheet.getRange(targetRow, 11).setFormula(`=HYPERLINK("${imgUrl}", "ดูรูปภาพชำรุด")`);
    }

    return { 
      success: true, 
      message: actionTaken.includes("เปิดใบแจ้งซ่อม") ? "บันทึกสำเร็จ! และ " + actionTaken : "บันทึกผลการตรวจสอบเรียบร้อย!",
      isGmrlCreated: actionTaken.includes("เปิดใบแจ้งซ่อม")
    };

  } catch (error) { 
    return { success: false, message: error.message }; 
  }
}
