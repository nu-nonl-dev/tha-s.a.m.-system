// =========================================================================
// 🛠️ MODULE: GENERAL MAINTENANCE REPAIR LOG (GMRL)
// =========================================================================

function saveGmrlData(formObj) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Gmrl_Log');
    if (!sheet) throw new Error("ไม่พบชีต Gmrl_Log กรุณาสร้างชีตก่อนครับ");

    // 1. สร้างรหัส Transaction ID อัตโนมัติ (Gmrl-0001)
    let nextId = "Gmrl-0001";
    if (sheet.getLastRow() > 1) {
      const lastIdStr = sheet.getRange(sheet.getLastRow(), 1).getValue().toString();
      const match = lastIdStr.match(/Gmrl-(\d+)/);
      if (match) { nextId = "Gmrl-" + ("0000" + (parseInt(match[1], 10) + 1)).slice(-4); }
    }

    // 2. ตั้งค่า Folder ID ตามที่เจ้านายกำหนด
    const GMRL_IMG_FOLDER_ID = '1tqTO8B4fPbAGuXThMjJYGRm7o2khJg1N';
    const GMRL_DOC_FOLDER_ID = '1V6Cpv0gxngvHLo2nERiSmrokoWH9hhJD';
    
    // (ไม่ต้องสร้าง ts ซ้อน เพราะใน processAndUploadFile มีการใส่ Timestamp ให้แล้ว)

    // 3. จัดการอัปโหลดไฟล์ โยนเข้าฟังก์ชัน Upload_Sys.gs (ถ้ามี)
    let imgUrl = "ไม่ได้แนบไฟล์"; 
    if (formObj.gmrlImage) {
      // ใส่ Prefix เป็น GMRL_IMG เพื่อให้ชื่อไฟล์เป็นระเบียบ เช่น GMRL_IMG_20260412_...
      imgUrl = processAndUploadFile(formObj.gmrlImage, GMRL_IMG_FOLDER_ID, "GMRL_IMG");
    }
    
    let docUrl = "ไม่ได้แนบไฟล์"; 
    if (formObj.gmrlDocument) {
      // ใส่ Prefix เป็น GMRL_DOC
      docUrl = processAndUploadFile(formObj.gmrlDocument, GMRL_DOC_FOLDER_ID, "GMRL_DOC");
    }

    // 4. วันที่ (รับมาจากหน้าเว็บเป็น DD/MM/YYYY สำเร็จรูปอยู่แล้ว โยนลงชีตได้เลย)
    let sheetStartDate = formObj.startDate || "";
    let sheetEndDate = formObj.endDate || "";

    // 5. เตรียมข้อมูล 13 คอลัมน์ (A ถึง M)
    const newRow = [
      nextId,                     // A: Transaction ID
      sheetStartDate,             // B: Start Date
      sheetEndDate,               // C: End Date
      formObj.category,           // D: Category
      formObj.type,               // E: Type
      formObj.location,           // F: Location
      formObj.details,            // G: Details
      formObj.cost !== "" ? formObj.cost : "", // H: Cost (ถ้าว่างให้เป็นค่าว่าง ไม่พิมพ์ 0)
      formObj.technician,         // I: Technician
      formObj.issuer,             // J: Issuer/Receiver
      formObj.remarks,            // K: Remarks
      imgUrl,                     // L: Image Link
      docUrl                      // M: Document Link
    ];
    
    // 6. บันทึกลงชีต
    sheet.appendRow(newRow);
    const targetRow = sheet.getLastRow();

    // 7. แปลงลิงก์ไฟล์ให้เป็น HYPERLINK กดดูง่ายๆ
    if (imgUrl !== "ไม่ได้แนบไฟล์" && !imgUrl.startsWith("Error")) {
      sheet.getRange(targetRow, 12).setFormula(`=HYPERLINK("${imgUrl}", "ดูรูปภาพ")`);
    }
    if (docUrl !== "ไม่ได้แนบไฟล์" && !docUrl.startsWith("Error")) {
      sheet.getRange(targetRow, 13).setFormula(`=HYPERLINK("${docUrl}", "ดูเอกสาร")`);
    }

    return { success: true, message: "บันทึกงานซ่อมทั่วไปเรียบร้อย รหัส: " + nextId };
  } catch (error) { 
    return { success: false, message: error.message }; 
  }
}
