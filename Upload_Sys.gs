// =========================================================================
// 📦 MODULE 1: UPLOAD SYSTEM
// =========================================================================
function processAndUploadFile(fileData, gdriveId, prefix) {
  try {
    if (!fileData || !fileData.base64) return "Error: ไม่ได้รับข้อมูลไฟล์";
    const decodedFile = Utilities.base64Decode(fileData.base64);
    let blob = Utilities.newBlob(decodedFile, fileData.mimeType, "temp_file");
    let lastDot = fileData.name ? fileData.name.lastIndexOf('.') : -1;
    let origName = lastDot !== -1 ? fileData.name.substring(0, lastDot) : (fileData.name || "File");
    origName = origName.replace(/[^\w\u0E00-\u0E7F-]/g, "_").substring(0, 20); 
    let finalExt = lastDot !== -1 ? fileData.name.substring(lastDot) : (fileData.mimeType.includes("pdf") ? ".pdf" : ".jpg");

    if (fileData.mimeType.includes("image")) {
      const tempFile = DriveApp.getFolderById(gdriveId).createFile(blob);
      try {
        let file = Drive.Files.get(tempFile.getId(), { fields: "thumbnailLink" });
        if (!file || !file.thumbnailLink) { Utilities.sleep(1500); file = Drive.Files.get(tempFile.getId(), { fields: "thumbnailLink" }); }
        if (file && file.thumbnailLink) {
          const thumbUrl = file.thumbnailLink.replace(/=s\d+/, "=s800");
          blob = UrlFetchApp.fetch(thumbUrl).getBlob().setContentType("image/jpeg");
          finalExt = ".jpg"; 
        }
      } catch (e) {} finally { tempFile.setTrashed(true); }
    }

    const fileName = `${prefix}_${Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss")}_${origName}${finalExt}`;
    const finalFile = DriveApp.getFolderById(gdriveId).createFile(Utilities.newBlob(blob.getBytes(), fileData.mimeType.includes("image") ? "image/jpeg" : fileData.mimeType, fileName));
    try { finalFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return 'https://drive.google.com/uc?id=' + finalFile.getId();
  } catch (error) { return "Error: " + error.message; }
}
