// =========================================================================
// 📊 MODULE 6: DASHBOARD ANALYTICS ENGINE (J.A.R.V.I.S. PRO V.9 - APEX READY)
// =========================================================================

function getDashboardData(filterMonth, filterYear) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assetSheet = ss.getSheetByName('Asset_Master');
    const maintSheet = ss.getSheetByName('Maintenance_Log');
    const transferSheet = ss.getSheetByName('Transfer_Log'); 
    const borrowSheet = ss.getSheetByName('Borrow_Return_Log');
    
    let today = new Date();
    today.setHours(0,0,0,0);
    
    let targetMonth = (filterMonth !== undefined && filterMonth !== null && filterMonth !== "") ? parseInt(filterMonth) : today.getMonth();
    let targetYear = (filterYear !== undefined && filterYear !== null && filterYear !== "") ? parseInt(filterYear) : today.getFullYear();
    
    let dashData = {
      totalAssets: 0, totalValue: 0, maintenanceCost: 0,
      pmDueSoon: 0, pmOverdue: 0, borrowed: 0, overdueReturn: 0,
      transferInMonth: 0, transferOutMonth: 0,
      categoryChart: {}, techPerformance: {}
    };

    // 1. วิเคราะห์ Asset_Master
    if (assetSheet && assetSheet.getLastRow() > 1) {
      const data = assetSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0] && !data[i][1]) continue; 
        dashData.totalAssets++;
        let price = parseFloat(String(data[i][9]).replace(/,/g, '')) || 0;
        dashData.totalValue += price;
        let cat = data[i][1] || "อื่นๆ";
        dashData.categoryChart[cat] = (dashData.categoryChart[cat] || 0) + 1;
        
        let currentStatus = String(data[i][20] || "");
        if (currentStatus.includes("กำลังยืม")) dashData.borrowed++;
        
        // ⚡ J.A.R.V.I.S. FIX: นับ "โอนย้ายออก" จาก Asset_Master (คอลัมน์ U) เพื่อให้ได้สถานะ Real-time เป๊ะๆ
        if (currentStatus.includes("โอนย้ายออก")) dashData.transferOutMonth++;
      }
    }

    // 2. วิเคราะห์ Maintenance_Log
    if (maintSheet && maintSheet.getLastRow() > 1) {
      const data = maintSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        let mDate = parseThaiDate(data[i][1]);
        if (!mDate) continue;
        if (mDate.getMonth() === targetMonth && mDate.getFullYear() === targetYear) {
          dashData.maintenanceCost += parseFloat(String(data[i][5]).replace(/,/g, '')) || 0;
          let tech = data[i][7];
          if (tech) dashData.techPerformance[tech] = (dashData.techPerformance[tech] || 0) + 1;
        }
        if (data[i][6]) {
          let nextPm = parseThaiDate(data[i][6]);
          if (nextPm) {
            let diff = Math.ceil((nextPm - today) / (1000 * 60 * 60 * 24));
            if (diff >= 0 && diff <= 7) dashData.pmDueSoon++;
            else if (diff < 0) dashData.pmOverdue++;
          }
        }
      }
    }

    // 3. วิเคราะห์ Borrow_Return_Log
    if (borrowSheet && borrowSheet.getLastRow() > 1) {
      const bData = borrowSheet.getDataRange().getValues();
      for (let i = 1; i < bData.length; i++) {
        if (String(bData[i][8]).includes("กำลังยืม") && bData[i][6]) {
          let dueDate = parseThaiDate(bData[i][6]);
          if (dueDate && dueDate < today) dashData.overdueReturn++;
        }
      }
    }

    // 4. วิเคราะห์ Transfer_Log (นับเฉพาะ "รับโอนเข้า" ประจำเดือน)
    if (transferSheet && transferSheet.getLastRow() > 1) {
      const tData = transferSheet.getDataRange().getValues();
      for (let i = 1; i < tData.length; i++) {
        let tDate = parseThaiDate(tData[i][1]);
        if (tDate && tDate.getMonth() === targetMonth && tDate.getFullYear() === targetYear) {
          let type = String(tData[i][2] || "");
          // ⚡ J.A.R.V.I.S. FIX: ดักจับคำว่า 'เข้า' หรือ 'รับโอน'
          if (type.includes("เข้า") || type.includes("รับโอน")) dashData.transferInMonth++;
        }
      }
    }

    dashData.categoryChart = Object.fromEntries(Object.entries(dashData.categoryChart).sort((a,b)=>b[1]-a[1]).slice(0,5));
    dashData.techPerformance = Object.fromEntries(Object.entries(dashData.techPerformance).sort((a,b)=>b[1]-a[1]).slice(0,3));

    return { status: 'success', data: dashData };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function parseThaiDate(dateStr) {
  if (!dateStr) return null;
  let d;
  if (dateStr instanceof Date) { d = new Date(dateStr.getTime()); }
  else {
    let p = String(dateStr).split(' ')[0].split('/'); 
    if (p.length === 3) {
      let year = parseInt(p[2]);
      if (year > 2400) year -= 543;
      d = new Date(year, parseInt(p[1]) - 1, parseInt(p[0]));
    } else { d = new Date(dateStr); }
  }
  if (d && d.getFullYear() > 2400) d.setFullYear(d.getFullYear() - 543);
  return d;
}

/** 🔍 2. ระบบค้นหาประวัติรายตัว (V.13 - FIXED DATE) */
function getAssetHistoryLog(searchKey) {
  try {
    if (!searchKey) return { status: 'error', message: 'No search key' };
    let key = String(searchKey).trim().toLowerCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const maintSheet = ss.getSheetByName('Maintenance_Log');
    const assetSheet = ss.getSheetByName('Asset_Master');
    
    let assetFullName = "ไม่ทราบชื่อสินทรัพย์";
    let assetImgUrl = ""; 

    if (assetSheet) {
      const aData = assetSheet.getDataRange().getValues();
      for (let i = 1; i < aData.length; i++) {
        let aId = String(aData[i][0]).toLowerCase();
        let aName = String(aData[i][2]).toLowerCase();
        if (aId === key || aName.includes(key) || key.includes(aId)) {
          assetFullName = aData[i][2];
          assetImgUrl = String(aData[i][22] || "").trim(); 
          key = aId;
          break;
        }
      }
    }

    let records = [];
    let totalCost = 0;
    if (maintSheet) {
      const mData = maintSheet.getDataRange().getValues();
      for (let i = 1; i < mData.length; i++) {
        if (String(mData[i][2]).toLowerCase().includes(key)) {
          let cleanDate = parseThaiDate(mData[i][1]); 
          let dateStr = cleanDate ? Utilities.formatDate(cleanDate, "GMT+7", "dd/MM/yyyy") : String(mData[i][1]);
          let costVal = parseFloat(String(mData[i][5]).replace(/,/g, '')) || 0;
          totalCost += costVal;
          
          records.push({
            date: dateStr,
            type: String(mData[i][3] || "-"),
            details: String(mData[i][4] || "-"),
            cost: costVal,
            technician: String(mData[i][7] || "-")
          });
        }
      }
    }
    records.reverse(); 
    return { 
      status: 'success', 
      assetName: assetFullName, 
      totalCost: totalCost, 
      data: records,
      imageUrl: assetImgUrl };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

// ====================================================================================
// 🏢 MODULE GMRL: DASHBOARD & PENDING JOBS ENGINE (J.A.R.V.I.S. PRO - CSV MATCHED)
// ====================================================================================

// 1️⃣ ฟังก์ชันดึงข้อมูลสรุปภาพรวมสำหรับ Dashboard ฝั่ง GMRL (✨ อัปเกรด Pro Version)
function getGmrlDashboardData(filterMonth, filterYear) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gmrlLogSheet = ss.getSheetByName('Gmrl_Log');
    
    let today = new Date();
    let targetMonth = (filterMonth !== undefined && filterMonth !== null && filterMonth !== "") ? parseInt(filterMonth) : today.getMonth();
    let targetYear = (filterYear !== undefined && filterYear !== null && filterYear !== "") ? parseInt(filterYear) : today.getFullYear();
    
    let dashData = {
      pendingJobs: 0,
      gmrlCost: 0,
      // ⚡ อัปเกรด 1: เพิ่ม "ทีมช่าง T | H | A" รองรับ Auto-Rank 3 อันดับ
      techPerformance: { 
          "ช่างโจ": {count: 0, cost: 0}, 
          "ช่างอาร์ม": {count: 0, cost: 0},
          "ทีมช่าง T | H | A": {count: 0, cost: 0}
      },
      // ⚡ อัปเกรด 2: สร้างกล่องรับข้อมูลสำหรับ TOP 10 Job Type
      typePerformance: {} 
    };

    if (gmrlLogSheet && gmrlLogSheet.getLastRow() > 1) {
      const data = gmrlLogSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        let logId = String(data[i][0] || "").trim(); // คอลัมน์ A (0)
        if (!logId) continue;

        let startDate = data[i][1]; // คอลัมน์ B (1) - วันที่เริ่ม
        let endDate = data[i][2];   // คอลัมน์ C (2) - วันที่เสร็จ
        
        // ⚡ ถ้าระยะเวลาสิ้นสุด (End Date) เป็นช่องว่าง แปลว่า "งานค้าง"
        if (!endDate || String(endDate).trim() === "") {
           dashData.pendingJobs++;
        }
        
        let checkDate = parseThaiDate(endDate) ? parseThaiDate(endDate) : parseThaiDate(startDate); 

        if (checkDate && checkDate.getMonth() === targetMonth && checkDate.getFullYear() === targetYear) {
          let cost = parseFloat(String(data[i][7]).replace(/,/g, '')) || 0; // คอลัมน์ H (7) - ค่าใช้จ่าย
          dashData.gmrlCost += cost;
          
          let tech = String(data[i][8] || "").trim(); // คอลัมน์ I (8) - ชื่อช่าง
          if (tech.includes("โจ")) {
              dashData.techPerformance["ช่างโจ"].count++;
              dashData.techPerformance["ช่างโจ"].cost += cost;
          } else if (tech.includes("อาร์ม")) {
              dashData.techPerformance["ช่างอาร์ม"].count++;
              dashData.techPerformance["ช่างอาร์ม"].cost += cost;
          } else {
              // ถ้าช่างไม่ระบุชื่อ หรือเป็นชื่ออื่น ให้ลงยอดที่ "ทีมช่าง T | H | A"
              dashData.techPerformance["ทีมช่าง T | H | A"].count++;
              dashData.techPerformance["ทีมช่าง T | H | A"].cost += cost;
          }

          // ⚡ อัปเกรด 3: ระบบเก็บข้อมูลนับยอด "ประเภทงาน" (Job Type)
          // ⚠️ คำเตือนจากจาร์วิส: ผมอ้างอิง "ประเภทงาน" ไว้ที่ คอลัมน์ E (Index 4) 
          // หากเจ้านายเก็บข้อมูลประเภทงานไว้คอลัมน์อื่น ให้เปลี่ยนตัวเลข 4 ตรง data[i][4] ได้เลยครับ
          let jobType = String(data[i][4] || "งานซ่อมทั่วไป").trim(); 
          if (jobType !== "") {
              if (!dashData.typePerformance[jobType]) {
                  dashData.typePerformance[jobType] = { count: 0, cost: 0 };
              }
              dashData.typePerformance[jobType].count++;
              dashData.typePerformance[jobType].cost += cost;
          }
          
        }
      }
    }
    return { status: 'success', ...dashData };
  } catch (e) { 
      return { status: 'error', message: e.toString() }; 
  }
}

// 2️⃣ ฟังก์ชันดึงรายการงานค้างไปโชว์ใน Modal (พร้อม Location Badge 📍)
function getPendingGmrlJobs() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Gmrl_Log');
        if (!sheet) return { success: false, message: 'ไม่พบชีต Gmrl_Log' };

        const data = sheet.getDataRange().getValues();
        let pendingList = [];

        for (let i = data.length - 1; i >= 1; i--) {
            let logId = String(data[i][0] || "").trim(); // คอลัมน์ A (0)
            if (!logId) continue;
            
            let endDate = data[i][2]; // คอลัมน์ C (2) - วันที่เสร็จ
            
            // ⚡ FIX: งานค้างคือ งานที่ยังไม่ได้กรอกวันที่เสร็จสิ้น
            if (!endDate || String(endDate).trim() === "") {
                pendingList.push({
                    logId: logId,                               // คอลัมน์ A (0)
                    category: String(data[i][3] || "-"),        // คอลัมน์ D (3) - หมวดหมู่งาน
                    assetName: "งานซ่อมบำรุงทั่วไป",            
                    location: String(data[i][5] || ""),         // ✨ คอลัมน์ F (5) - จุดดำเนินการ (เอาไปทำ Badge)
                    type: String(data[i][4] || "-"),            // คอลัมน์ E (4) - ประเภทงาน
                    details: String(data[i][6] || "-"),         // คอลัมน์ G (6) - รายละเอียด
                    rowIdx: i + 1
                });
            }
        }
        return { success: true, data: pendingList };
    } catch (e) { return { success: false, message: e.toString() }; }
}

// 3️⃣ ฟังก์ชันบันทึกการปิดงานจาก Dashboard
function closeGmrlJob(logId, formattedDate) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Gmrl_Log');
        if (!sheet) return { success: false, message: 'ไม่พบชีต Gmrl_Log' };

        const data = sheet.getDataRange().getValues();
        let targetRow = -1;

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === String(logId).trim()) {
                targetRow = i + 1;
                break;
            }
        }

        if (targetRow === -1) return { success: false, message: 'ไม่พบรหัสงานนี้ในระบบ' };

        // ⚡ FIX: บันทึกวันที่ปิดงานลงไปใน คอลัมน์ C (คอลัมน์ที่ 3) ซึ่งก็คือ End Date
        sheet.getRange(targetRow, 3).setValue(formattedDate);
        
        return { success: true };
    } catch (e) { return { success: false, message: e.toString() }; }
}
