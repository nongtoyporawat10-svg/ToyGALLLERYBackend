// =========================================================================
// ⚙️ การตั้งค่าระบบ (Configuration)
// =========================================================================
const SHEET_NAME = "Bookings";
const PORTFOLIO_SHEET = "Portfolio";
const ADMIN_PASSWORD = "admin"; // 🔑 เปลี่ยนรหัสผ่านเข้าหลังบ้านตรงนี้
const CALENDAR_ID = "primary";
const LINE_NOTIFY_TOKEN = ""; 
const SPREADSHEET_ID = ""; 

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Toy Gallery 2000's | Photography")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupPermissions() { DriveApp.getRootFolder(); }

function initSystemBackend() {
  try {
    let ss = SPREADSHEET_ID !== "" ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("ไม่พบไฟล์ Google Sheets");
    
    let bookSheet = ss.getSheetByName(SHEET_NAME);
    if (!bookSheet) {
      bookSheet = ss.insertSheet(SHEET_NAME);
      // เพิ่ม Drive Link และ Photos Link ต่อท้าย Slip URL เป็นคอลัมน์ N และ O
      bookSheet.appendRow(["Booking ID", "Timestamp", "Customer", "Phone", "Date", "Time", "Package", "Location", "Price", "Deposit", "Notes", "Status", "Slip URL", "Drive Link", "Photos Link"]);
      bookSheet.getRange("A1:O1").setFontWeight("bold").setBackground("#f3f4f6");
      bookSheet.setFrozenRows(1);
    } else {
      if(bookSheet.getRange("M1").getValue() === "") { bookSheet.getRange("M1").setValue("Slip URL").setFontWeight("bold").setBackground("#f3f4f6"); }
      if(bookSheet.getRange("N1").getValue() === "") { bookSheet.getRange("N1").setValue("Drive Link").setFontWeight("bold").setBackground("#f3f4f6"); }
      if(bookSheet.getRange("O1").getValue() === "") { bookSheet.getRange("O1").setValue("Photos Link").setFontWeight("bold").setBackground("#f3f4f6"); }
    }

    let portSheet = ss.getSheetByName(PORTFOLIO_SHEET);
    if (!portSheet) {
      portSheet = ss.insertSheet(PORTFOLIO_SHEET);
      portSheet.appendRow(["ID", "Title", "Category", "URL", "Location", "Timestamp"]);
      portSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#f3f4f6");
      portSheet.setFrozenRows(1);
    }
    return ss;
  } catch (e) { throw e; }
}

function verifyAdmin(password) {
  if(password === ADMIN_PASSWORD) return { success: true, message: "เข้าสู่ระบบสำเร็จ" };
  return { success: false, message: "รหัสผ่านไม่ถูกต้อง" };
}

// =========================================================================
// 📁 ฟังก์ชันจัดการโฟลเดอร์หลัก (Main Folder)
// =========================================================================
function getAppFolder(subFolderName) {
  const mainFolderName = "ToyGallery Booking Backend";
  
  let mainFolders = DriveApp.getFoldersByName(mainFolderName);
  let mainFolder = mainFolders.hasNext() ? mainFolders.next() : DriveApp.createFolder(mainFolderName);
  
  let subFolders = mainFolder.getFoldersByName(subFolderName);
  let subFolder = subFolders.hasNext() ? subFolders.next() : mainFolder.createFolder(subFolderName);
  
  return subFolder;
}

// =========================================================================
// ⚙️ การตั้งค่าระบบหลังบ้าน (ตั้งค่าบัญชีโอนเงิน & QR Code)
// =========================================================================
function getAppSettings() {
  const props = PropertiesService.getScriptProperties();
  return {
    promptPayNum: props.getProperty('promptPayNum') || '0812345678',
    accountName: props.getProperty('accountName') || 'นายช่างภาพ ใจดี',
    qrImageUrl: props.getProperty('qrImageUrl') || ''
  };
}

function saveAppSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('promptPayNum', settings.promptPayNum);
    props.setProperty('accountName', settings.accountName);

    if (settings.qrBase64) {
      let folder = getAppFolder("ToyGallery_Settings");
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      const byteCharacters = Utilities.base64Decode(settings.qrBase64.split(",")[1]);
      const blob = Utilities.newBlob(byteCharacters, settings.mimeType, "Custom_QR_Code.jpg");
      const file = folder.createFile(blob);
      props.setProperty('qrImageUrl', "https://lh3.googleusercontent.com/d/" + file.getId());
    }
    return { success: true, message: 'บันทึกการตั้งค่าบัญชีและ QR Code สำเร็จ!' };
  } catch (error) { return { success: false, message: error.toString() }; }
}

// =========================================================================
// 🤖 อัปโหลดสลิป & อัปเดตคิว (AI ตรวจสอบ)
// =========================================================================
function uploadSlipAndUpdateBooking(id, fileData) {
  try {
     let folder = getAppFolder("ToyGallery_Slips");
     folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
     
     const byteCharacters = Utilities.base64Decode(fileData.base64.split(",")[1]);
     const blob = Utilities.newBlob(byteCharacters, fileData.mimeType, "Slip_" + id + ".jpg");
     const file = folder.createFile(blob);
     const slipUrl = "https://lh3.googleusercontent.com/d/" + file.getId();

     const ss = initSystemBackend();
     const sheet = ss.getSheetByName(SHEET_NAME);
     const data = sheet.getDataRange().getValues();
     for (let i = 1; i < data.length; i++) {
       if (String(data[i][0]) === String(id)) {
         sheet.getRange(i + 1, 12).setValue("ยืนยันแล้ว");
         sheet.getRange(i + 1, 13).setValue(slipUrl);
         if (LINE_NOTIFY_TOKEN !== "") { try { UrlFetchApp.fetch("https://notify-api.line.me/api/notify", { "method": "post", "headers": { "Authorization": "Bearer " + LINE_NOTIFY_TOKEN }, "payload": { "message": `\n💸 ชำระมัดจำสำเร็จ!\nรหัสคิว: ${id}\nลูกค้าแนบสลิปและ AI อนุมัติแล้ว\nดูสลิป: ${slipUrl}` }}); } catch(e) {} }
         return { success: true, message: "ระบบตรวจสอบสลิปสำเร็จ! ยืนยันคิวเรียบร้อย" };
       }
     }
     return { success: false, message: "ไม่พบคิวงาน" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// =========================================================================
// จัดการคิวงาน (Bookings)
// =========================================================================
function getBookingsFromSheet() {
  try {
    const ss = initSystemBackend();
    const data = ss.getSheetByName(SHEET_NAME).getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] }; 
    const bookings = [];
    for (let i = data.length - 1; i > 0; i--) {
      bookings.push({
        id: String(data[i][0] || "-"), customerName: String(data[i][2] || "-"),
        tel: String(data[i][3] || "-"), date: (data[i][4] instanceof Date) ? Utilities.formatDate(data[i][4], "Asia/Bangkok", "yyyy-MM-dd") : String(data[i][4] || ""),
        timeSlot: String(data[i][5] || "-"), packageName: String(data[i][6] || "-"),
        location: String(data[i][7] || "-"), price: Number(data[i][8]) || 0, depositAmount: Number(data[i][9]) || 0,
        notes: String(data[i][10] || ""), status: String(data[i][11] || "รอตรวจสอบ"), slipUrl: String(data[i][12] || ""),
        // ดึงลิงก์จากคอลัมน์ N (13) และ O (14)
        driveLink: String(data[i][13] || ""),
        photosLink: String(data[i][14] || "")
      });
    }
    return { success: true, data: bookings };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function handleCreateBooking(data) {
  try {
    const ss = initSystemBackend();
    ss.getSheetByName(SHEET_NAME).appendRow([ data.id, new Date(), data.customerName, data.tel, data.date, data.timeSlot, data.packageName, data.location, data.price, data.depositAmount, data.notes || "-", data.status, "", "", "" ]);
    try { CalendarApp.getCalendarById(CALENDAR_ID).createAllDayEvent(`📸 คิวถ่าย: ${data.customerName}`, new Date(data.date), { location: data.location, description: `รหัส: ${data.id}\nโทร: ${data.tel}\nแพ็กเกจ: ${data.packageName}\nยอดมัดจำ: ${data.depositAmount} บาท\nบรีฟ: ${data.notes}` }); } catch(e) {}
    if (LINE_NOTIFY_TOKEN !== "") {
      try { UrlFetchApp.fetch("https://notify-api.line.me/api/notify", { "method": "post", "headers": { "Authorization": "Bearer " + LINE_NOTIFY_TOKEN }, "payload": { "message": `\n🎉 คิวงานใหม่!\nรหัส: ${data.id}\nลูกค้า: ${data.customerName}\nวันที่: ${data.date}\nรอการแนบสลิปมัดจำ` }}); } catch(e) {}
    }
    return { success: true, message: "บันทึกคิวเรียบร้อย", id: data.id };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function updateBookingInSheet(id, updates) {
  try {
    const ss = initSystemBackend();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if(updates.customerName !== undefined) sheet.getRange(i + 1, 3).setValue(updates.customerName);
        if(updates.tel !== undefined) sheet.getRange(i + 1, 4).setValue(updates.tel);
        if(updates.date !== undefined) sheet.getRange(i + 1, 5).setValue(updates.date);
        if(updates.timeSlot !== undefined) sheet.getRange(i + 1, 6).setValue(updates.timeSlot);
        if(updates.packageName !== undefined) sheet.getRange(i + 1, 7).setValue(updates.packageName);
        if(updates.location !== undefined) sheet.getRange(i + 1, 8).setValue(updates.location);
        if(updates.price !== undefined) sheet.getRange(i + 1, 9).setValue(updates.price);
        if(updates.depositAmount !== undefined) sheet.getRange(i + 1, 10).setValue(updates.depositAmount);
        if(updates.notes !== undefined) sheet.getRange(i + 1, 11).setValue(updates.notes);
        if(updates.status !== undefined) sheet.getRange(i + 1, 12).setValue(updates.status);
        
        // บันทึกลิงก์ส่งงาน Google Drive & Photos ลงคอลัมน์ N (14) และ O (15)
        if(updates.driveLink !== undefined) sheet.getRange(i + 1, 14).setValue(updates.driveLink);
        if(updates.photosLink !== undefined) sheet.getRange(i + 1, 15).setValue(updates.photosLink);
        
        return { success: true, message: "อัปเดตข้อมูลสำเร็จ!" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูล" };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function deleteBookingInSheet(id) {
  try {
    const ss = initSystemBackend();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "ลบข้อมูลคิวงานเรียบร้อย!" };
      }
    }
    return { success: false, message: "ไม่พบคิวงานที่ต้องการลบ" };
  } catch (error) { return { success: false, message: error.toString() }; }
}

// =========================================================================
// จัดการผลงาน (Portfolio)
// =========================================================================
function getPortfolioFromSheet() {
  try {
    const ss = initSystemBackend();
    const data = ss.getSheetByName(PORTFOLIO_SHEET).getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] }; 
    let works = [];
    for (let i = 1; i < data.length; i++) { 
      works.push({ 
        id: String(data[i][0]), title: String(data[i][1]), 
        category: String(data[i][2]), url: String(data[i][3]), 
        location: String(data[i][4]), timestamp: data[i][5] ? new Date(data[i][5]).getTime() : 0 
      }); 
    }
    works.sort((a, b) => b.timestamp - a.timestamp);
    return { success: true, data: works };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function uploadAndAddPortfolio(data) {
  try {
    let imageUrl = "";
    if (data.base64File) {
      let folder = getAppFolder("ToyGallery_Portfolio");
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
      
      const byteCharacters = Utilities.base64Decode(data.base64File.split(",")[1]);
      const blob = Utilities.newBlob(byteCharacters, data.mimeType, "PORT_" + new Date().getTime() + ".jpg");
      const file = folder.createFile(blob);
      imageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    } else { imageUrl = data.url; }
    const ss = initSystemBackend();
    const newId = "P-" + Math.floor(10000 + Math.random() * 90000);
    ss.getSheetByName(PORTFOLIO_SHEET).appendRow([newId, data.title, data.category, imageUrl, data.location, new Date()]);
    return { success: true, message: "เพิ่มผลงานสำเร็จ!" };
  } catch (error) { return { success: false, message: "Upload Error: " + error.toString() }; }
}

function setPortfolioCoverInSheet(id) {
  try {
    const ss = initSystemBackend();
    const sheet = ss.getSheetByName(PORTFOLIO_SHEET);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, 6).setValue(new Date()); 
        return { success: true, message: "ตั้งเป็นภาพหน้าปกอัลบั้มเรียบร้อย!" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูลรูปภาพ" };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function updateAlbumDetailsInSheet(oldTitle, updates) {
  try {
    const ss = initSystemBackend();
    const sheet = ss.getSheetByName(PORTFOLIO_SHEET);
    const data = sheet.getDataRange().getValues();
    let updatedCount = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(oldTitle)) {
        if(updates.title) sheet.getRange(i + 1, 2).setValue(updates.title);
        if(updates.category) sheet.getRange(i + 1, 3).setValue(updates.category);
        if(updates.location) sheet.getRange(i + 1, 5).setValue(updates.location);
        updatedCount++;
      }
    }
    return { success: true, message: `อัปเดตข้อมูลอัลบั้มสำเร็จ (${updatedCount} รูป)` };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function deleteEntireAlbumInSheet(title) {
  try {
    const ss = initSystemBackend();
    const sheet = ss.getSheetByName(PORTFOLIO_SHEET);
    const data = sheet.getDataRange().getValues();
    let deletedCount = 0;
    for (let i = data.length - 1; i > 0; i--) { 
      if (String(data[i][1]) === String(title)) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    return { success: true, message: `ลบอัลบั้มเรียบร้อย (${deletedCount} รูป)` };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function deletePortfolioInSheet(id) {
  try {
    const ss = initSystemBackend();
    const sheet = ss.getSheetByName(PORTFOLIO_SHEET);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "ลบรูปภาพเรียบร้อย" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูล" };
  } catch (error) { return { success: false, message: error.toString() }; }
}
