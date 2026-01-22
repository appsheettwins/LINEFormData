// ==================== ⚙️ ตั้งค่าระบบ ====================
// ⚠️ สำคัญ: กรุณาแก้ไขค่าเหล่านี้ก่อนใช้งาน

const LINE_NOTIFY_TOKEN = 'ใส่ LINE Notify Token ของคุณที่นี่'; 
// วิธีสร้าง LINE Notify Token:
// 1. ไปที่ https://notify-bot.line.me/
// 2. คลิก My page > Generate token
// 3. เลือกกลุ่ม/แชทที่ต้องการส่งการแจ้งเตือน
// 4. คัดลอก Token มาวางที่นี่

const SPREADSHEET_ID = '1lUfArwkheK2JMntkO6zwzXmJL-F5T_qZ8e3VT5W4Sgc';
// Spreadsheet ID จาก URL: https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit

const SHEET_NAME = 'ลงทะเบียนกิจกรรม';
// ชื่อแผ่นงานที่ต้องการบันทึกข้อมูล (ถ้าไม่มีจะสร้างอัตโนมัติ)

const LINE_MESSAGING_TOKEN = 'pUcYHL7II8uYofiWV01d84F/gZJkFR3hoDMU/EE1+C7rWJhrYskpfpMsm8bTgs1pcSB1Htc7Waf34BM5biopYhTEed9cCHiV1JscVL3YddvkDU2PXLQkjTv1t2Kwtmmh97sviafI8ft/TFgqXyMIVgdB04t89/1O/w1cDnyilFU=';
// Messaging API Channel Access Token จาก LINE Developers Console


// ==================== 🌐 ฟังก์ชันแสดงหน้า Web ====================

/**
 * ฟังก์ชันสำหรับแสดงหน้า HTML เมื่อเข้าถึง Web App
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ลงทะเบียนกิจกรรม')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}



// /**
//  * ฟังก์ชันหลักในการรับและประมวลผลข้อมูลจากฟอร์ม
//  */
// function processForm(data) {
//   try {
//     console.log('==================== เริ่มต้นประมวลผลฟอร์ม ====================');
//     const results = {};
    
//     // 1. บันทึกข้อมูลลง Google Sheet (เก็บส่วนนี้ไว้)
//     console.log('📊 กำลังบันทึกลง Google Sheet...');
//     try {
//       results.sheetResult = saveToSheet(data);
//       console.log('✅ บันทึก Google Sheet สำเร็จ');
//     } catch (error) {
//       console.error('❌ บันทึก Google Sheet ล้มเหลว:', error);
//       results.sheetResult = { success: false, message: error.message };
//       throw error; // ให้หยุดทำงานถ้าบันทึก Sheet ไม่สำเร็จ
//     }
    
//     /* --- ปิดส่วนการส่ง LINE ออกทั้งหมด ---
    
//     // 2. ส่ง LINE Notify
//     // (ปิดการใช้งาน)
    
//     // 3. ส่งข้อความตอบกลับ
//     // (ปิดการใช้งาน)
    
//     -------------------------------------- */
    
//     console.log('==================== ประมวลผลเสร็จสิ้น ====================');
    
//     return {
//       success: true,
//       message: 'บันทึกข้อมูลลง Google Sheet เรียบร้อยแล้ว',
//       timestamp: new Date().toLocaleString('th-TH'),
//       results: results
//     };
    
//   } catch (error) {
//     console.error('❌ เกิดข้อผิดพลาด:', error);
//     throw new Error('เกิดข้อผิดพลาด: ' + error.message);
//   }
// }






// ==================== 📊 ฟังก์ชันบันทึกข้อมูลลง Google Sheet ====================

/**
 * บันทึกข้อมูลลง Google Sheets
 * @param {Object} data - ข้อมูลที่จะบันทึก
 * @returns {Object} ผลลัพธ์การบันทึก
 */
// function saveToSheet(data) {
//   try {
//     // เปิด Spreadsheet
//     const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
//     let sheet = ss.getSheetByName(SHEET_NAME);
    
//     // สร้างแผ่นงานใหม่ถ้ายังไม่มี
//     if (!sheet) {
//       console.log('📄 สร้างแผ่นงานใหม่:', SHEET_NAME);
//       sheet = ss.insertSheet(SHEET_NAME);
      
//       // เพิ่มหัวตาราง
//       const headers = [
//         'วันที่-เวลา',
//         'ชื่อ-นามสกุล', 
//         'เบอร์โทร',
//         'กิจกรรม',
//         'รายการย่อย',
//         'รายละเอียด',
//         'LINE User ID',
//         'LINE Display Name',
//         'Picture URL'
//       ];
//       sheet.appendRow(headers);
      
//       // จัดรูปแบบหัวตาราง
//       const headerRange = sheet.getRange(1, 1, 1, headers.length);
//       headerRange.setBackground('#667eea');
//       headerRange.setFontColor('#ffffff');
//       headerRange.setFontWeight('bold');
//       headerRange.setHorizontalAlignment('center');
//       headerRange.setVerticalAlignment('middle');
      
//       // ตั้งค่าความกว้างคอลัมน์
//       sheet.setColumnWidth(1, 150); // วันที่-เวลา
//       sheet.setColumnWidth(2, 150); // ชื่อ-นามสกุล
//       sheet.setColumnWidth(3, 120); // เบอร์โทร
//       sheet.setColumnWidth(4, 120); // กิจกรรม
//       sheet.setColumnWidth(5, 180); // รายการย่อย
//       sheet.setColumnWidth(6, 200); // รายละเอียด
//       sheet.setColumnWidth(7, 250); // LINE User ID
//       sheet.setColumnWidth(8, 150); // LINE Display Name
//       sheet.setColumnWidth(9, 250); // Picture URL
      
//       // ตรึงแถวแรก
//       sheet.setFrozenRows(1);
      
//       console.log('✅ สร้างแผ่นงานและหัวตารางสำเร็จ');
//     }
    
//     // เพิ่มข้อมูลแถวใหม่
//     const rowData = [
//       data.timestamp || new Date().toLocaleString('th-TH'),
//       data.name || '',
//       data.phone || '',
//       data.option || '',
//       data.sub_option || '',
//       data.details || '-',
//       data.lineUserId || '',
//       data.displayName || '',
//       data.pictureUrl || ''
//     ];
    
//     sheet.appendRow(rowData);
    
//     // จัดรูปแบบแถวข้อมูล
//     const lastRow = sheet.getLastRow();
//     const dataRange = sheet.getRange(lastRow, 1, 1, rowData.length);
    
//     // เพิ่มเส้นขอบ
//     dataRange.setBorder(
//       true, true, true, true, true, true,
//       '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID
//     );
    
//     // จัดแนวข้อความ
//     dataRange.setVerticalAlignment('middle');
//     sheet.getRange(lastRow, 1, 1, 1).setHorizontalAlignment('center'); // วันที่-เวลา
//     sheet.getRange(lastRow, 2, 1, 1).setHorizontalAlignment('left');   // ชื่อ
//     sheet.getRange(lastRow, 3, 1, 1).setHorizontalAlignment('center'); // เบอร์
    
//     // สลับสีแถว
//     if (lastRow % 2 === 0) {
//       dataRange.setBackground('#f9fafb');
//     }
    
//     console.log('✅ บันทึกข้อมูลสำเร็จที่แถว:', lastRow);
    
//     return {
//       success: true,
//       message: 'บันทึกลง Google Sheet สำเร็จ',
//       row: lastRow,
//       sheetName: SHEET_NAME,
//       spreadsheetUrl: ss.getUrl()
//     };
    
//   } catch (error) {
//     console.error('❌ เกิดข้อผิดพลาดในการบันทึก Sheet:', error);
//     throw new Error('ไม่สามารถบันทึกลง Google Sheet: ' + error.message);
//   }
// }

// อย่าลืมประกาศตัวแปรเหล่านี้ไว้ด้านบนสุดของไฟล์
const SPREADSHEET_ID = '1lUfArwkheK2JMntkO6zwzXmJL-F5T_qZ8e3VT5W4Sgc';
const SHEET_NAME = 'ลงทะเบียนกิจกรรม';

function saveToSheet(data) {
  try {
    // 1. เปิด Spreadsheet และตรวจสอบ Sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // 2. ถ้าไม่มี Sheet ให้สร้างใหม่พร้อมตั้งค่าเริ่มต้น (ทำครั้งเดียว)
    if (!sheet) {
      console.log('📄 สร้างแผ่นงานใหม่:', SHEET_NAME);
      sheet = ss.insertSheet(SHEET_NAME);
      
      const headers = [
        'วันที่-เวลา', 'ชื่อ-นามสกุล', 'เบอร์โทร', 'กิจกรรม', 
        'รายการย่อย', 'รายละเอียด', 'LINE User ID', 
        'LINE Display Name', 'Picture URL'
      ];
      sheet.appendRow(headers);
      
      // จัดรูปแบบหัวตาราง (Header)
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#667eea')
                 .setFontColor('#ffffff')
                 .setFontWeight('bold')
                 .setHorizontalAlignment('center')
                 .setVerticalAlignment('middle');
      
      // ตั้งค่าความกว้างคอลัมน์
      const widths = [180, 150, 120, 120, 180, 200, 250, 150, 250];
      widths.forEach((width, index) => {
        sheet.setColumnWidth(index + 1, width);
      });
      
      sheet.setFrozenRows(1);
    }
    
    // 3. เตรียมข้อมูล (ใช้ new Date() เพื่อให้ Google Sheet มองเป็นวันที่จริงๆ)
    const rowData = [
      data.timestamp ? new Date(data.timestamp) : new Date(), 
      data.name || '',
      data.phone || '',
      data.option || '',
      data.sub_option || '',
      data.details || '-',
      data.lineUserId || '',
      data.displayName || '',
      data.pictureUrl || ''
    ];
    
    // 4. บันทึกข้อมูล
    sheet.appendRow(rowData);
    
    // 5. จัดรูปแบบแถวที่เพิ่งเพิ่ม (ทำเฉพาะที่จำเป็นเพื่อความเร็ว)
    const lastRow = sheet.getLastRow();
    const lastColumn = rowData.length;
    const dataRange = sheet.getRange(lastRow, 1, 1, lastColumn);
    
    // จัดตำแหน่งข้อความ
    dataRange.setVerticalAlignment('middle');
    sheet.getRange(lastRow, 1).setHorizontalAlignment('center'); // วันที่
    sheet.getRange(lastRow, 3).setHorizontalAlignment('center'); // เบอร์โทร
    
    // ใส่เส้นขอบแบบเบาๆ และสลับสี (Optional)
    dataRange.setBorder(true, true, true, true, true, true, '#e5e7eb', SpreadsheetApp.BorderStyle.SOLID);
    if (lastRow % 2 === 0) {
      dataRange.setBackground('#f9fafb');
    }

    console.log('✅ บันทึกข้อมูลสำเร็จที่แถว:', lastRow);
    
    return {
      success: true,
      row: lastRow,
      sheetName: SHEET_NAME
    };
    
  } catch (error) {
    console.error('❌ เกิดข้อผิดพลาด:', error);
    return { success: false, message: error.message };
  }
}
