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


// ==================== 📢 ฟังก์ชันส่ง LINE Notify ====================

// /**
//  * ส่งการแจ้งเตือนผ่าน LINE Notify ไปยังกลุ่ม/แชทที่กำหนด
//  * @param {Object} data - ข้อมูลที่จะส่ง
//  * @returns {Object} ผลลัพธ์การส่ง
//  */
// function sendLineNotify(data) {
//   try {
//     // ตรวจสอบว่ามี Token หรือไม่
//     if (!LINE_NOTIFY_TOKEN || LINE_NOTIFY_TOKEN === 'ใส่ LINE Notify Token ของคุณที่นี่') {
//       console.warn('⚠️ LINE Notify Token ยังไม่ได้ตั้งค่า');
//       return { 
//         success: false, 
//         message: 'LINE Notify Token ยังไม่ได้ตั้งค่า (ข้ามขั้นตอนนี้)' 
//       };
//     }
    
//     // สร้างข้อความแจ้งเตือน
//     const message = `
// 🎉 มีการลงทะเบียนใหม่!

// 👤 ชื่อ: ${data.name}
// 📞 เบอร์: ${data.phone}
// 🎯 กิจกรรม: ${data.option}
// 📍 รายการ: ${data.sub_option}
// ${data.details ? '📝 รายละเอียด: ' + data.details : ''}

// 👥 ชื่อใน LINE: ${data.displayName || '-'}
// 🆔 User ID: ${data.lineUserId}
// ⏰ เวลา: ${data.timestamp || new Date().toLocaleString('th-TH')}
//     `.trim();
    
//     // ส่งข้อความผ่าน LINE Notify API
//     const url = 'https://notify-api.line.me/api/notify';
//     const options = {
//       method: 'post',
//       headers: {
//         'Authorization': 'Bearer ' + LINE_NOTIFY_TOKEN,
//         'Content-Type': 'application/x-www-form-urlencoded'
//       },
//       payload: {
//         message: message
//       }
//     };
    
//     const response = UrlFetchApp.fetch(url, options);
//     const result = JSON.parse(response.getContentText());
    
//     console.log('📨 LINE Notify API response:', result);
    
//     if (result.status === 200) {
//       return {
//         success: true,
//         message: 'ส่ง LINE Notify สำเร็จ',
//         response: result
//       };
//     } else {
//       throw new Error('LINE Notify API returned status: ' + result.status);
//     }
    
//   } catch (error) {
//     console.error('❌ เกิดข้อผิดพลาดในการส่ง LINE Notify:', error);
//     // ไม่ throw error เพื่อไม่ให้การบันทึกข้อมูลล้มเหลว
//     return {
//       success: false,
//       message: 'ส่ง LINE Notify ไม่สำเร็จ: ' + error.message
//     };
//   }
// }


// // ==================== 💬 ฟังก์ชันส่งข้อความตอบกลับผ่าน Messaging API ====================

// /**
//  * ส่งข้อความตอบกลับไปหาผู้ใช้ผ่าน LINE Messaging API
//  * @param {Object} data - ข้อมูลที่จะส่ง
//  * @returns {Object} ผลลัพธ์การส่ง
//  */
// function sendLineMessage(data) {
//   try {
//     // ตรวจสอบว่ามี Token หรือไม่
//     if (!LINE_MESSAGING_TOKEN || LINE_MESSAGING_TOKEN === 'ใส่ Messaging API Token ของคุณที่นี่') {
//       console.warn('⚠️ LINE Messaging API Token ยังไม่ได้ตั้งค่า');
//       return { 
//         success: false, 
//         message: 'LINE Messaging API Token ยังไม่ได้ตั้งค่า (ข้ามขั้นตอนนี้)' 
//       };
//     }
    
//     // ตรวจสอบว่ามี User ID หรือไม่
//     if (!data.lineUserId) {
//       console.warn('⚠️ ไม่พบ LINE User ID');
//       return { 
//         success: false, 
//         message: 'ไม่พบ LINE User ID' 
//       };
//     }
    
//     // สร้างข้อความตอบกลับ
//     const replyMessage = `✅ ลงทะเบียนสำเร็จ!

// สวัสดีคุณ ${data.name} 👋

// 📋 ข้อมูลการลงทะเบียนของคุณ:
// 👤 ชื่อ: ${data.name}
// 📞 เบอร์: ${data.phone}
// 🎯 กิจกรรม: ${data.option}
// 📍 รายการ: ${data.sub_option}
// ${data.details ? '📝 รายละเอียด: ' + data.details : ''}

// ⏰ วันที่: ${data.timestamp || new Date().toLocaleString('th-TH')}

// ขอบคุณที่ลงทะเบียนค่ะ 🙏
// เราจะติดต่อกลับไปในเร็วๆ นี้`;
    
//     // ส่งข้อความผ่าน LINE Messaging API
//     const url = 'https://api.line.me/v2/bot/message/push';
//     const options = {
//       method: 'post',
//       headers: {
//         'Content-Type': 'application/json',
//         'Authorization': 'Bearer ' + LINE_MESSAGING_TOKEN
//       },
//       payload: JSON.stringify({
//         to: data.lineUserId,
//         messages: [
//           {
//             type: 'text',
//             text: replyMessage
//           }
//         ]
//       }),
//       muteHttpExceptions: true
//     };
    
//     const response = UrlFetchApp.fetch(url, options);
//     const responseCode = response.getResponseCode();
//     const result = JSON.parse(response.getContentText());
    
//     console.log('💬 LINE Messaging API response code:', responseCode);
//     console.log('💬 LINE Messaging API response:', result);
    
//     if (responseCode === 200) {
//       return {
//         success: true,
//         message: 'ส่งข้อความตอบกลับสำเร็จ',
//         response: result
//       };
//     } else {
//       throw new Error('LINE Messaging API returned code: ' + responseCode + ', message: ' + JSON.stringify(result));
//     }
    
//   } catch (error) {
//     console.error('❌ เกิดข้อผิดพลาดในการส่งข้อความ:', error);
//     // ไม่ throw error เพื่อไม่ให้การบันทึกข้อมูลล้มเหลว
//     return {
//       success: false,
//       message: 'ส่งข้อความตอบกลับไม่สำเร็จ: ' + error.message
//     };
//   }
// }


// ==================== 🧪 ฟังก์ชันทดสอบระบบ ====================

/**
 * ฟังก์ชันสำหรับทดสอบระบบทั้งหมด
 * ใช้โดยเลือกฟังก์ชันนี้ใน Editor แล้วกด Run
 */
function testProcessForm() {
  console.log('🧪 เริ่มต้นการทดสอบระบบ...');
  
  // ข้อมูลทดสอบ
  const testData = {
    name: 'ทดสอบ ระบบ',
    phone: '081-234-5678',
    option: 'ธรรมยาตรา',
    sub_option: '3 ม.ค. โลตัสแลนด์',
    details: 'นี่คือข้อมูลทดสอบระบบ',
    lineUserId: 'Utest1234567890abcdef',
    displayName: 'Test User',
    pictureUrl: 'https://example.com/picture.jpg',
    timestamp: new Date().toLocaleString('th-TH')
  };
  
  console.log('📋 ข้อมูลทดสอบ:', testData);
  
  try {
    const result = processForm(testData);
    console.log('✅ ผลการทดสอบ:', result);
    console.log('');
    console.log('==================== สรุปผลการทดสอบ ====================');
    console.log('Sheet:', result.results.sheetResult.success ? '✅ สำเร็จ' : '❌ ล้มเหลว');
    console.log('LINE Notify:', result.results.lineNotifyResult.success ? '✅ สำเร็จ' : '❌ ล้มเหลว');
    console.log('LINE Message:', result.results.messageResult.success ? '✅ สำเร็จ' : '❌ ล้มเหลว');
    console.log('======================================================');
    return result;
  } catch (error) {
    console.error('❌ การทดสอบล้มเหลว:', error);
    return { success: false, error: error.message };
  }
}

/**
 * ทดสอบการบันทึก Google Sheet เท่านั้น
 */
function testSheetOnly() {
  console.log('📊 ทดสอบการบันทึก Google Sheet...');
  
  const testData = {
    name: 'ทดสอบ Sheet',
    phone: '081-111-2222',
    option: 'ธรรมยาตรา',
    sub_option: '3 ม.ค. โลตัสแลนด์',
    details: 'ทดสอบ Sheet',
    lineUserId: 'Utest123',
    displayName: 'Test',
    timestamp: new Date().toLocaleString('th-TH')
  };
  
  try {
    const result = saveToSheet(testData);
    console.log('✅ ผลการทดสอบ:', result);
    return result;
  } catch (error) {
    console.error('❌ ทดสอบล้มเหลว:', error);
    return { success: false, error: error.message };
  }
}

/**
 * ทดสอบ LINE Notify เท่านั้น
 */
function testLineNotifyOnly() {
  console.log('📢 ทดสอบการส่ง LINE Notify...');
  
  const testData = {
    name: 'ทดสอบ Notify',
    phone: '081-111-2222',
    option: 'ธรรมยาตรา',
    sub_option: '3 ม.ค. โลตัสแลนด์',
    details: 'ทดสอบ LINE Notify',
    lineUserId: 'Utest123',
    displayName: 'Test User',
    timestamp: new Date().toLocaleString('th-TH')
  };
  
  try {
    const result = sendLineNotify(testData);
    console.log('✅ ผลการทดสอบ:', result);
    return result;
  } catch (error) {
    console.error('❌ ทดสอบล้มเหลว:', error);
    return { success: false, error: error.message };
  }
}


// ==================== 📋 ฟังก์ชันเสริม ====================

/**
 * ดูข้อมูลการตั้งค่า
 */
function checkConfiguration() {
  console.log('==================== ตรวจสอบการตั้งค่า ====================');
  console.log('Spreadsheet ID:', SPREADSHEET_ID);
  console.log('Sheet Name:', SHEET_NAME);
  console.log('LINE Notify Token:', LINE_NOTIFY_TOKEN ? '✅ ตั้งค่าแล้ว' : '❌ ยังไม่ได้ตั้งค่า');
  console.log('LINE Messaging Token:', LINE_MESSAGING_TOKEN ? '✅ ตั้งค่าแล้ว' : '❌ ยังไม่ได้ตั้งค่า');
  console.log('===========================================================');
}

/**
 * ทดสอบการเชื่อมต่อ Spreadsheet
 */
function testSpreadsheetConnection() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ เชื่อมต่อ Spreadsheet สำเร็จ');
    console.log('ชื่อไฟล์:', ss.getName());
    console.log('URL:', ss.getUrl());
    return true;
  } catch (error) {
    console.error('❌ ไม่สามารถเชื่อมต่อ Spreadsheet:', error.message);
    return false;
  }
}
