// 1. ฟังก์ชันเปิดหน้าเว็บ
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบเลือกโปรแกรมตรวจสุขภาพ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. ฟังก์ชันค้นหาข้อมูลผู้ใช้ (Sheet 1)
function searchUser(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; // แผ่นงานที่ 1 (ข้อมูลพนักงาน)
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    // แปลงเป็น String เพื่อป้องกัน Error
    if (String(data[i][0]) === String(id)) {
      
      // จัดการวันที่ (ดึงค่ามาเป็น Text หรือ Date แล้วแปลงเป็น dd/MM/yyyy)
      var dobRaw = data[i][4]; 
      var dobFormatted = "";
      
      if (dobRaw && (dobRaw instanceof Date)) {
        dobFormatted = Utilities.formatDate(dobRaw, "GMT+7", "dd/MM/yyyy");
      } else {
        dobFormatted = String(dobRaw); 
      }

      return {
        found: true,
        row: i + 1,
        prefix: data[i][1],      
        name: data[i][2],        
        surname: data[i][3],     
        dob: dobFormatted,       
        age: data[i][5],         
        program: data[i][6]      
      };
    }
  }
  return { found: false };
}

// 3. ฟังก์ชันบันทึกข้อมูลกลับลง Sheet (Sheet 1)
function updateUserData(form) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var row = form.row;

  sheet.getRange(row, 2).setValue(form.prefix);  
  sheet.getRange(row, 3).setValue(form.name);    
  sheet.getRange(row, 4).setValue(form.surname); 
  sheet.getRange(row, 5).setValue(form.dob);     // บันทึกวันที่เป็น Text (dd/MM/yyyy)
  sheet.getRange(row, 6).setValue(form.age);     
  sheet.getRange(row, 7).setValue(form.program); 

  return "✅ บันทึกข้อมูลเรียบร้อยแล้ว!";
}

// 4. ฟังก์ชันดึงรายละเอียดโปรแกรมตรวจ (Sheet 2)
function getProgramDetails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1]; // แผ่นงานที่ 2 (ตารางโปรแกรมตรวจ)
  
  // ดึงข้อมูลช่วง B4:L19 (ครอบคลุมชื่อรายการตรวจ และ ราคาด้านล่าง)
  // B=ชื่อรายการ, E=โปรแกรม1 ... L=โปรแกรม8
  var data = sheet.getRange("B4:L19").getValues();
  
  var programs = {};
  
  // วนลูปสร้างข้อมูลโปรแกรม 1-8
  for (var p = 1; p <= 8; p++) {
    var colIndex = p + 2; // คอลัมน์ E เริ่มที่ index 3 (ใน array นี้)
    
    var items = [];
    // วนลูปรายการตรวจ (แถว 0-13 คือรายการแพทย์ - PSA)
    for (var i = 0; i < 14; i++) {
      var checkMark = data[i][colIndex];
      // ถ้ามีเครื่องหมายถูก หรือมีค่า
      if (checkMark && String(checkMark).trim() !== "") {
        items.push(data[i][0]); // เก็บชื่อรายการ (Column B)
      }
    }
    
    // ดึงราคาส่วนเกิน (แถวที่ 15 ใน array คือ row 19 ใน Excel 'ราคาส่วนเกินตามสิทธิ')
    // *หมายเหตุ: ถ้าต้องการใช้ราคาบุคคลทั่วไป (row 20) ให้เปลี่ยนเลข 14 เป็น 15
    var price = data[14][colIndex];
    if (price === "" || price == null) price = 0;

    programs[p] = {
      items: items,
      price: price
    };
  }
  
  return programs;
}
