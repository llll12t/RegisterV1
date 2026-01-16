// ====================================================================================
// การตั้งค่าหลัก (Configuration)
// ====================================================================================
var SPREADSHEET_ID = '1g7uhBnYkpmCXt5Dkf89qIus86CrHkIZK0b6yw1T2pp8';
var DRIVE_FOLDER_ID = '1zCDYqKso6U-T69wWqDwm5KI6_Ds8BzVs';
var SHEET_NAME_MEMBER = 'Member';

// ====================================================================================
// Column Index Configuration (0-based index)
// กำหนดตำแหน่งคอลัมน์ - สามารถเปลี่ยนชื่อหัวคอลัมน์ในชีตได้อิสระ
// ระบบจะอ้างอิงจาก index ไม่ใช่ชื่อ
// ====================================================================================
var COL = {
  ID: 0,            // คอลัมน์ A (index 0): id หรือ รหัส
  USER_LINE_ID: 1,  // คอลัมน์ B (index 1): userIDline หรือ LINE ID
  NAME: 2,          // คอลัมน์ C (index 2): ชื่อ สกุล
  CODE: 3,          // คอลัมน์ D (index 3): รหัสประจำตัว
  DEPARTMENT: 4,    // คอลัมน์ E (index 4): แผนก
  IMAGE: 5,         // คอลัมน์ F (index 5): รูปโปรไฟล์
  TIMESTAMP: 6      // คอลัมน์ G (index 6): บันทึกเวลา
};

// หัวคอลัมน์เริ่มต้น (ใช้เมื่อสร้างชีตใหม่)
var DEFAULT_HEADERS = ['ID', 'IDline', 'ชื่อ สกุล', 'รหัสประจำตัว', 'แผนก', 'รูปโปรไฟล์', 'บันทึกเวลา'];

// ====================================================================================
// ฟังก์ชันจัดการรูปภาพ
// ====================================================================================
function generateFileUrl(data, folderId) {
  var imageName = `${data.name}_${Utilities.formatDate(new Date(), 'GMT+7', 'dd-MM-yyyy_HH:mm:ss')}.png`;
  var decodedImage = Utilities.base64Decode(data.base64);
  var blob = Utilities.newBlob(decodedImage, 'image/png', imageName);
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://lh5.googleusercontent.com/d/${file.getId()}`;
}

function uploadBase64Image(base64Image, name) {
  var imageName = `${name}_${Utilities.formatDate(new Date(), 'GMT+7', 'dd-MM-yyyy_HH:mm:ss')}.png`;
  var decodedImage = Utilities.base64Decode(base64Image);
  var blob = Utilities.newBlob(decodedImage, 'image/png', imageName);
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return `https://lh5.googleusercontent.com/d/${file.getId()}`;
}

// ====================================================================================
// ฟังก์ชันเพิ่ม/อัปเดตข้อมูลสมาชิก
// ====================================================================================
function handleMemberDataInsertion(data) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MEMBER);

  // สร้างชีตใหม่ถ้าไม่มี
  if (!sheet) {
    sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_NAME_MEMBER);
    sheet.appendRow(DEFAULT_HEADERS);
  }

  var rows = sheet.getDataRange().getValues();
  var existingRowIndex = -1;

  // ค้นหา userlineId ที่มีอยู่แล้ว
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][COL.USER_LINE_ID] === data.userlineId) {
      existingRowIndex = i + 1;
      break;
    }
  }

  var fileLink = generateFileUrl(data, DRIVE_FOLDER_ID);
  var timestamp = Utilities.formatDate(new Date(), 'GMT+7', 'dd-MM-yyyy HH:mm:ss');

  if (existingRowIndex > 0) {
    // อัปเดตข้อมูลที่มีอยู่
    sheet.getRange(existingRowIndex, COL.NAME + 1).setValue(data.nameId);
    sheet.getRange(existingRowIndex, COL.CODE + 1).setValue("'" + data.keynumberId);
    sheet.getRange(existingRowIndex, COL.DEPARTMENT + 1).setValue("'" + data.keynumber2Id);
    sheet.getRange(existingRowIndex, COL.IMAGE + 1).setValue(fileLink);
    sheet.getRange(existingRowIndex, COL.TIMESTAMP + 1).setValue(timestamp);
  } else {
    // เพิ่มข้อมูลใหม่
    var nextRow = sheet.getLastRow() + 1;
    var keyId = nextRow - 1;
    var rowData = [];
    rowData[COL.ID] = keyId;
    rowData[COL.USER_LINE_ID] = data.userlineId;
    rowData[COL.NAME] = data.nameId;
    rowData[COL.CODE] = "'" + data.keynumberId;
    rowData[COL.DEPARTMENT] = "'" + data.keynumber2Id;
    rowData[COL.IMAGE] = fileLink;
    rowData[COL.TIMESTAMP] = timestamp;
    sheet.getRange(nextRow, 1, 1, 7).setValues([rowData]);
  }
}

// ====================================================================================
// ฟังก์ชันจัดการคำขอ HTTP
// ====================================================================================
function doGet(e) {
  if (e.parameter.method === 'getUserData' && e.parameter.userlineId) {
    return fetchDataResponseByUserlineId(e.parameter.userlineId);
  }
  if (e.parameter.action === 'fetchDataByName' && e.parameter.nameId) {
    return fetchDataResponseByName(e.parameter.nameId);
  }
  return fetchDataResponse();
}

function doPost(e) {
  try {
    // ตรวจสอบ JSON payload
    if (e.postData && e.postData.contents) {
      try {
        var obj = JSON.parse(e.postData.contents);
        if (obj.action === "fetchData") {
          return fetchDataResponse();
        }
        if (obj.hasOwnProperty('nameId')) {
          handleMemberDataInsertion(obj);
          return ContentService.createTextOutput("Data saved successfully").setMimeType(ContentService.MimeType.TEXT);
        }
      } catch (jsonError) {
        console.log('JSON parsing error:', jsonError);
      }
    }

    // ใช้ URL parameters
    var method = e.parameter.method;
    if (method === 'insertData') {
      handleMemberDataInsertion({
        userlineId: e.parameter.userlineId,
        nameId: e.parameter.nameId,
        keynumberId: e.parameter.keynumberId,
        keynumber2Id: e.parameter.keynumber2Id,
        base64: e.parameter.base64
      });
      return ContentService.createTextOutput("Data inserted successfully").setMimeType(ContentService.MimeType.TEXT);
    } else if (method === 'updateData') {
      return doUpdate(e.parameter);
    } else if (method === 'deleteData') {
      return doDelete(e.parameter);
    }
    return ContentService.createTextOutput('Invalid request').setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    console.error('Error in doPost:', error);
    return ContentService.createTextOutput(`Error: ${error.message}`).setMimeType(ContentService.MimeType.TEXT);
  }
}

// ====================================================================================
// ฟังก์ชันดึงข้อมูล
// ====================================================================================
function fetchDataResponse() {
  var data = fetchDataFromSheet();
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function fetchDataResponseByName(nameId) {
  var data = fetchDataFromSheetByName(nameId);
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function fetchDataResponseByUserlineId(userlineId) {
  var data = fetchDataFromSheetByUserlineId(userlineId);
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function fetchDataFromSheet() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MEMBER);
  return sheet.getDataRange().getValues();
}

function fetchDataFromSheetByName(nameId) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MEMBER);
  var rows = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][COL.NAME] === nameId) {
      result.push({
        keyid: rows[i][COL.ID],
        userlineId: rows[i][COL.USER_LINE_ID],
        nameId: rows[i][COL.NAME],
        keynumberId: rows[i][COL.CODE],
        keynumber2Id: rows[i][COL.DEPARTMENT],
        fileLink: rows[i][COL.IMAGE],
        timestamp: rows[i][COL.TIMESTAMP]
      });
    }
  }
  return result;
}

function fetchDataFromSheetByUserlineId(userlineId) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MEMBER);
  var rows = sheet.getDataRange().getValues();
  var result = [];
  
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][COL.USER_LINE_ID] === userlineId) {
      var rowData = {
        keyid: rows[i][COL.ID],
        userlineId: rows[i][COL.USER_LINE_ID],
        nameId: rows[i][COL.NAME],
        keynumberId: rows[i][COL.CODE],
        keynumber2Id: rows[i][COL.DEPARTMENT],
        fileLink: rows[i][COL.IMAGE],
        timestamp: rows[i][COL.TIMESTAMP]
      };
      Logger.log('Found data: ' + JSON.stringify(rowData));
      result.push(rowData);
    }
  }
  
  Logger.log('Total data found: ' + result.length);
  return result;
}

// ====================================================================================
// ฟังก์ชันอัปเดตและลบข้อมูล
// ====================================================================================
function doUpdate(data) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MEMBER);
  var rows = sheet.getDataRange().getValues();

  var imageUrl = data.imageUrl;

  // อัปโหลดรูปใหม่ถ้าเป็น Base64 (รองรับทั้ง JPEG และ PNG)
  if (data.imageUrl && !data.imageUrl.startsWith('http')) {
    imageUrl = uploadBase64Image(data.imageUrl, data.name);
  }

  // ค้นหาและอัปเดตแถว
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][COL.ID] == data.id) {
      sheet.getRange(i + 1, COL.USER_LINE_ID + 1).setValue(data.userID);
      sheet.getRange(i + 1, COL.NAME + 1).setValue(data.name);
      sheet.getRange(i + 1, COL.CODE + 1).setValue("'" + data.employeeID);
      sheet.getRange(i + 1, COL.DEPARTMENT + 1).setValue("'" + data.department);
      sheet.getRange(i + 1, COL.IMAGE + 1).setValue(imageUrl);
      break;
    }
  }

  return ContentService.createTextOutput("Data updated").setMimeType(ContentService.MimeType.TEXT);
}

function doDelete(data) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME_MEMBER);
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][COL.ID] == data.id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return ContentService.createTextOutput("Data deleted").setMimeType(ContentService.MimeType.TEXT);
}
