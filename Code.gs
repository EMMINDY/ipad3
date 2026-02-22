// ==========================================
// 1. ส่วนตั้งค่าระบบ (CONFIGURATION)
// ==========================================

const SPREADSHEET_ID = '1Kr1GOn5F8rBNJGA7_Sqp4be7JirrRvVci0AGhtkA5hQ'; // ID ของ Google Sheet
const FOLDER_ID = '1pPtZlI8XYBle02byB5lthAhtLX8012Pa'; // ID ของ Google Drive Folder

const SHEET_NAMES = {
  STUDENTS: [
    'รายชื่อนักเรียน ม.3', 
    'รายชื่อนักเรียน ม.4', 
    'รายชื่อนักเรียน ม.5', 
    'รายชื่อนักเรียน ม.6'
  ],
  TEACHERS: 'รายชื่อครู',
  ASSETS: 'รายงานทะเบียนทรัพย์สิน', // ชื่อชีตต้องตรงเป๊ะ
  DATA_DB: 'ข้อมูล',
  LOGS: 'Log',
  ADMIN: 'แอดมิน',
  ADVISOR: 'ครูที่ปรึกษา'
};

// ==========================================
// 2. ฟังก์ชันพื้นฐาน & AI Helper
// ==========================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบตรวจสอบสถานะ iPad โรงเรียนอรัญประเทศ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

// *** ฟังก์ชันล้างชื่อขั้นสูง (Normalize) ***
// ตัดคำนำหน้า ยศทหาร และอักขระพิเศษออก เพื่อให้เทียบชื่อได้แม่นยำ
/**
 * ฟังก์ชันปรับชื่อให้สะอาดที่สุดเพื่อการจับคู่ที่แม่นยำ
 * ใช้ได้ทั้งกับชื่อนักเรียนและครู
 */
function normalizeName(name) {
  if (!name) return "";
  let n = name.toString().normalize('NFC'); // บังคับมาตรฐาน Unicode
  
  // ลบคำนำหน้า (Regex ตัวเดิมที่คุณครูใช้ ซึ่งครอบคลุมดีแล้ว)
  const titleRegex = /^(?:ว่าที่\s*ร(?:้อย)?\.?\s*[ตทพ]\.?|จ(?:่า)?\.?ส(?:ิบ)?\.?[อทต]\.?|ส(?:ิบ)?\.?[อทต]\.?|พล(?:ทหาร)?\.?|ส\.อ\.|จ\.ส\.อ\.|ร\.ต\.|ดร\.?|ผศ\.?|รศ\.?|ศ\.?|เด็กชาย|เด็กหญิง|นางสาว|ด\.?\s*ช\.?|ด\.?\s*ญ\.?|น\.?\s*ส\.?|นาย|นาง|ครู|อ\.?|mr\.?|mrs\.?|ms\.?|miss)[\s\.]*/gi;
  n = n.replace(titleRegex, ''); 

  // 🚨 ลบทุกอย่างที่ไม่ใช่ ก-ฮ, สระไทย, และ A-Z (รวมถึงช่องว่าง)
  n = n.replace(/[^ก-๙a-zA-Z]/g, ''); 
  
  return n.toLowerCase();
}

// *** ฟังก์ชันคำนวณความต่างของคำ (Levenshtein Distance) ***
// ใช้สำหรับหาชื่อที่สะกดผิดเล็กน้อย
function getEditDistance(a, b) {
  if (a.length === 0) return b.length; 
  if (b.length === 0) return a.length; 

  var matrix = [];

  // สร้างตารางเมทริกซ์
  for (var i = 0; i <= b.length; i++) { matrix[i] = [i]; }
  for (var j = 0; j <= a.length; j++) { matrix[0][j] = j; }

  // คำนวณระยะห่าง
  for (var i = 1; i <= b.length; i++) {
    for (var j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) == a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          Math.min(
            matrix[i][j - 1] + 1, // insertion
            matrix[i - 1][j] + 1  // deletion
          )
        );
      }
    }
  }
  return matrix[b.length][a.length];
}

// ==========================================
// 3. ฟังก์ชันดึงข้อมูลหลัก (MAIN DATA ENGINE)
// ==========================================

function getAllSystemData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // --- A. ดึงข้อมูล Asset (ทะเบียนทรัพย์สิน) ---
  const assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  let assetMap = {};
  let assetKeys = []; // เก็บรายชื่อที่ล้างแล้วไว้ทำ Fuzzy Match

  if (assetSheet) {
    const lastRow = assetSheet.getLastRow();
    if (lastRow > 1) {
      // อ่านข้อมูล Fast Mode (อ่านถึงคอลัมน์ J = 10 คอลัมน์)
      const assetData = assetSheet.getRange(1, 1, lastRow, 10).getValues();
      
      for (let i = 1; i < assetData.length; i++) {
        // ชื่ออยู่ที่ Col E (Index 4)
        let rawName = assetData[i][4]; 
        // Serial อยู่ที่ Col C (Index 2) หรือ D (Index 3)
        let serial = assetData[i][2] || assetData[i][3]; 

        if (rawName) {
          let cName = normalizeName(rawName); // ล้างชื่อ
          if (cName.length > 0) {
            assetMap[cName] = { 
              serial: serial ? serial.toString() : '-', 
              status: 'ยืมอยู่' 
            };
            assetKeys.push(cName); // เก็บเข้าลิสต์
          }
        }
      }
    }
  }

  // --- B. ดึงข้อมูล Log (Database) ---
  const dbSheet = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  let dbMap = {}; 
  
  if (dbSheet) {
    const dbLastRow = dbSheet.getLastRow();
    if (dbLastRow > 1) {
      // อ่านข้อมูลถึงคอลัมน์ N (14 คอลัมน์)
      const dbData = dbSheet.getRange(1, 1, dbLastRow, 14).getValues();
      
      for (let i = 1; i < dbData.length; i++) {
        let id = dbData[i][1]; // รหัสประจำตัว
        if (!id) continue;
        id = id.toString();

        let rowSerial = dbData[i][5] ? dbData[i][5].toString() : '';
        let rowStatus = dbData[i][8] ? dbData[i][8].toString() : '';
        let hasFiles = (dbData[i][9] || dbData[i][10] || dbData[i][11] || dbData[i][12]);
        
        if (!dbMap[id]) dbMap[id] = { borrowStatus: 'ยังไม่ยืม', docStatus: 'ยังไม่ส่ง', serial: '-', files: {} };
        
        if (rowSerial && rowSerial !== '-' && rowSerial !== '') dbMap[id].serial = rowSerial;
        
        if (hasFiles) {
          dbMap[id].files = { 
            agreement: dbData[i][9], card_std: dbData[i][10], 
            card_parent: dbData[i][11], house: dbData[i][12], phone: dbData[i][13] 
          };
          if (!rowStatus.includes('ADMIN') && !rowStatus.includes('ADVISOR')) {
            dbMap[id].docStatus = 'รอตรวจสอบ';
          }
        }

        if (rowStatus.includes('เอกสารผ่าน')) dbMap[id].docStatus = 'เอกสารผ่าน';
        else if (rowStatus.includes('ไม่ผ่าน')) dbMap[id].docStatus = 'เอกสารไม่ผ่าน';
        else if (rowStatus.includes('รอตรวจสอบ')) dbMap[id].docStatus = 'รอตรวจสอบ';

        if (rowStatus.includes('ยืมอยู่') || rowStatus === 'ยืมได้') dbMap[id].borrowStatus = 'ยืมอยู่';
        else if (rowStatus.includes('คืน')) dbMap[id].borrowStatus = 'คืนแล้ว';
        else if (rowStatus.includes('ซ่อม')) dbMap[id].borrowStatus = 'ส่งซ่อม';
        else if (rowStatus.includes('สละ')) dbMap[id].borrowStatus = 'สละสิทธิ์';
        else if (rowStatus === 'ยังไม่ยืม') dbMap[id].borrowStatus = 'ยังไม่ยืม';
      }
    }
  }

  // --- C. รวมข้อมูล (Merge + Fuzzy Match) ---
  let allPeople = [];
  
  const processPerson = (type, no, id, name, room, source) => {
    if (!name) return;
    id = id.toString();
    
    let cleanedName = normalizeName(name); // ล้างชื่อต้นทาง
    
    let finalBorrow = 'ยังไม่ยืม';
    let finalDoc = 'ยังไม่ส่ง';
    let finalSerial = '-';
    let finalFiles = {};
    let isInAssetSheet = false;

    // 1. ลองหาแบบ "ตรงเป๊ะ" (Exact Match) - เร็วสุด
    if (assetMap[cleanedName]) { 
      finalBorrow = assetMap[cleanedName].status; 
      finalSerial = assetMap[cleanedName].serial;
      isInAssetSheet = true;
    } else {
      // 2. ถ้าไม่เจอ -> ใช้ระบบ "Fuzzy Match" (หาความใกล้เคียง)
      for (let i = 0; i < assetKeys.length; i++) {
        let assetKey = assetKeys[i];
        
        // คำนวณความต่าง
        let dist = getEditDistance(cleanedName, assetKey);
        
        // กติกา: ชื่อยาวเกิน 5 ยอมให้ผิด 2 จุด, ถ้าน้อยกว่ายอมให้ผิด 1 จุด
        let allowedErrors = cleanedName.length > 5 ? 2 : 1;

        if (dist <= allowedErrors) {
          // เจอคู่ที่ใกล้เคียงแล้ว!
          finalBorrow = assetMap[assetKey].status;
          finalSerial = assetMap[assetKey].serial;
          isInAssetSheet = true;
          break; // เจอแล้วหยุดหา
        }
      }
    }

    // อัปเดตข้อมูลจาก Log (Database)
    if (dbMap[id]) {
      if (dbMap[id].borrowStatus !== 'ยังไม่ยืม') {
        finalBorrow = dbMap[id].borrowStatus;
      } else if (finalBorrow === 'ยังไม่ยืม' && isInAssetSheet) {
        // ถ้า DB ยังไม่ยืม แต่ Asset มีชื่อ -> ยึด Asset
        finalBorrow = 'ยืมอยู่'; 
      }
      
      finalDoc = dbMap[id].docStatus;
      if (dbMap[id].serial !== '-') finalSerial = dbMap[id].serial;
      finalFiles = dbMap[id].files;
    }

    allPeople.push({ 
      type: type, 
      no: no, 
      id: id, 
      name: name, 
      room: room, 
      source_sheet: source, 
      serial: finalSerial, 
      borrowStatus: finalBorrow, 
      docStatus: finalDoc, 
      files: finalFiles, 
      inAsset: isInAssetSheet 
    });
  };

  // วนลูปอ่านรายชื่อนักเรียน
  SHEET_NAMES.STUDENTS.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) { 
      let lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        let data = sheet.getRange(1, 1, lastRow, 4).getValues(); 
        for (let i = 1; i < data.length; i++) { 
           processPerson('student', data[i][0], data[i][1], data[i][2], data[i][3], sheetName); 
        } 
      }
    }
  });

  // วนลูปอ่านรายชื่อครู
  let teacherSheet = ss.getSheetByName(SHEET_NAMES.TEACHERS);
  if (teacherSheet) { 
    let lastRow = teacherSheet.getLastRow();
    if (lastRow > 1) {
       let tData = teacherSheet.getRange(1, 1, lastRow, 2).getValues(); 
       for (let i = 1; i < tData.length; i++) { 
         processPerson('teacher', tData[i][0], 'T-'+tData[i][0], tData[i][1], 'ห้องพักครู', SHEET_NAMES.TEACHERS); 
       } 
    }
  }

  return allPeople;
}

// ==========================================
// 4. ฟังก์ชันสำหรับระบบจับคู่ชื่อ (AUDIT & FIX)
// ==========================================

function getAssetAuditData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let peopleList = [];
  
  // 1. ดึงข้อมูลคนในระบบ (ตรวจสอบชื่อชีตและช่วงข้อมูลให้ดี)
  const fetchPeople = (sheetNames) => {
    if(!Array.isArray(sheetNames)) sheetNames = [sheetNames];
    sheetNames.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName);
      if(sheet && sheet.getLastRow() > 1) {
        // อ่านถึง Col D (No, ID, Name, Room) -> Index 0, 1, 2, 3
        let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
        data.forEach(r => {
          let originalName = r[2]; // ชื่อใน Col C
          if(originalName) { 
             let norm = normalizeName(originalName);
             if (norm) {
               peopleList.push({
                 id: r[1], 
                 name: originalName, 
                 room: r[3], 
                 sheet: sheetName, 
                 norm: norm
               });
             }
          }
        });
      }
    });
  };

  fetchPeople(SHEET_NAMES.STUDENTS);
  fetchPeople(SHEET_NAMES.TEACHERS);

  // สร้าง Map แทน Set เพื่อเก็บข้อมูลต้นฉบับไว้เทียบ
  let peopleNormMap = new Map();
  peopleList.forEach(p => peopleNormMap.set(p.norm, p.name));

  let assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  let orphans = [];

  if(assetSheet && assetSheet.getLastRow() > 1) {
    // อ่านข้อมูล Asset (ดึงเผื่อถึง Col J)
    let assetData = assetSheet.getRange(2, 1, assetSheet.getLastRow() - 1, 10).getValues(); 
    
    assetData.forEach((r, index) => {
      let assetName = r[4]; // ชื่อคนถือครองใน Col E (Index 4)
      let serial = r[2] || r[3] || 'ไม่มี Serial'; 
      
      if(assetName) {
        let assetNorm = normalizeName(assetName);
        if(!assetNorm) return;

        // --- 🚨 จุดตรวจสอบสำคัญ: ถ้ามีใน Map (Exact Match) ให้ข้ามไปเลย ---
        if(peopleNormMap.has(assetNorm)) {
          return; 
        }

        // ถ้ามาถึงตรงนี้คือ "ไม่เชื่อมโยง" (Orphan)
        let suggestions = [];
        for(let p of peopleList) {
          let dist = getEditDistance(assetNorm, p.norm);
          let isPartial = assetNorm.includes(p.norm) || p.norm.includes(assetNorm);
          
          // ปรับเกณฑ์ความคล้ายให้ยืดหยุ่นขึ้น (ผิดได้ 30% ของความยาวชื่อ)
          let threshold = Math.ceil(Math.max(assetNorm.length, p.norm.length) * 0.3);

          if(isPartial || dist <= threshold) {
            suggestions.push({
              name: p.name,
              sheet: p.sheet,
              room: p.room || '-',
              diff: isPartial ? 0 : dist,
              isPartial: isPartial
            });
          }
        }
        
        // เรียงคนที่น่าจะใช่ที่สุดขึ้นก่อน
        suggestions.sort((a,b) => (a.isPartial === b.isPartial) ? a.diff - b.diff : (a.isPartial ? -1 : 1));

        orphans.push({
          row: index + 2, // เก็บเลขแถวไว้ดูใน Sheets
          assetName: assetName,
          serial: serial.toString(),
          suggestions: suggestions.slice(0, 5) 
        });
      }
    });
  }
  return orphans;
}

// ฟังก์ชันแก้ชื่อใน "ทะเบียนทรัพย์สิน" (ตามคำสั่ง)
function adminFixAssetName(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  
  if(!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };
  
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToUpdate = -1;
    
    // วนหาแถวที่ต้องแก้ในชีตทรัพย์สิน
    for(let i=0; i<rows.length; i++) {
       // เงื่อนไข: ชื่อตรงกับชื่อเดิม (ที่ผิด) AND Serial ตรงกัน
       let currentName = rows[i][4];
       let currentSerial = rows[i][2] || rows[i][3];
       
       if(currentName == data.oldAssetName && String(currentSerial) == String(data.serial)) {
          rowToUpdate = i + 1;
          break;
       }
    }

    if(rowToUpdate > -1) {
      // แก้ไขชื่อใน Col E (คอลัมน์ที่ 5) เป็นชื่อที่ถูกต้อง
      sheet.getRange(rowToUpdate, 5).setValue(data.correctName);
      return { success: true, message: "แก้ไขชื่อในทะเบียนทรัพย์สินเรียบร้อยแล้ว" };
    } else {
      return { success: false, message: "ไม่พบข้อมูลเดิมในทะเบียนทรัพย์สิน (อาจถูกแก้ไขไปแล้ว)" };
    }
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}
// ฟังก์ชันใหม่: แก้ไขชื่อทีละหลายรายการ (Bulk Update) - ปรับปรุงใหม่
function adminFixAssetNameBulk(updateList) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  if(!sheet) return { success: false, message: "ไม่พบแผ่นงานรายงานทะเบียนทรัพย์สิน" };

  // ใช้ LockService ป้องกันการชนกันของข้อมูล
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอคิวได้สูงสุด 10 วินาที

    const dataRange = sheet.getDataRange();
    const rows = dataRange.getDisplayValues(); 
    let updateCount = 0;

    // วนลูปรายการที่จะแก้
    updateList.forEach(item => {
      // เตรียมข้อมูลเปรียบเทียบ (แปลงเป็น String และตัดช่องว่างหน้าหลังออก)
      let targetName = String(item.oldAssetName).trim();
      let targetSerial = String(item.serial).trim();

      for(let i=1; i<rows.length; i++) {
        let currentName = String(rows[i][4]).trim();           // Col E
        let currentSerial = String(rows[i][2] || rows[i][3]).trim(); // Col C or D

        // เปรียบเทียบแบบแม่นยำขึ้น
        if(currentName === targetName && currentSerial === targetSerial) {
          // อัปเดตข้อมูล (i + 1 คือแถว, 5 คือคอลัมน์ E)
          sheet.getRange(i + 1, 5).setValue(item.correctName);
          updateCount++;
          // อัปเดตค่าในตัวแปร rows ด้วย เพื่อกันพลาดกรณีชื่อซ้ำในรอบถัดไป
          rows[i][4] = item.correctName; 
          break; 
        }
      }
    });

    // *** หัวใจสำคัญ: บังคับบันทึกข้อมูลลง Sheet ทันที ***
    SpreadsheetApp.flush(); 

    return { success: true, count: updateCount };

  } catch(e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock(); // ปล่อยล็อกเสมอ
  }
}

// ==========================================
// 5. STANDARD HELPERS (Form, Auth, etc.)
// ==========================================

function processForm(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  const folder = DriveApp.getFolderById(FOLDER_ID);
  
  try {
    const timestamp = new Date();
    const uploadFile = (fileBlob, prefix) => {
      if (!fileBlob || fileBlob.name == "") return "";
      let fileName = prefix + "_" + formObject.userName + "_" + timestamp.getTime();
      return folder.createFile(fileBlob).setName(fileName).getUrl();
    };

    let url_agreement = uploadFile(formObject.file_agreement, "AGREEMENT");
    let url_card_std = "", url_card_parent = "", parent_phone = ""; // ตัด url_house ออกไป
    
    if (formObject.userType === 'student') {
      url_card_std = uploadFile(formObject.file_card_std, "CARD_STD");
      url_card_parent = uploadFile(formObject.file_card_parent, "CARD_PARENT");
      // ไม่มีการอัปโหลดทะเบียนบ้านอีกต่อไป
      parent_phone = "'" + formObject.parent_phone;
    }

    let statusToSave = formObject.statusSelect;
    if (url_agreement !== "") {
      statusToSave = statusToSave + " | รอตรวจสอบเอกสาร";
    }

    sheetData.appendRow([
      timestamp, 
      formObject.userId, 
      formObject.userName, 
      formObject.userType, 
      formObject.userRoom, 
      formObject.userSerial, 
      "USER_UPDATE", 
      formObject.note || "", 
      statusToSave, 
      url_agreement, 
      url_card_std, 
      url_card_parent, 
      "",              // คงช่องว่างไว้เพื่อรักษาตำแหน่งคอลัมน์ของทะเบียนบ้านเดิม
      parent_phone
    ]);
    return { success: true, message: "บันทึกข้อมูลเรียบร้อย" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

function verifyAdmin(u, p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADMIN);
  if (!sheet) return { success: false, message: "No Admin Sheet" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { success: true, role: 'admin' };
    }
  }
  return { success: false, message: "Login Failed" };
}

function verifyAdvisor(u, p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ADVISOR);
  if (!sheet) return { success: false, message: "ไม่พบชีตครูที่ปรึกษา" };
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(u).trim() && String(data[i][1]).trim() === String(p).trim()) {
      return { success: true, role: 'advisor', level: data[i][2], room: data[i][3], name: data[i][4] || "คุณครูที่ปรึกษา" };
    }
  }
  return { success: false, message: "ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง" };
}

function adminUpdateData(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetLog = ss.getSheetByName(SHEET_NAMES.DATA_DB); 
  
  try {
    const timestamp = new Date();
    // ตรวจสอบบทบาทผู้แก้ไขเพื่อบันทึกชื่อให้ถูกต้อง
    let editor = data.editorRole === 'advisor' ? ("ครูที่ปรึกษา: " + data.editorName) : ("ADMIN: " + (data.editorName || "ผู้ดูแลระบบ"));
    
    // 1. บันทึกประวัติลงใน Log (Sheet DATA_DB)
    if (sheetLog && data.borrowStatusSelect) {
      sheetLog.appendRow([
        timestamp, 
        data.userId, 
        data.userName, 
        data.userType, 
        data.userRoom, 
        data.userSerial, 
        editor, 
        data.note || "-", 
        data.borrowStatusSelect, 
        "", "", "", "", "" // เว้นคอลัมน์ที่เหลือ
      ]);
    }
    
    // 2. อัปเดตสถานะกลับไปที่ชีตรายชื่อชั้นเรียนหลัก (เช่น "รายชื่อนักเรียน ม.3")
    const targetSheetName = data.source_sheet; 
    if (targetSheetName) {
      const targetSheet = ss.getSheetByName(targetSheetName);
      
      // ตรวจสอบว่ามีชื่อชีตนี้อยู่จริงหรือไม่ ป้องกัน Error
      if (targetSheet) {
        const sheetValues = targetSheet.getDataRange().getValues();
        
        for (let i = 1; i < sheetValues.length; i++) {
          // ค้นหาด้วย ID (คอลัมน์ A หรือ Index 0)
          if (String(sheetValues[i][0]) === String(data.userId)) {
            
            // หมายเหตุ: เลขคอลัมน์ด้านล่างคือตำแหน่งใน Google Sheets (A=1, B=2...)
            
            // อัปเดตหมายเหตุ (คอลัมน์ H = 8)
            if (data.note !== undefined) {
              targetSheet.getRange(i + 1, 8).setValue(data.note);
            }
            
            // อัปเดตสถานะเครื่อง (คอลัมน์ K = 11)
            if (data.borrowStatusSelect) {
              targetSheet.getRange(i + 1, 11).setValue(data.borrowStatusSelect);
            }
            
            // อัปเดตสถานะเอกสาร (คอลัมน์ L = 12)
            if (data.docStatusSelect && data.docStatusSelect !== "") {
              targetSheet.getRange(i + 1, 12).setValue(data.docStatusSelect);
            }
            
            break; // เจอแล้ว อัปเดตแล้ว ให้หยุดลูปทันที
          }
        }
      } else {
        throw new Error("หาแผ่นงาน '" + targetSheetName + "' ไม่เจอในไฟล์นี้");
      }
    }

    return { success: true, message: "อัปเดตข้อมูลและสถานะใน Sheet สำเร็จ" };
  } catch (e) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() }; 
  }
}

function adminDeleteUser(data) {
  if (data.editorRole === 'advisor') return { success: false, message: "ครูที่ปรึกษาไม่ได้รับอนุญาตให้ลบข้อมูล" };
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.source_sheet);
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const rows = sheet.getDataRange().getDisplayValues();
    let rowToDelete = -1;
    for (let i = 0; i < rows.length; i++) {
      if (data.source_sheet === SHEET_NAMES.TEACHERS) { 
        if (rows[i][1] == data.name) { rowToDelete = i + 1; break; } 
      } else { 
        if (rows[i][1] == data.id) { rowToDelete = i + 1; break; } 
      }
    }
    if (rowToDelete > -1) { 
      sheet.deleteRow(rowToDelete); 
      return { success: true, message: "ลบข้อมูลเรียบร้อยแล้ว" }; 
    } else { return { success: false, message: "ไม่พบข้อมูล" }; }
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

function adminAddUser(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(data.targetSheet); 
  if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน" };
  try {
    const nextNo = sheet.getLastRow(); 
    if (data.targetSheet === SHEET_NAMES.TEACHERS) {
      sheet.appendRow([nextNo, data.name]); 
    } else {
      sheet.appendRow([nextNo, data.id, data.name, data.room]);
    }
    return { success: true, message: "เพิ่มรายชื่อเรียบร้อยแล้ว" };
  } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
}

// ==========================================
// 6. DASHBOARD STATS SYSTEM 
// ==========================================
function getDashboardStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const assetSheet = ss.getSheetByName(SHEET_NAMES.ASSETS);
  
  // ล็อคยอดทั้งหมดไว้ที่ 2085 เครื่อง
  let stats = {
    total: 2085,
    borrowed: 0,
    available: 0
  };
  
  if (assetSheet) {
    const lastRow = assetSheet.getLastRow();
    if (lastRow > 1) {
      // นับแถวที่มีชื่อในคอลัมน์ E
      const data = assetSheet.getRange(2, 5, lastRow - 1, 1).getValues();
      let count = 0;
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() !== "") {
          count++;
        }
      }
      stats.borrowed = count;
    }
  }
  
  stats.available = stats.total - stats.borrowed;
  return stats;
}
// ==========================================
// ==========================================
// ฟังก์ชันสำหรับครูที่ปรึกษาทำเรื่องคืน iPad
// ==========================================
function processAdvisorReturn(formObject) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetData = ss.getSheetByName(SHEET_NAMES.DATA_DB);
  
  // โฟลเดอร์สำหรับเก็บเอกสารการคืนเครื่อง
  const folderReturn = DriveApp.getFolderById("16Rn35Lv0gC3HRt2ohUWmi_DdQShRcUfn");
  
  try {
    const timestamp = new Date();
    let url_return = "";
    
    // อัปโหลดไฟล์เอกสาร/รูปภาพการคืน
    if (formObject.file_return && formObject.file_return.name != "") {
      // เพิ่มเลขห้องเข้าไปในชื่อไฟล์ด้วย แอดมินจะได้ค้นหาง่ายขึ้น
      let fileName = "RETURN_ห้อง" + formObject.userRoom + "_" + formObject.userName + "_" + timestamp.getTime();
      url_return = folderReturn.createFile(formObject.file_return).setName(fileName).getUrl();
    }
    
    // เปลี่ยนสถานะเป็น "อยู่ระหว่างการส่งคืน"
    let statusToSave = "อยู่ระหว่างการส่งคืน";

    // แทรกลิงก์เอกสารคืนลงในช่อง Note เพื่อให้แอดมินกดเปิดดูได้ง่ายๆ ตอนกดปุ่มแก้ไข
    let newNote = formObject.note || "";
    if (url_return !== "") {
        newNote += "\n[หลักฐานการคืน]: " + url_return;
    }

    // บันทึกลง Sheet (เปลี่ยน Action เป็น ADVISOR_RETURN)
    sheetData.appendRow([
      timestamp, 
      formObject.userId, 
      formObject.userName, 
      formObject.userType, 
      formObject.userRoom, 
      formObject.userSerial, 
      "ADVISOR_RETURN", 
      newNote, 
      statusToSave, 
      "", 
      "", 
      "", 
      "", 
      ""
    ]);
    return { success: true, message: "ส่งเรื่องคืนเรียบร้อย รอแอดมินตรวจสอบ" };
  } catch (error) { 
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}
