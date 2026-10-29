// ============================================================
//  DomiCha Contract Backend — Google Apps Script
//  วิธีใช้:
//    1. เปิด script.google.com → New Project → วางโค้ดนี้
//    2. รัน setupSheetColumns() ครั้งเดียวเพื่อเพิ่ม Column ใหม่
//    3. Deploy → New Deployment → Web App
//       Execute as: Me | Who has access: Anyone
//    4. คัดลอก Web App URL ไปใส่ใน HTML form
// ============================================================

const CONFIG = {
  SHEET_ID:         '117qoLgTd1LJnBFwzVcaGVeYiuG0-AXN621pzmtH7eBQ',
  SHEET_NAME:       'ContractData',
  PARENT_FOLDER_ID: '1gecUuzj0JCIBmpxpKwz27VfgNReGNL41'  // โฟลเดอร์ สัญญา_
};

// ── Column mapping (1-indexed) ──────────────────────────────
const COL = {
  BRANCH_NO:          1,   // A
  CUSTOMER_NAME:      2,   // B
  ADDRESS:            3,   // C
  PHONE:              4,   // D
  ID_CARD:            5,   // E
  BRANCH_NAME:        6,   // F
  PROVINCE:           7,   // G
  FRANCHISE_FEE:      8,   // H
  CONTRACT_DATE:      9,   // I
  START_DATE:         10,  // J
  LINK_LOCATION:      11,  // K
  STATUS:             12,  // L
  DOC_FILE_ID:        13,  // M  (ต้นฉบับ)
  COPY_FILE_ID:       14,  // N  (คู่ฉบับ)
  FOLDER_LINK:        15,  // O
  DATE_GENERATED:     16,  // P
  DATE_SENT:          17,  // Q
  DATE_SIGNED:        18,  // R
  DATE_PAID:          19,  // S
  DATE_TRAINING:      20,  // T
  DATE_OPENED:        21,  // U
  NOTES:              22,  // V
  CANCEL_TYPE:        23,  // W  ← NEW (BY_FRANCHISEE | BY_FRANCHISOR)
  CANCEL_DATE:        24,  // X  ← NEW
  CANCEL_REASON:      25,  // Y  ← NEW
  CANCEL_CONTRACT_ID: 26   // Z  ← NEW (File ID ของสัญญายกเลิก)
};

// ── Entry points ────────────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'createContract':  return createContract(data);
      case 'updateStatus':    return updateStatus(data);
      case 'cancelContract':  return cancelContract(data);
      case 'getBranchInfo':   return getBranchInfo(data);
      case 'getAllBranches':  return getAllBranches(data);
      default: return jsonErr('Unknown action: ' + data.action);
    }
  } catch (err) {
    return jsonErr(err.toString());
  }
}

function doGet() {
  return ContentService
    .createTextOutput('✅ DomiCha Contract API is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Action: สร้างสัญญาใหม่ ──────────────────────────────────

function createContract(data) {
  // 1. สร้าง folder: DMCxxxx_BranchName
  const parent = DriveApp.getFolderById(CONFIG.PARENT_FOLDER_ID);
  const folderName = data.branchNo + '_' + data.branchName;
  const folder = parent.createFolder(folderName);

  // 2. บันทึกไฟล์ ต้นฉบับ
  const origFile = saveDocx(folder,
    'สัญญาแฟรนไชส์โดมิชา_' + data.branchNo + '_ต้นฉบับ.docx',
    data.docxOriginal);

  // 3. บันทึกไฟล์ คู่ฉบับ
  const copyFile = saveDocx(folder,
    'สัญญาแฟรนไชส์โดมิชา_' + data.branchNo + '_คู่ฉบับ.docx',
    data.docxCopy);

  // 4. เพิ่ม row ใน Sheet
  const sheet = getSheet();
  const today = thaiDate();

  sheet.appendRow([
    data.branchNo,        // A: เลขสาขา
    data.customerName,    // B: CustomerName
    data.address,         // C: Address
    data.phone,           // D: Phone
    data.idCard,          // E: IDCard
    data.branchName,      // F: BranchName
    data.province,        // G: Province
    data.franchiseFee,    // H: FranchiseFee
    data.contractDate,    // I: ContractDate
    data.startDate,       // J: StartDate
    '',                   // K: Link Location
    'GENERATED',          // L: Status
    origFile.getId(),     // M: DocFileId (ต้นฉบับ)
    copyFile.getId(),     // N: CopyFileId (คู่ฉบับ)
    folder.getUrl(),      // O: FolderLink
    today,                // P: DateGenerated
    '', '', '', '', '', '', // Q-V: dates + notes (blank)
    '', '', '', ''         // W-Z: cancel fields (blank)
  ]);

  return jsonOk({
    folderUrl:    folder.getUrl(),
    folderId:     folder.getId(),
    origFileId:   origFile.getId(),
    copyFileId:   copyFile.getId(),
    dateCreated:  today
  });
}

// ── Action: อัปเดต Status ───────────────────────────────────

function updateStatus(data) {
  // data = { branchNo, status, notes? }
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  const today = thaiDate();

  // Status → date column mapping
  const dateCols = {
    SENT:     COL.DATE_SENT,
    SIGNED:   COL.DATE_SIGNED,
    PAID:     COL.DATE_PAID,
    TRAINING: COL.DATE_TRAINING,
    ACTIVE:   COL.DATE_OPENED
  };

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(data.branchNo).trim()) {
      const r = i + 1; // 1-indexed row
      sheet.getRange(r, COL.STATUS).setValue(data.status);
      if (dateCols[data.status]) {
        sheet.getRange(r, dateCols[data.status]).setValue(today);
      }
      if (data.notes) {
        sheet.getRange(r, COL.NOTES).setValue(data.notes);
      }
      return jsonOk({ updatedRow: r, status: data.status, date: today });
    }
  }

  return jsonErr('ไม่พบสาขา: ' + data.branchNo);
}

// ── Action: ยกเลิกสัญญา ─────────────────────────────────────
//
//  data = {
//    action:        'cancelContract',
//    branchNo:      'DMC0001',
//    cancelType:    'BY_FRANCHISEE' | 'BY_FRANCHISOR',
//    cancelDate:    '01/01/2568',   // วันที่ยกเลิก (Thai format)
//    cancelReason:  'ข้อความ...',
//    cancelDocBase64: '<base64 string>'  // optional: สัญญายกเลิก DOCX
//  }

function cancelContract(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();

  const newStatus = data.cancelType === 'BY_FRANCHISEE'
    ? 'CANCELLED_BY_FRANCHISEE'
    : 'CANCELLED_BY_FRANCHISOR';

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(data.branchNo).trim()) {
      const r          = i + 1;
      const folderLink = rows[i][COL.FOLDER_LINK - 1];
      let   cancelFileId = '';

      // บันทึกสัญญายกเลิกลงใน folder เดิม (ถ้ามี)
      if (data.cancelDocBase64 && folderLink) {
        try {
          // ดึง folder จาก FolderLink column
          const folderId = extractFolderIdFromUrl(folderLink);
          if (folderId) {
            const folder = DriveApp.getFolderById(folderId);
            const suffix = data.cancelType === 'BY_FRANCHISEE'
              ? 'แฟรนไชส์ซียกเลิก' : 'แฟรนไชส์ซอร์ยกเลิก';
            const cancelFile = saveDocx(folder,
              'สัญญายกเลิก_' + data.branchNo + '_' + suffix + '.docx',
              data.cancelDocBase64);
            cancelFileId = cancelFile.getId();
          }
        } catch (fe) {
          Logger.log('⚠️ ไม่สามารถบันทึกสัญญายกเลิกลง Drive: ' + fe.toString());
        }
      }

      // อัปเดต Sheet
      sheet.getRange(r, COL.STATUS).setValue(newStatus);
      sheet.getRange(r, COL.CANCEL_TYPE).setValue(data.cancelType);
      sheet.getRange(r, COL.CANCEL_DATE).setValue(data.cancelDate || thaiDate());
      sheet.getRange(r, COL.CANCEL_REASON).setValue(data.cancelReason || '');
      if (cancelFileId) {
        sheet.getRange(r, COL.CANCEL_CONTRACT_ID).setValue(cancelFileId);
      }

      // ไฮไลต์ row สีแดงอ่อน
      try {
        const lastCol = sheet.getLastColumn();
        sheet.getRange(r, 1, 1, lastCol).setBackground('#ffcdd2');
      } catch (ce) { /* ข้ามถ้า error */ }

      return jsonOk({
        updatedRow:    r,
        status:        newStatus,
        cancelFileId:  cancelFileId,
        cancelDate:    data.cancelDate || thaiDate()
      });
    }
  }

  return jsonErr('ไม่พบสาขา: ' + data.branchNo);
}

// ── Action: ดึงข้อมูลสาขา ────────────────────────────────────
//  data = { action: 'getBranchInfo', branchNo: 'DMC0001' }

function getBranchInfo(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(data.branchNo).trim()) {
      const row = rows[i];
      return jsonOk({
        branchNo:         row[COL.BRANCH_NO - 1],
        customerName:     row[COL.CUSTOMER_NAME - 1],
        branchName:       row[COL.BRANCH_NAME - 1],
        province:         row[COL.PROVINCE - 1],
        status:           row[COL.STATUS - 1],
        folderLink:       row[COL.FOLDER_LINK - 1],
        dateGenerated:    row[COL.DATE_GENERATED - 1],
        dateSent:         row[COL.DATE_SENT - 1],
        dateSigned:       row[COL.DATE_SIGNED - 1],
        datePaid:         row[COL.DATE_PAID - 1],
        dateTraining:     row[COL.DATE_TRAINING - 1],
        dateOpened:       row[COL.DATE_OPENED - 1],
        notes:            row[COL.NOTES - 1],
        cancelType:       row[COL.CANCEL_TYPE - 1],
        cancelDate:       row[COL.CANCEL_DATE - 1],
        cancelReason:     row[COL.CANCEL_REASON - 1],
        cancelContractId: row[COL.CANCEL_CONTRACT_ID - 1]
      });
    }
  }

  return jsonErr('ไม่พบสาขา: ' + data.branchNo);
}

// ── Action: ดึงข้อมูลสาขาทั้งหมด (Dashboard) ─────────────────
//  data = { action: 'getAllBranches' }
//  Returns: array of all branch rows

function getAllBranches(data) {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  const branches = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[COL.BRANCH_NO - 1]) continue; // skip empty rows
    branches.push({
      branchNo:      String(row[COL.BRANCH_NO - 1]).trim(),
      customerName:  row[COL.CUSTOMER_NAME - 1],
      branchName:    row[COL.BRANCH_NAME - 1],
      province:      row[COL.PROVINCE - 1],
      franchiseFee:  row[COL.FRANCHISE_FEE - 1],
      contractDate:  row[COL.CONTRACT_DATE - 1],
      startDate:     row[COL.START_DATE - 1],
      status:        row[COL.STATUS - 1],
      dateGenerated: row[COL.DATE_GENERATED - 1],
      dateSent:      row[COL.DATE_SENT - 1],
      dateSigned:    row[COL.DATE_SIGNED - 1],
      datePaid:      row[COL.DATE_PAID - 1],
      dateTraining:  row[COL.DATE_TRAINING - 1],
      dateOpened:    row[COL.DATE_OPENED - 1],
      notes:         row[COL.NOTES - 1],
      cancelType:    row[COL.CANCEL_TYPE - 1],
      cancelDate:    row[COL.CANCEL_DATE - 1],
      cancelReason:  row[COL.CANCEL_REASON - 1],
      folderLink:    row[COL.FOLDER_LINK - 1]
    });
  }

  return jsonOk({ branches: branches, total: branches.length });
}

// ── Setup: รันครั้งเดียวเพื่อเพิ่ม Column ใหม่ใน Sheet ──────

function setupSheetColumns() {
  const sheet   = getSheet();
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
                        .getValues()[0];

  // Phase 1: O-V (original columns)
  const phase1Headers = [
    'FolderLink',    // O (15)
    'DateGenerated', // P (16)
    'DateSent',      // Q (17)
    'DateSigned',    // R (18)
    'DatePaid',      // S (19)
    'DateTraining',  // T (20)
    'DateOpened',    // U (21)
    'Notes'          // V (22)
  ];

  // Phase 2: W-Z (cancel columns)
  const phase2Headers = [
    'CancelType',       // W (23)
    'CancelDate',       // X (24)
    'CancelReason',     // Y (25)
    'CancelContractId'  // Z (26)
  ];

  let added = 0;

  phase1Headers.forEach(function(header, i) {
    const col = 15 + i;
    if (!existing[col - 1] || existing[col - 1] === '') {
      const cell = sheet.getRange(1, col);
      cell.setValue(header);
      cell.setFontWeight('bold');
      cell.setBackground('#c8e6c9');
      added++;
    }
  });

  phase2Headers.forEach(function(header, i) {
    const col = 23 + i;
    if (!existing[col - 1] || existing[col - 1] === '') {
      const cell = sheet.getRange(1, col);
      cell.setValue(header);
      cell.setFontWeight('bold');
      cell.setBackground('#ffcdd2'); // สีชมพูสำหรับ cancel columns
      added++;
    }
  });

  // ทำให้ Status column โดดเด่น
  sheet.getRange(1, COL.STATUS).setBackground('#fff176');

  Logger.log('✅ เพิ่ม ' + added + ' columns เรียบร้อย (O–Z)');
  SpreadsheetApp.flush();
}

// ── Helpers ─────────────────────────────────────────────────

function saveDocx(folder, filename, base64) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    filename
  );
  return folder.createFile(blob);
}

function getSheet() {
  return SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);
}

function thaiDate() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy');
}

function extractFolderIdFromUrl(url) {
  // https://drive.google.com/drive/folders/FOLDER_ID
  const m = String(url).match(/folders\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

function jsonOk(obj) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data: obj }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
