// --- KONFIGURASI UTAMA ---
const SHEET_ID = '1fqAdBg-IzmFe-JLVcLxd5d6rFl04YANf75JpkHCHjdo'; // ID Spreadsheet Anda
const SHEET_NAME = 'outbound SAT';
const FOLDER_ID = '1Ll4-DKAnOOLKPa_zmRmFkxQun24vcJnM'; // ID Folder Drive Anda

function setupPermissions() {
  DriveApp.getFolderById(FOLDER_ID);
  SpreadsheetApp.openById(SHEET_ID);
}

// Helper: Ambil atau Buat Tab Users
function getOrCreateUsersSheet(ss) {
  let sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Username", "Email", "Role", "Date Created"]);
  }
  return sheet;
}

// --- FUNGSI STANDAR (GET) ---
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  var action = e.parameter.action;
  
  if (action === 'getRiwayat') {
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      
      // Tanggal mentah dari Vercel
      const rawTarget = e.parameter.target; 
      let targetDmy = "";
      let targetYmd = "";

      // LOGIKA ANTI-MELESET (Menerjemahkan semua jenis format tanggal)
      if (!rawTarget) {
          targetDmy = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
          targetYmd = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
      } else if (rawTarget.includes('/')) {
          targetDmy = rawTarget;
          const p = rawTarget.split('/');
          if (p.length === 3) targetYmd = p[2] + "-" + p[1] + "-" + p[0];
      } else if (rawTarget.includes('-')) {
          targetYmd = rawTarget;
          const p = rawTarget.split('-');
          if (p.length === 3) targetDmy = p[2] + "/" + p[1] + "/" + p[0];
      }
      
      const rows = sheet.getDataRange().getDisplayValues();
      if (rows.length <= 1) return output.setContent(JSON.stringify({ status: 'ok', data: [] }));

      const data = rows.slice(1).map((row, index) => ({
        rowIndex:   index + 2,
        timestamp:  row[0],
        group:      row[1],
        area:       row[2],
        nopol:      row[3],
        chiller:    parseInt(row[4]) || 0,
        freezer:    parseInt(row[5]) || 0,
        coklat:     parseInt(row[6]) || 0,
        dus:        parseInt(row[7]) || 0,
        sterofoam:  parseInt(row[8]) || 0,
        total:      parseInt(row[9]) || 0,
        qc:         row[10],
        loader:     row[11],
        keterangan: row[12],
        photoLink:  row[13],
        tanggal:    row[14],
        jamMuat:    row[15] || "",
        jamSelesai: row[16] || ""
      })).filter(row => {
        if (targetYmd && row.tanggal && row.tanggal.includes(targetYmd)) return true;
        if (targetDmy && row.tanggal && row.tanggal.includes(targetDmy)) return true;
        if (targetDmy && row.timestamp && row.timestamp.includes(targetDmy)) return true;
        if (targetYmd && row.timestamp && row.timestamp.includes(targetYmd)) return true;
        return false;
      }).reverse();

      output.setContent(JSON.stringify({ status: 'ok', data: data }));
    } catch(err) {
      output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
    }
  } 
  
  else if (action === 'getAllData') {
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      const rows = sheet.getDataRange().getDisplayValues();
      if (rows.length <= 1) return output.setContent(JSON.stringify({ status: 'ok', data: [] }));

      const data = rows.slice(1).map((row, index) => ({
        rowIndex:   index + 2,
        timestamp:  row[0],
        group:      row[1],
        area:       row[2],
        nopol:      row[3],
        chiller:    parseInt(row[4]) || 0,
        freezer:    parseInt(row[5]) || 0,
        coklat:     parseInt(row[6]) || 0,
        dus:        parseInt(row[7]) || 0,
        sterofoam:  parseInt(row[8]) || 0,
        total:      parseInt(row[9]) || 0,
        qc:         row[10],
        loader:     row[11],
        keterangan: row[12],
        photoLink:  row[13],
        tanggal:    row[14],
        jamMuat:    row[15] || "",
        jamSelesai: row[16] || ""
      })).reverse();

      output.setContent(JSON.stringify({ status: 'ok', data: data }));
    } catch(err) {
      output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
    }
  }
  
  else if (action === 'getUsers') {
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = getOrCreateUsersSheet(ss);
      const rows = sheet.getDataRange().getDisplayValues();
      if (rows.length <= 1) return output.setContent(JSON.stringify({ status: 'ok', data: [] }));

      const data = rows.slice(1).map((row, index) => ({
        rowIndex: index + 2,
        username: row[0],
        email: row[1],
        role: row[2],
        dateCreated: row[3]
      }));

      output.setContent(JSON.stringify({ status: 'ok', data: data }));
    } catch(err) {
      output.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
    }
  }

  return output;
}

// --- FUNGSI UTAMA (POST) UNTUK MENERIMA DATA ---
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.payload;

    if (action === 'uploadFoto') {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const bytes = Utilities.base64Decode(payload.base64Data);
      const blob = Utilities.newBlob(bytes, 'image/jpeg', payload.fileName);
      
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      output.setContent(JSON.stringify({ status: "ok", url: file.getUrl() }));
      return output;

    } else if (action === 'simpanData') {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
      
      const rowData = [
        timestamp,                    // A - Timestamp
        payload.group,                // B - Group
        payload.area,                 // C - Area
        payload.nopol,                // D - Nopol
        payload.chiller,              // E - Chiller
        payload.freezer,              // F - Freezer
        payload.coklat,               // G - Coklat
        payload.dus || 0,             // H - Dus
        payload.sterofoam || 0,       // I - Sterofoam
        payload.total,                // J - Total
        payload.qc,                   // K - QC
        payload.loader,               // L - Loader
        payload.keterangan,           // M - Keterangan
        payload.photoLink || "",      // N - Photo Link
        payload.tanggal,              // O - Tanggal Carian
        payload.jamMuat || "",        // P - Jam Muat
        payload.jamSelesai || ""      // Q - Jam Selesai Muat
      ];

      if (payload.rowIndex && payload.rowIndex !== "") {
        sheet.getRange(parseInt(payload.rowIndex), 1, 1, 17).setValues([rowData]);
        output.setContent(JSON.stringify({ status: "ok", message: "Data Diperbarui!" }));
      } else {
        sheet.appendRow(rowData);
        output.setContent(JSON.stringify({ status: "ok", message: "Data Disimpan!" }));
      }
      return output;
    } 
    
    else if (action === 'deleteRecord') {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
      const rowIndex = parseInt(payload.rowIndex);
      if (rowIndex >= 2) {
        sheet.deleteRow(rowIndex);
        output.setContent(JSON.stringify({ status: "ok", message: "Data Berhasil Dihapus!" }));
      } else {
        output.setContent(JSON.stringify({ status: "error", message: "Baris tidak valid!" }));
      }
      return output;
    }
    
    else if (action === 'saveUser') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = getOrCreateUsersSheet(ss);
      const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
      
      const rowData = [
        payload.username,
        payload.email,
        payload.role || "petugas",
        payload.dateCreated || timestamp
      ];
      
      if (payload.rowIndex && payload.rowIndex !== "") {
        sheet.getRange(parseInt(payload.rowIndex), 1, 1, 4).setValues([rowData]);
        output.setContent(JSON.stringify({ status: "ok", message: "User Diperbarui!" }));
      } else {
        sheet.appendRow(rowData);
        output.setContent(JSON.stringify({ status: "ok", message: "User Disimpan!" }));
      }
      return output;
    }
    
    else if (action === 'deleteUser') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = getOrCreateUsersSheet(ss);
      const rowIndex = parseInt(payload.rowIndex);
      if (rowIndex >= 2) {
        sheet.deleteRow(rowIndex);
        output.setContent(JSON.stringify({ status: "ok", message: "User Berhasil Dihapus!" }));
      } else {
        output.setContent(JSON.stringify({ status: "error", message: "Baris tidak valid!" }));
      }
      return output;
    }

  } catch (error) {
    output.setContent(JSON.stringify({ status: "error", message: error.toString() }));
    return output;
  }
}

function doOptions(e) {
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}
