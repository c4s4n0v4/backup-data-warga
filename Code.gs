/**
 * DATABASE CONNECTION
 */
var SPREADSHEET_ID = '13yjgZ0-T8Ik2nnVh2Hz46ptdx_pnnUSdRSOrwka5VXs'; // Ensure this Match User ID
var TARGET_GID = 1228857705; 

function getSheet() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e2) {
      throw new Error("Gagal membuka Spreadsheet.");
    }
  }

  // Find By GID
  var sheets = ss.getSheets();
  var sheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === TARGET_GID) {
      sheet = sheets[i];
      break;
    }
  }

  // Fallback (Safe Create)
  if (!sheet) {
     sheet = ss.getSheetByName('DataWarga');
     if (!sheet) {
        sheet = ss.insertSheet('DataWarga');
        // Default Headers
        sheet.appendRow([
          'ID', 'NIK', 'Nama Lengkap', 'Jenis Kelamin', 'Tempat Lahir', 'Tanggal Lahir', 
          'Alamat', 'No HP', 'Wilayah', 'Status', 'Tanggal Baptis', 'Golongan Darah', 
          'Tanggal Kematian', 'Pendidikan Terakhir', 'Catatan'
        ]);
     }
  }
  return sheet;
}

/**
 * AUTHENTICATION & USERS SHEET
 */
function getUsersSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['Username', 'Password', 'Role', 'Wilayah']);
    // Default Super Admin
    sheet.appendRow(['admin', '123456', 'SUPER_ADMIN', 'All']);
    // Default Regional Admins (Examples)
    sheet.appendRow(['admin_barat', 'barat123', 'ADMIN_WILAYAH', 'Barat']);
    sheet.appendRow(['admin_timur', 'timur123', 'ADMIN_WILAYAH', 'Timur']);
  }
  return sheet;
}

function loginUser(username, password) {
  try {
    var sheet = getUsersSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
       var u = String(data[i][0]).trim();
       var p = String(data[i][1]).trim();
       if (u === username && p === password) {
         return {
           success: true,
           username: u,
           role: data[i][2],
           wilayah: String(data[i][3]).trim()
         };
       }
    }
    return { success: false, message: "Username/Password salah." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Data Warga GKJ Wonogiri Utara')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * DYNAMIC HEADER MAPPING
 */
function getColumnMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  
  // Normalize headers (trim, lowercase) for robust matching
  headers.forEach(function(h, index) {
    var key = String(h).toLowerCase().trim();
    map[key] = index; // Store 0-based index
  });
  
  return {
    headers: headers,
    map: map
  };
}

// Helper to get value safely by header name
function getVal(row, map, headerNames) {
  // headerNames can be an array of possible names for robust matching
  for (var i = 0; i < headerNames.length; i++) {
    var key = headerNames[i].toLowerCase().trim();
    if (map.hasOwnProperty(key)) {
      return row[map[key]];
    }
  }
  return ""; // Not found
}

function safeString(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return ""; 
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  return String(val);
}

/**
 * CRUD
 */

function getData() {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return []; // No data

    // 1. Parse Headers
    var headerInfo = getColumnMap(sheet);
    var map = headerInfo.map;

    // 2. Parse Data (skip row 0)
    
    return data.slice(1).map(function(row, index) {
      var rowNum = index + 2; // +1 for 0-index, +1 for header
      var realId = safeString(getVal(row, map, ['ID']));
      if(!realId || realId.trim() === "") {
         realId = "ROW_" + rowNum; // 1-based index (Header is 1, so row is i+2)
      }

      return {
        _row:            rowNum,
        id:              realId, 
        nik:             safeString(getVal(row, map, ['NIK'])),
        noKK:            safeString(getVal(row, map, ['No KK', 'Nomor KK', 'No. KK'])),
        nama:            safeString(getVal(row, map, ['Nama Lengkap', 'Nama'])),
        jk:              safeString(getVal(row, map, ['Jenis Kelamin', 'V_JK'])), // V_JK fallback
        hubKeluarga:     safeString(getVal(row, map, ['Hubungan Keluarga', 'Status Hubungan', 'Hubungan'])),
        tempatLahir:     safeString(getVal(row, map, ['Tempat Lahir'])),
        tanggalLahir:    safeString(getVal(row, map, ['Tanggal Lahir'])),
        namaSuami:       safeString(getVal(row, map, ['Nama Suami', 'Suami'])),
        namaIstri:       safeString(getVal(row, map, ['Nama Istri', 'Istri'])),
        namaAyah:        safeString(getVal(row, map, ['Nama Ayah', 'Ayah'])),
        namaIbu:         safeString(getVal(row, map, ['Nama Ibu', 'Ibu'])),
        alamat:          safeString(getVal(row, map, ['Alamat', 'Alamat Lengkap'])),
        noHp:            safeString(getVal(row, map, ['No HP', 'No. HP', 'Nomor HP'])),
        wilayah:         safeString(getVal(row, map, ['Wilayah'])),
        statusKeluarga:  safeString(getVal(row, map, ['Status Keluarga'])),
        statusBekerja:   safeString(getVal(row, map, ['Status Bekerja', 'Status'])),
        tanggalBaptis:   safeString(getVal(row, map, ['Tanggal Baptis', 'Tanggal Baptis/Sidi', 'Tgl Baptis'])),
        golDarah:        safeString(getVal(row, map, ['Golongan Darah', 'Gol Darah', 'Gol. Darah'])),
        tanggalKematian: safeString(getVal(row, map, ['Tanggal Kematian', 'Tgl Kematian'])),
        pendidikan:      safeString(getVal(row, map, ['Pendidikan Terakhir', 'Pendidikan'])),
        kategorial:      safeString(getVal(row, map, ['Kategorial'])),
        pekerjaan:       safeString(getVal(row, map, ['Pekerjaan'])),
        catatan:         safeString(getVal(row, map, ['Catatan', 'Keterangan']))
      };
    }).filter(function(item) {
      return item.nama !== ""; 
    });

  } catch (e) {
    console.error(e);
    throw new Error("Server Error: " + e.message); 
  }
}

function addData(formObject, userCtx) {
  try {
    // Permission Check
    if (!userCtx || !userCtx.role) throw new Error("Akses ditolak. Silakan login.");
    if (userCtx.role === 'MEMBER') throw new Error("Anggota tidak dapat menambah data.");
    if (userCtx.role === 'ADMIN_WILAYAH') {
       var fWil = superNormalize(formObject.wilayah);
       var uWil = superNormalize(userCtx.wilayah);
       if (fWil !== uWil) {
         throw new Error("Anda hanya boleh menambah data wilayah " + userCtx.wilayah);
       }
    }

    var sheet = getSheet();
    var headerInfo = getColumnMap(sheet);
    var map = headerInfo.map;
    var headers = headerInfo.headers;
    
    var id = generateCustomId(formObject.tanggalLahir, formObject.wilayah);
    
    var newRow = new Array(headers.length).fill("");
    sheet.appendRow(newRow);
    var rowNum = sheet.getLastRow();
    
    var setVal = (names, val) => {
       for(var i=0; i<names.length; i++) {
         var key = names[i].toLowerCase().trim();
         if(map.hasOwnProperty(key)) {
           var colNum = map[key] + 1;
           var cell = sheet.getRange(rowNum, colNum);
           cell.clearDataValidations(); 
           cell.setValue(val);
           return;
         }
       }
    };

    setVal(['ID'], id);
    setVal(['NIK'], "'" + formObject.nik);
    setVal(['No KK', 'No. KK'], "'" + formObject.noKK);
    setVal(['Nama Lengkap', 'Nama'], formObject.nama);
    setVal(['Jenis Kelamin'], formObject.jk);
    setVal(['Hubungan Keluarga', 'Hubungan'], formObject.hubKeluarga);
    setVal(['Tempat Lahir'], formObject.tempatLahir);
    setVal(['Tanggal Lahir'], formObject.tanggalLahir);
    setVal(['Nama Suami', 'Suami'], formObject.namaSuami);
    setVal(['Nama Istri', 'Istri'], formObject.namaIstri);
    setVal(['Nama Ayah'], formObject.namaAyah);
    setVal(['Nama Ibu'], formObject.namaIbu);
    setVal(['Alamat'], formObject.alamat);
    setVal(['No HP'], "'" + formObject.noHp);
    setVal(['Wilayah'], formObject.wilayah);
    setVal(['Status Keluarga'], formObject.statusKeluarga);
    setVal(['Status Bekerja', 'Status'], formObject.statusBekerja);
    setVal(['Tanggal Baptis'], formObject.tanggalBaptis);
    setVal(['Golongan Darah'], formObject.golDarah);
    setVal(['Tanggal Kematian'], formObject.tanggalKematian);
    setVal(['Pendidikan Terakhir', 'Pendidikan'], formObject.pendidikan);
    setVal(['Kategorial'], formObject.kategorial);
    setVal(['Pekerjaan'], formObject.pekerjaan);
    setVal(['Catatan'], formObject.catatan);
    
    return "Sukses! ID: " + id;
  } catch (e) {
    throw new Error("Error: " + e.message);
  }
}

function updateData(formObject, userCtx) {
  try {
     // Permission Check
    if (!userCtx || !userCtx.role) throw new Error("Akses ditolak. Silakan login.");
    if (userCtx.role === 'MEMBER') throw new Error("Anggota tidak dapat mengubah data.");
    
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var headerInfo = getColumnMap(sheet);
    var map = headerInfo.map;
    
    var idIndex = map['id'];
    if (idIndex === undefined) throw new Error("Kolom ID tidak ditemukan di Spreadsheet.");

    var targetId = String(formObject.id);

    for (var i = 1; i < data.length; i++) {
      var currentId = String(data[i][idIndex]);
      var rowNum = i + 1;
      
      var isRowIdMatch = targetId.startsWith("ROW_") && targetId === ("ROW_" + rowNum);
      
      if (currentId === targetId || isRowIdMatch) {
         
         // Regional Check
         if (userCtx.role === 'ADMIN_WILAYAH') {
            var wilIndex = map['wilayah'];
            var uWil = superNormalize(userCtx.wilayah || "");
             if (wilIndex !== undefined) {
                var existingWil = superNormalize(data[i][wilIndex]);
                if (existingWil !== "" && existingWil !== uWil) {
                   throw new Error("Akses Ditolak: Bukan wilayah Anda (" + existingWil + " vs " + uWil + ")");
                }
             }
             var fWil = superNormalize(formObject.wilayah || "");
             if (fWil !== uWil) throw new Error("Anda tidak boleh memindah wilayah warga.");
          }

         if (isRowIdMatch && currentId === "") { 
           var newRealId = generateCustomId(formObject.tanggalLahir, formObject.wilayah);
           var colId = map['id'] + 1;
           sheet.getRange(rowNum, colId).setValue(newRealId);
        }

        var updateCell = (names, val) => {
           for(var k=0; k<names.length; k++) {
             var key = names[k].toLowerCase().trim();
             if(map.hasOwnProperty(key)) {
               var colNum = map[key] + 1;
               var cell = sheet.getRange(rowNum, colNum);
               cell.clearDataValidations(); 
               cell.setValue(val);
               return;
             }
           }
        };

        updateCell(['NIK'], "'" + formObject.nik);
        updateCell(['No KK', 'No. KK'], "'" + formObject.noKK);
        updateCell(['Nama Lengkap', 'Nama'], formObject.nama);
        updateCell(['Jenis Kelamin'], formObject.jk);
        updateCell(['Hubungan Keluarga', 'Hubungan'], formObject.hubKeluarga);
        updateCell(['Tempat Lahir'], formObject.tempatLahir);
        updateCell(['Tanggal Lahir'], formObject.tanggalLahir);
        updateCell(['Nama Suami', 'Suami'], formObject.namaSuami);
        updateCell(['Nama Istri', 'Istri'], formObject.namaIstri);
        updateCell(['Nama Ayah'], formObject.namaAyah);
        updateCell(['Nama Ibu'], formObject.namaIbu);
        updateCell(['Alamat'], formObject.alamat);
        updateCell(['No HP'], "'" + formObject.noHp);
        updateCell(['Wilayah'], formObject.wilayah);
        updateCell(['Status Keluarga'], formObject.statusKeluarga);
        updateCell(['Status Bekerja', 'Status'], formObject.statusBekerja);
        updateCell(['Tanggal Baptis'], formObject.tanggalBaptis);
        updateCell(['Golongan Darah'], formObject.golDarah);
        updateCell(['Tanggal Kematian'], formObject.tanggalKematian);
        updateCell(['Pendidikan Terakhir', 'Pendidikan'], formObject.pendidikan);
        updateCell(['Kategorial'], formObject.kategorial);
        updateCell(['Pekerjaan'], formObject.pekerjaan);
        updateCell(['Catatan'], formObject.catatan);
        
        return "Sukses update data!";
      }
    }
    throw new Error("ID tidak ditemukan.");

  } catch (e) {
    throw new Error("Gagal: " + e.message);
  }
}


function deleteData(id, userCtx) {
  try {
     // Permission Check
    if (!userCtx || !userCtx.role) throw new Error("Akses ditolak.");
    if (userCtx.role === 'MEMBER') throw new Error("Akses ditolak.");

    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var headerInfo = getColumnMap(sheet);
    var map = headerInfo.map;
    var idIndex = map['id'];
    
    if (idIndex === undefined) throw new Error("Kolom ID tidak ditemukan.");

    for (var i = 1; i < data.length; i++) {
      var currentId = String(data[i][idIndex]);
      var rowNum = i + 1;
      var isRowIdMatch = id.startsWith("ROW_") && id === ("ROW_" + rowNum);

      if (currentId === String(id) || isRowIdMatch) {
         // Regional Check
         if (userCtx.role === 'ADMIN_WILAYAH') {
            var wilIndex = map['wilayah'];
             if (wilIndex !== undefined) {
                var existingWil = String(data[i][wilIndex]).trim();
                // Allow if matches OR if orphaned (empty Wilayah)
                if (existingWil !== "" && existingWil !== userCtx.wilayah) {
                   throw new Error("Akses Ditolak: Bukan wilayah Anda.");
                }
             }
          }

        sheet.deleteRow(rowNum);
        return "Sukses hapus!";
      }
    }
    throw new Error("ID tidak ditemukan.");
  } catch (e) {
    throw new Error("Error: " + e.message);
  }
}

function generateCustomId(birthDateString, wilayah) {
  var nowStr = new Date().getTime().toString();
  if (!birthDateString) return nowStr;

  var d = new Date(birthDateString);
  if (isNaN(d.getTime())) return nowStr;

  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  var datePrefix = y.toString() + m.toString() + day.toString();

  // Updated Map based on user request
  var mapWilayah = {
    "Barat": "01",
    "Timur": "02",
    "Tengah": "03",
    "Tandon": "04",
    "Gegeran": "05",
    "Gemantar": "06"
  };
  var wilCode = mapWilayah[wilayah] || "99";
  var prefix = datePrefix + wilCode;

  return prefix + Math.floor(Math.random() * 1000).toString().padStart(3, '0');
}
