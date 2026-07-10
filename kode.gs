function doGet(e) {
  initDatabase(); 
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('LMS Modern Edmodo - CAT ASN')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function formatDateTimeLocal(dateVal) {
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  }
  return dateVal ? dateVal.toString() : "";
}

function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsDefinition = {
    "Users": ["Username", "Password", "Role", "Nama Lengkap", "Tingkat", "Kelas", "Nomor Handphone", "Alamat Rumah", "Email", "Foto Profil"],
    "Topik": ["ID Topik", "Nama Topik", "Materi PDF URL", "Soal Diskusi", "Soal Tugas", "Waktu Rilis", "Waktu Berakhir", "Tingkat", "Rilis Pre", "Rilis Diskusi", "Rilis Materi", "Rilis Tugas", "Rilis Post"],
    "Bank Soal": ["ID Topik", "Jenis", "Pertanyaan", "Opsi A", "Opsi B", "Opsi C", "Opsi D", "Opsi E", "Kunci"],
    "Nilai Tugas": ["Username", "Nama", "Tingkat", "Kelas", "ID Topik", "Nilai Pre", "Pre Attempts", "Nilai Post", "Post Attempts", "Jawaban Diskusi", "Nilai Diskusi", "Link Tugas", "Nilai Tugas"],
    "Komentar Diskusi": ["ID Topik", "Target Username", "Penulis Username", "Penulis Nama", "Komentar"],
    "Groups": ["ID Topik", "Nama Kelompok", "Ketua", "Anggota", "FileURLs"]
  };
  
  for (var sheetName in sheetsDefinition) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheetsDefinition[sheetName]);
      insertSampleData(sheetName, sheet);
    }
  }
}

function insertSampleData(sheetName, sheet) {
  var hariIni = new Date();
  var mingguDepan = new Date();
  mingguDepan.setDate(hariIni.getDate() + 7);
  var rilisStr = Utilities.formatDate(hariIni, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  var berakhirStr = Utilities.formatDate(mingguDepan, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  
  if (sheetName === "Users") {
    sheet.appendRow(["guru1", "123", "Guru", "Budi Utomo, S.Pd", "-", "-", "", "", "", ""]);
    sheet.appendRow(["siswa1", "123", "Siswa", "Andi Wijaya", "X", "AKL 1", "", "", "", ""]);
    sheet.appendRow(["siswa2", "123", "Siswa", "Rani Saputri", "XI", "TKJ 1", "", "", "", ""]);
  } else if (sheetName === "Topik") {
    var diskusiX = JSON.stringify(["Apa perbedaan mendasar Hub dan Switch?", "Sebutkan contoh topologi fisik!"]);
    var tugasX = JSON.stringify(["Buat ringkasan topologi jaringan PDF!", "Unggah gambar desain jaringan paket tracer."]);
    sheet.appendRow(["T01", "Dasar Jaringan Komputer", "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf", diskusiX, tugasX, rilisStr, berakhirStr, "X", rilisStr, rilisStr, rilisStr, rilisStr, rilisStr]);
    
    var diskusiXI = JSON.stringify(["Sebutkan jenis-jenis Web Server!"]);
    var tugasXI = JSON.stringify(["Tugas instalasi linux (Screenshoot)."]);
    sheet.appendRow(["T02", "Administrasi Server Lanjut", "https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf", diskusiXI, tugasXI, rilisStr, berakhirStr, "XI", rilisStr, rilisStr, rilisStr, rilisStr, rilisStr]);
  } else if (sheetName === "Bank Soal") {
    sheet.appendRow(["T01", "Pre", "Apa kepanjangan dari LAN?", "Local Area Network", "Log Area Network", "Local Array Net", "Link Area Network", "List Area Network", "A"]);
    sheet.appendRow(["T01", "Pre", "Topologi berbentuk cincin adalah...", "Bus", "Star", "Ring", "Mesh", "Tree", "C"]);
    sheet.appendRow(["T01", "Post", "Perangkat pembagi cerdas di layer 2 adalah...", "Hub", "Switch", "Router", "Kabel", "Repeater", "B"]);
  }
  // Groups sheet sample kosong
}

function loginUser(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === username && data[i][1].toString() === password) {
      return { 
        success: true, 
        username: data[i][0], 
        role: data[i][2], 
        nama: data[i][3], 
        tingkat: data[i][4], 
        kelas: data[i][5],
        hp: data[i][6] || "",
        alamat: data[i][7] || "",
        email: data[i][8] || "",
        foto: data[i][9] || ""
      };
    }
  }
  return { success: false, message: "Kombinasi User & Password salah!" };
}

function updateBiodata(username, password, hp, alamat, email, fotoUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === username) {
      if (password !== "") {
        sheet.getRange(i + 1, 2).setValue(password);
      }
      sheet.getRange(i + 1, 7).setValue(hp || "");
      sheet.getRange(i + 1, 8).setValue(alamat || "");
      sheet.getRange(i + 1, 9).setValue(email || "");
      if(fotoUrl) sheet.getRange(i + 1, 10).setValue(fotoUrl);
      return "Biodata berhasil diperbarui!";
    }
  }
  return "User tidak ditemukan.";
}

function uploadProfilePicServer(base64Data, username) {
  try {
    var parts = base64Data.split(',');
    var contentType = parts[0].split(':')[1].split(';')[0];
    var decoded = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(decoded, contentType, "foto_profil_" + username);
    
    var folderName = "Foto Profil LMS";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) { folder = folders.next(); } 
    else { folder = DriveApp.createFolder(folderName); }
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: "https://drive.google.com/uc?export=view&id=" + file.getId() };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getDaftarTopik() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Topik");
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var d = data[i][3], t = data[i][4];
    var dArr, tArr;
    try { dArr = JSON.parse(d); } catch(e) { dArr = d ? [d] : []; }
    if(!Array.isArray(dArr)) dArr = d ? [d] : [];
    
    try { tArr = JSON.parse(t); } catch(e) { tArr = t ? [t] : []; }
    if(!Array.isArray(tArr)) tArr = t ? [t] : [];

    list.push({ 
      id: data[i][0], nama: data[i][1], pdf: data[i][2], 
      diskusi: dArr, tugas: tArr, 
      rilis: formatDateTimeLocal(data[i][5]), berakhir: formatDateTimeLocal(data[i][6]),
      tingkat: data[i][7] || "X",
      rilisPre: formatDateTimeLocal(data[i][8]),
      rilisDiskusi: formatDateTimeLocal(data[i][9]),
      rilisMateri: formatDateTimeLocal(data[i][10]),
      rilisTugas: formatDateTimeLocal(data[i][11]),
      rilisPost: formatDateTimeLocal(data[i][12])
    });
  }
  return list;
}

function getStatusSiswaTopik(username, idTopik) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Nilai Tugas");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][4] === idTopik) {
      var dStr = data[i][9], tStr = data[i][11];
      var dArr, tObj;
      try { dArr = JSON.parse(dStr); } catch(e) { dArr = dStr ? [dStr] : []; }
      try { tObj = JSON.parse(tStr); } catch(e) { tObj = tStr ? {"0": tStr} : {}; }

      return {
        preScore: data[i][5], preAttempts: parseInt(data[i][6]) || 0,
        postScore: data[i][7], postAttempts: parseInt(data[i][8]) || 0,
        jawabanDiskusiArr: dArr, linkTugasObj: tObj
      };
    }
  }
  return { preScore: "", preAttempts: 0, postScore: "", postAttempts: 0, jawabanDiskusiArr: [], linkTugasObj: {} };
}

function getSoal(idTopik, jenis) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bank Soal");
  var data = sheet.getDataRange().getValues();
  var soalList = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idTopik && data[i][1] === jenis) {
      var opsiAsli = [
        { key: "A", text: data[i][3] }, { key: "B", text: data[i][4] },
        { key: "C", text: data[i][5] }, { key: "D", text: data[i][6] }, { key: "E", text: data[i][7] }
      ];
      soalList.push({ pertanyaan: data[i][2], options: shuffleArray(opsiAsli), kunci: data[i][8] });
    }
  }
  return shuffleArray(soalList);
}

function shuffleArray(arr) {
  var newArr = arr.slice();
  for (var i = newArr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = newArr[i]; newArr[i] = newArr[j]; newArr[j] = temp;
  }
  return newArr;
}

// --- FUNGSI UNTUK KELOMPOK TUGAS ---

// Helper: dapatkan nama dari username
function getNamaFromUsername(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Users");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      return data[i][3]; // Nama Lengkap
    }
  }
  return username;
}

// Mendapatkan daftar siswa satu kelas yang belum memiliki kelompok untuk topik ini
function getAvailableClassmates(username, idTopik) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsers = ss.getSheetByName("Users");
  var sheetGroups = ss.getSheetByName("Groups");
  var userData = sheetUsers.getDataRange().getValues();
  var groupData = sheetGroups.getDataRange().getValues();
  
  var currentUser = null;
  var currentKelas = "";
  var currentTingkat = "";
  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0] === username) {
      currentUser = userData[i];
      currentKelas = userData[i][5];
      currentTingkat = userData[i][4];
      break;
    }
  }
  if (!currentUser) return [];
  
  var usersInGroup = {};
  for (var j = 1; j < groupData.length; j++) {
    if (groupData[j][0] === idTopik) {
      var members = groupData[j][3] ? groupData[j][3].split(',') : [];
      members.forEach(function(m) { usersInGroup[m] = true; });
      if (groupData[j][2]) usersInGroup[groupData[j][2]] = true;
    }
  }
  
  var result = [];
  for (var k = 1; k < userData.length; k++) {
    var u = userData[k];
    if (u[0] === username) continue;
    if (u[4] !== currentTingkat) continue;
    if (u[5] !== currentKelas) continue;
    if (u[2] !== "Siswa") continue;
    if (usersInGroup[u[0]]) continue;
    result.push({ username: u[0], nama: u[3] });
  }
  return result;
}

// Membuat kelompok baru
function createGroup(idTopik, leaderUsername, memberUsernames) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGroups = ss.getSheetByName("Groups");
  
  var data = sheetGroups.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idTopik) {
      var members = data[i][3] ? data[i][3].split(',') : [];
      if (data[i][2] === leaderUsername || members.indexOf(leaderUsername) !== -1) {
        return "Anda sudah memiliki kelompok untuk topik ini.";
      }
    }
  }
  
  var groupName = "Grup-" + leaderUsername + "-" + new Date().getTime();
  var allMembers = [leaderUsername].concat(memberUsernames);
  var membersStr = allMembers.join(',');
  var fileURLs = "{}";
  
  sheetGroups.appendRow([idTopik, groupName, leaderUsername, membersStr, fileURLs]);
  return "Kelompok berhasil dibuat!";
}

// Mendapatkan informasi kelompok dengan nama anggota
function getGroupInfoForTopic(idTopik, username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGroups = ss.getSheetByName("Groups");
  var data = sheetGroups.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== idTopik) continue;
    var members = data[i][3] ? data[i][3].split(',') : [];
    if (data[i][2] === username || members.indexOf(username) !== -1) {
      var membersWithNames = members.map(function(m) {
        return { username: m, nama: getNamaFromUsername(m) };
      });
      var leaderNama = getNamaFromUsername(data[i][2]);
      return {
        groupName: data[i][1],
        leader: data[i][2],
        leaderNama: leaderNama,
        members: membersWithNames,
        fileURLs: data[i][4] ? JSON.parse(data[i][4]) : {}
      };
    }
  }
  return null;
}

// Upload file untuk kelompok (izinkan semua anggota upload ulang, tapi upload pertama hanya ketua)
function uploadFileTugasGroup(base64Data, fileName, username, nama, tingkat, kelas, idTopik, taskIndex) {
  try {
    var groupInfo = getGroupInfoForTopic(idTopik, username);
    if (!groupInfo) return { success: false, error: "Anda tidak berada dalam kelompok manapun untuk topik ini." };
    
    var existingFiles = groupInfo.fileURLs || {};
    var isLeader = (groupInfo.leader === username);
    
    if (!existingFiles[taskIndex] && !isLeader) {
      return { success: false, error: "Hanya ketua yang dapat mengupload tugas pertama." };
    }
    
    var parts = base64Data.split(',');
    var contentType = parts[0].split(':')[1].split(';')[0];
    var decoded = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(decoded, contentType, fileName);
    var folderName = "Tugas Tingkat " + tingkat;
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) { folder = folders.next(); } 
    else { folder = DriveApp.createFolder(folderName); }
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetGroups = ss.getSheetByName("Groups");
    var dataGroups = sheetGroups.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < dataGroups.length; i++) {
      if (dataGroups[i][0] === idTopik && dataGroups[i][2] === groupInfo.leader) {
        rowIdx = i + 1;
        break;
      }
    }
    if (rowIdx === -1) return { success: false, error: "Gagal menemukan data kelompok." };
    
    var currentURLs = dataGroups[rowIdx-1][4] ? JSON.parse(dataGroups[rowIdx-1][4]) : {};
    currentURLs[taskIndex] = fileUrl;
    sheetGroups.getRange(rowIdx, 5).setValue(JSON.stringify(currentURLs));
    
    var sheetNilai = ss.getSheetByName("Nilai Tugas");
    var dataNilai = sheetNilai.getDataRange().getValues();
    var allMembers = groupInfo.members.map(function(m) { return m.username; });
    allMembers.push(groupInfo.leader);
    var uniqueMembers = [...new Set(allMembers)];
    
    for (var j = 0; j < uniqueMembers.length; j++) {
      var member = uniqueMembers[j];
      var foundRow = -1;
      for (var k = 1; k < dataNilai.length; k++) {
        if (dataNilai[k][0] === member && dataNilai[k][4] === idTopik) {
          foundRow = k + 1;
          break;
        }
      }
      
      var existingLinks = {};
      if (foundRow !== -1) {
        var cellVal = sheetNilai.getRange(foundRow, 12).getValue();
        try { existingLinks = JSON.parse(cellVal) || {}; } catch(e) { existingLinks = {"0": cellVal}; }
        existingLinks[taskIndex] = fileUrl;
        sheetNilai.getRange(foundRow, 12).setValue(JSON.stringify(existingLinks));
      } else {
        existingLinks[taskIndex] = fileUrl;
        var barisBaru = [member, nama, tingkat, kelas, idTopik, "", 0, "", 0, "", "", JSON.stringify(existingLinks), ""];
        sheetNilai.appendRow(barisBaru);
      }
    }
    
    return { success: true, url: fileUrl };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// --- FUNGSI LAMA (tidak diubah) ---

function uploadFileTugasIndex(base64Data, fileName, username, nama, tingkat, kelas, idTopik, taskIndex) {
  // Tetap untuk kompatibilitas jika diperlukan
  try {
    var parts = base64Data.split(',');
    var contentType = parts[0].split(':')[1].split(';')[0];
    var decoded = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(decoded, contentType, fileName);
    var folderName = "Tugas Tingkat " + tingkat;
    var folders = DriveApp.getFoldersByName(folderName);
    var folder;
    if (folders.hasNext()) { folder = folders.next(); } 
    else { folder = DriveApp.createFolder(folderName); }
    
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = file.getUrl();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Nilai Tugas");
    var data = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][4] === idTopik) { rowIdx = i + 1; break; }
    }
    
    var existingLinks = {};
    if (rowIdx !== -1) {
      var cellVal = sheet.getRange(rowIdx, 12).getValue();
      try { existingLinks = JSON.parse(cellVal) || {}; } catch(e) { existingLinks = {"0": cellVal}; }
      existingLinks[taskIndex] = fileUrl;
      sheet.getRange(rowIdx, 12).setValue(JSON.stringify(existingLinks));
    } else {
      existingLinks[taskIndex] = fileUrl;
      var barisBaru = [username, nama, tingkat, kelas, idTopik, "", 0, "", 0, "", "", JSON.stringify(existingLinks), ""];
      sheet.appendRow(barisBaru);
    }
    return { success: true, url: fileUrl };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function simpanJawabanDiskusiArray(username, nama, tingkat, kelas, idTopik, jawabanArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Nilai Tugas");
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][4] === idTopik) { rowIdx = i + 1; break; }
  }
  var val = JSON.stringify(jawabanArray);
  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 10).setValue(val);
  } else {
    var barisBaru = [username, nama, tingkat, kelas, idTopik, "", 0, "", 0, val, "", "", ""];
    sheet.appendRow(barisBaru);
  }
}

function simpanNilaiUjian(username, nama, tingkat, kelas, idTopik, jenis, nilai) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Nilai Tugas");
  var data = sheet.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][4] === idTopik) { rowIdx = i + 1; break; }
  }
  
  var scoreCol = (jenis === "Pre") ? 6 : 8;
  var attemptCol = (jenis === "Pre") ? 7 : 9;
  if (rowIdx !== -1) {
    var currentAttempts = parseInt(sheet.getRange(rowIdx, attemptCol).getValue()) || 0;
    sheet.getRange(rowIdx, scoreCol).setValue(nilai);
    sheet.getRange(rowIdx, attemptCol).setValue(currentAttempts + 1);
  } else {
    var barisBaru = [username, nama, tingkat, kelas, idTopik, "", 0, "", 0, "", "", "", ""];
    barisBaru[scoreCol - 1] = nilai;
    barisBaru[attemptCol - 1] = 1;
    sheet.appendRow(barisBaru);
  }
}

function getNilaiSiswaForGuru() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Nilai Tugas");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    list.push({
      rowNum: i + 1, username: data[i][0], nama: data[i][1], tingkat: data[i][2], kelas: data[i][3], idTopik: data[i][4], 
      nilaiPre: data[i][5], preAttempts: data[i][6], nilaiPost: data[i][7], postAttempts: data[i][8],
      jawabanDiskusi: data[i][9], nilaiDiskusi: data[i][10], linkTugas: data[i][11], nilaiTugas: data[i][12]
    });
  }
  return list;
}

function getRekapNilaiSiswa(username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNilai = ss.getSheetByName("Nilai Tugas");
  var sheetTopik = ss.getSheetByName("Topik");
  if (!sheetNilai) return [];
  
  var dataNilai = sheetNilai.getDataRange().getValues();
  var dataTopik = sheetTopik ? sheetTopik.getDataRange().getValues() : [];
  
  var mapTopik = {};
  for (var j = 1; j < dataTopik.length; j++) {
    mapTopik[dataTopik[j][0]] = dataTopik[j][1];
  }
  
  var list = [];
  for (var i = 1; i < dataNilai.length; i++) {
    if(dataNilai[i][0] === username) {
      var idT = dataNilai[i][4];
      var namaT = mapTopik[idT] ? mapTopik[idT] : idT;
      list.push({ namaTopik: namaT, pre: dataNilai[i][5], post: dataNilai[i][7], nDiskusi: dataNilai[i][10], nTugas: dataNilai[i][12] });
    }
  }
  return list;
}

function simpanNilaiEsaiDariGuru(rowNum, colType, skor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Nilai Tugas");
  var colIdx = (colType === 'diskusi') ? 11 : 13; 
  sheet.getRange(rowNum, colIdx).setValue(skor);
  return "Nilai Berhasil Disimpan!";
}

function tambahAtauUpdateTopik(id, nama, tingkat, rilis, berakhir, rilisPre, rilisDiskusi, rilisMateri, rilisTugas, rilisPost) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Topik");
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  
  for(var i=1; i<data.length; i++) {
    if(data[i][0] === id) { foundRow = i + 1; break; }
  }
  
  if(foundRow !== -1) {
    sheet.getRange(foundRow, 2).setValue(nama);
    sheet.getRange(foundRow, 6, 1, 8).setValues([[rilis, berakhir, tingkat, rilisPre, rilisDiskusi, rilisMateri, rilisTugas, rilisPost]]);
    return "Topik Berhasil Diperbarui!";
  } else {
    var defDis = JSON.stringify(["Masukkan soal diskusi di sini..."]);
    var defTgs = JSON.stringify(["Masukkan perintah tugas di sini..."]);
    sheet.appendRow([id, nama, "", defDis, defTgs, rilis, berakhir, tingkat, rilisPre, rilisDiskusi, rilisMateri, rilisTugas, rilisPost]);
    return "Topik Baru Berhasil Ditambahkan!";
  }
}

function hapusTopikDariSheet(idTopik) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Topik");
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] === idTopik) {
      sheet.deleteRow(i + 1);
      return "Topik berhasil dihapus.";
    }
  }
  return "Topik tidak ditemukan.";
}

function getSoalGuru(idTopik, jenis) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bank Soal");
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idTopik && data[i][1] === jenis) {
      list.push({ rowNum: i + 1, pertanyaan: data[i][2], a: data[i][3], b: data[i][4], c: data[i][5], d: data[i][6], e: data[i][7], kunci: data[i][8] });
    }
  }
  return list;
}

function tambahSoalGuru(idTopik, jenis, q, a, b, c, d, e, kunci) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bank Soal");
  sheet.appendRow([idTopik, jenis, q, a, b, c, d, e, kunci]);
  return "Soal berhasil ditambahkan!";
}

function editSoalGuruBerdasarkanRow(rowNum, idTopik, jenis, q, a, b, c, d, e, kunci) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bank Soal");
  sheet.getRange(rowNum, 1, 1, 9).setValues([[idTopik, jenis, q, a, b, c, d, e, kunci]]);
  return "Soal berhasil diperbarui!";
}

function hapusSoalGuruBerdasarkanRow(rowNum) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Bank Soal");
  sheet.deleteRow(rowNum);
  return "Soal berhasil dihapus!";
}

function updateElementTopikSpesifik(idTopik, tipe, nilaiStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Topik");
  var data = sheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][0] === idTopik) {
      var colIdx = (tipe === 'Materi') ? 3 : ((tipe === 'Diskusi') ? 4 : 5);
      sheet.getRange(i + 1, colIdx).setValue(nilaiStr);
      return "Komponen " + tipe + " Berhasil Diperbarui!";
    }
  }
  return "Topik tidak ditemukan.";
}

function getDiskusiKelas(idTopik, tingkat, kelas) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNilai = ss.getSheetByName("Nilai Tugas");
  var sheetKomentar = ss.getSheetByName("Komentar Diskusi");
  var dataNilai = sheetNilai.getDataRange().getValues();
  var dataKomentar = sheetKomentar ? sheetKomentar.getDataRange().getValues() : [];
  var listDiskusi = [];
  for (var i = 1; i < dataNilai.length; i++) {
    if (dataNilai[i][4] === idTopik && dataNilai[i][2] === tingkat && dataNilai[i][3] === kelas) {
      if (dataNilai[i][9] && dataNilai[i][9].toString().trim() !== "") {
        var usernameTarget = dataNilai[i][0];
        var kList = [];
        
        for (var j = 1; j < dataKomentar.length; j++) {
          if (dataKomentar[j][0] === idTopik && dataKomentar[j][1] === usernameTarget) {
            kList.push({ penulisNama: dataKomentar[j][3], komentar: dataKomentar[j][4] });
          }
        }
        
        var jStr = dataNilai[i][9];
        var jArr = [];
        try { jArr = JSON.parse(jStr); } catch(e) { jArr = [jStr]; }
        
        listDiskusi.push({ username: usernameTarget, nama: dataNilai[i][1], jawabanArr: jArr, komentarList: kList });
      }
    }
  }
  return listDiskusi;
}

function submitKomentarDiskusi(idTopik, targetUsername, penulisUsername, penulisNama, komentar) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Komentar Diskusi");
  sheet.appendRow([idTopik, targetUsername, penulisUsername, penulisNama, komentar]);
  return "Komentar diskusi terkirim!";
}
// === FUNGSI BARU UNTUK MENAMPILKAN STATUS KELOMPOK SISWA ===
function getClassmatesWithGroupStatus(username, idTopik) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetUsers = ss.getSheetByName("Users");
  var sheetGroups = ss.getSheetByName("Groups");
  var userData = sheetUsers.getDataRange().getValues();
  var groupData = sheetGroups.getDataRange().getValues();

  // Cari data user yang login
  var currentUser = null;
  var currentKelas = "";
  var currentTingkat = "";
  for (var i = 1; i < userData.length; i++) {
    if (userData[i][0] === username) {
      currentUser = userData[i];
      currentKelas = userData[i][5];
      currentTingkat = userData[i][4];
      break;
    }
  }
  if (!currentUser) return [];

  // Buat map: username -> { hasGroup, groupName, leader }
  var groupStatus = {};
  for (var j = 1; j < groupData.length; j++) {
    if (groupData[j][0] !== idTopik) continue;
    var members = groupData[j][3] ? groupData[j][3].split(',') : [];
    var leader = groupData[j][2];
    var groupName = groupData[j][1];
    // Tambahkan leader
    groupStatus[leader] = { hasGroup: true, groupName: groupName, leader: leader };
    // Tambahkan anggota
    members.forEach(function(m) {
      groupStatus[m] = { hasGroup: true, groupName: groupName, leader: leader };
    });
  }

  var result = [];
  for (var k = 1; k < userData.length; k++) {
    var u = userData[k];
    if (u[0] === username) continue; // exclude diri sendiri
    if (u[4] !== currentTingkat) continue;
    if (u[5] !== currentKelas) continue;
    if (u[2] !== "Siswa") continue;

    var status = groupStatus[u[0]] || { hasGroup: false, groupName: "", leader: "" };
    result.push({
      username: u[0],
      nama: u[3],
      hasGroup: status.hasGroup,
      groupName: status.groupName,
      leader: status.leader
    });
  }
  return result;
}
function leaveGroup(idTopik, username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGroups = ss.getSheetByName("Groups");
  var sheetNilai = ss.getSheetByName("Nilai Tugas");
  var dataGroups = sheetGroups.getDataRange().getValues();
  var foundRow = -1;
  var members = [];
  var leader = "";
  
  // Cari baris kelompok berdasarkan idTopik dan cek apakah username ada di anggota
  for (var i = 1; i < dataGroups.length; i++) {
    if (dataGroups[i][0] !== idTopik) continue;
    var m = dataGroups[i][3] ? dataGroups[i][3].split(',') : [];
    if (dataGroups[i][2] === username) {
      return "Ketua tidak dapat keluar dari kelompok. Hapus kelompok atau pindahkan ketua.";
    }
    if (m.indexOf(username) !== -1) {
      foundRow = i + 1;
      leader = dataGroups[i][2];
      members = m;
      break;
    }
  }
  if (foundRow === -1) return "Anda tidak berada dalam kelompok untuk topik ini.";
  
  // Hapus username dari daftar anggota
  var newMembers = members.filter(function(m) { return m !== username; });
  if (newMembers.length === 0) {
    // Jika tidak ada anggota tersisa, hapus baris kelompok (atau biarkan kosong)
    sheetGroups.deleteRow(foundRow);
  } else {
    sheetGroups.getRange(foundRow, 4).setValue(newMembers.join(','));
  }
  
  // Hapus link tugas siswa dari sheet Nilai Tugas (reset linkTugas dan nilaiTugas)
  var dataNilai = sheetNilai.getDataRange().getValues();
  for (var j = 1; j < dataNilai.length; j++) {
    if (dataNilai[j][0] === username && dataNilai[j][4] === idTopik) {
      var rowNum = j + 1;
      sheetNilai.getRange(rowNum, 12).setValue("{}"); // Link Tugas
      sheetNilai.getRange(rowNum, 13).setValue("");   // Nilai Tugas
      break;
    }
  }
  
  return "Anda berhasil keluar dari kelompok. File tugas telah dihapus.";
}
function leaveGroupAdvanced(idTopik, username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGroups = ss.getSheetByName("Groups");
  var sheetNilai = ss.getSheetByName("Nilai Tugas");
  var dataGroups = sheetGroups.getDataRange().getValues();
  var foundRow = -1;
  var members = [];
  var leader = "";
  var groupName = "";
  
  // Cari baris kelompok
  for (var i = 1; i < dataGroups.length; i++) {
    if (dataGroups[i][0] !== idTopik) continue;
    var m = dataGroups[i][3] ? dataGroups[i][3].split(',') : [];
    var l = dataGroups[i][2];
    // Cek apakah username adalah leader atau anggota
    if (l === username || m.indexOf(username) !== -1) {
      foundRow = i + 1;
      leader = l;
      members = m;
      groupName = dataGroups[i][1];
      break;
    }
  }
  if (foundRow === -1) return "Anda tidak berada dalam kelompok untuk topik ini.";
  
  // Jika yang keluar adalah leader, pilih anggota pertama sebagai leader baru (jika ada)
  if (leader === username) {
    if (members.length === 0) {
      // Tidak ada anggota, hapus kelompok
      sheetGroups.deleteRow(foundRow);
    } else {
      // Anggota pertama menjadi leader baru
      var newLeader = members[0];
      // Hapus newLeader dari daftar anggota
      members.splice(0, 1);
      // Update leader dan anggota
      sheetGroups.getRange(foundRow, 3).setValue(newLeader);
      sheetGroups.getRange(foundRow, 4).setValue(members.join(','));
    }
  } else {
    // Anggota keluar: hapus dari daftar anggota
    var newMembers = members.filter(function(m) { return m !== username; });
    if (newMembers.length === 0 && leader !== username) {
      // Tidak ada anggota tersisa selain leader, hapus kelompok? 
      // Atau biarkan leader tanpa anggota? Lebih baik hapus kelompok.
      sheetGroups.deleteRow(foundRow);
    } else {
      sheetGroups.getRange(foundRow, 4).setValue(newMembers.join(','));
    }
  }
  
  // Hapus file tugas dari siswa yang keluar (reset linkTugas dan nilaiTugas)
  var dataNilai = sheetNilai.getDataRange().getValues();
  for (var j = 1; j < dataNilai.length; j++) {
    if (dataNilai[j][0] === username && dataNilai[j][4] === idTopik) {
      var rowNum = j + 1;
      sheetNilai.getRange(rowNum, 12).setValue("{}"); // Link Tugas
      sheetNilai.getRange(rowNum, 13).setValue("");   // Nilai Tugas
      break;
    }
  }
  
  return "Anda berhasil keluar dari kelompok. File tugas Anda telah dihapus.";
}
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var args = body.args || [];
    var result = this[action].apply(this, args);
    return ContentService.createTextOutput(JSON.stringify({ data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
