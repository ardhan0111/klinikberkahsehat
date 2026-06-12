// Ganti ID ini dengan ID spreadsheet Google Sheets kamu
// ID bisa dilihat dari URL spreadsheet:
// https://docs.google.com/spreadsheets/d/[ID_DI_SINI]/edit
const SPREADSHEET_ID = '1YV6l3mqFhr7r6AGhbAApQd11v37HKuHUeqTmXtH4F7E';

// Coba getActiveSpreadsheet dulu (jika container-bound),
// jika gagal gunakan openById (jika standalone)
function getSpreadsheet() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch(e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
const SS = getSpreadsheet();

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Klinik Berkah Sehat')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- FUNGSI AMBIL DATA ---
function getDataFromSheet(sheetName) {
  try {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return [];
    const range = sheet.getDataRange();
    const data = range.getValues();
    if (data.length <= 1) return [];

    const headers = data.shift();
    return data.map((row, index) => {
      let obj = { rowNumber: index + 2 };
      headers.forEach((header, i) => {
        let value = row[i];
        if (value instanceof Date) {
          // Tahun < 1900 → format Jam (serial time di Sheets)
          if (value.getFullYear() < 1900) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
          } else {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
        }
        obj[header] = value;
      });
      return obj;
    });
  } catch (e) {
    console.log("Error pada sheet " + sheetName + ": " + e);
    return [];
  }
}

// --- AUTH ---
// Sheet Users: ID_User | Username | Password | Role | ID_Pasien_Linked
// Password bisa tersimpan sebagai angka (misal 1234) atau string — gunakan String() untuk komparasi
function checkLogin(username, password) {
  const users = getDataFromSheet("Users");
  const user = users.find(u =>
    String(u.Username).trim().toLowerCase() === String(username).trim().toLowerCase() &&
    String(u.Password).trim() === String(password).trim()
  );
  if (user) return { success: true, role: user.Role, nama: user.Username };
  return { success: false, msg: "Username atau Password salah!" };
}

// --- CEK PASIEN ---
// Sheet Pasien: ID_Pasien | Nama | NIK | JK | Gol_Darah | TTL | Alamat | No_WA | Pekerjaan | Kontak_Darurat
// Sheet RekamMedis: ID_Rekam | ID_Pasien | Tanggal | Keluhan | Diagnosa | Tindakan | Catatan_Fisio | ID_Terapis
// Sheet Jadwal: ID_Jadwal | ID_Pasien | Tanggal_Terapi | Jam | Status | Status_Pengajuan | ...
function getPasienPublicData(idPasien) {
  const semuaPasien = getDataFromSheet("Pasien");
  const pasien = semuaPasien.find(p => String(p.ID_Pasien).toUpperCase() === idPasien.toUpperCase());
  if (!pasien) return { success: false, msg: "ID Pasien tidak ditemukan!" };

  const rmData = getDataFromSheet("RekamMedis").filter(rm => String(rm.ID_Pasien) === String(pasien.ID_Pasien));
  const jadwalData = getDataFromSheet("Jadwal").filter(j =>
    String(j.ID_Pasien) === String(pasien.ID_Pasien) && j.Status !== "Selesai"
  );

  return { success: true, profil: pasien, riwayat: rmData, jadwal: jadwalData };
}

// --- CRUD ADMIN ---
function getManagementData() {
  return {
    pasien: getDataFromSheet("Pasien"),
    staff: getDataFromSheet("Users")
  };
}

function hapusData(sheetName, rowNumber) {
  try {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return "Sheet tidak ditemukan!";
    sheet.deleteRow(parseInt(rowNumber));
    return "Data berhasil dihapus.";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// Sheet Pasien kolom: ID_Pasien | Nama | NIK | JK | Gol_Darah | TTL | Alamat | No_WA | Pekerjaan | Kontak_Darurat | Jenis_Terapi | Terapis
function tambahPasienBackend(d) {
  try {
    const sheet = SS.getSheetByName("Pasien");
    if (!sheet) return "Error: Sheet Pasien tidak ditemukan!";
    
    // Auto-create headers if missing
    if (sheet.getLastColumn() < 12) {
      sheet.getRange(1, 11).setValue("Jenis_Terapi");
      sheet.getRange(1, 12).setValue("Terapis");
    }
    
    const nextId = "PT-" + (2500 + sheet.getLastRow());
    const ttl = d.tempat + ", " + d.tglLahir;
    sheet.appendRow([nextId, d.nama, d.nik, d.jk, d.goldar, ttl, d.alamat, d.wa, d.pekerjaan, d.darurat, d.jenisTerapi || "", d.terapis || ""]);
    return nextId;
  } catch(e) {
    return "Error: " + e.message;
  }
}

// Sheet Users kolom: ID_User | Username | Password | Role | ID_Pasien_Linked
function tambahStaffBackend(d) {
  try {
    const sheet = SS.getSheetByName("Users");
    if (!sheet) return { success: false, msg: "Sheet Users tidak ditemukan!" };
    const nextId = "USR-" + (1000 + sheet.getLastRow());
    sheet.appendRow([nextId, d.username, d.password, d.role, ""]);
    return { success: true, msg: "Staff " + d.username + " berhasil ditambahkan." };
  } catch(e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

// Sheet Pengaturan_Web: ID_Konten | Nama_Klinik | Deskripsi_Sistem | Alamat_Klinik | Kontak_WA | Email_Klinik
// Data ada di baris 2, kolom 2-6
function updateBerandaBackend(d) {
  try {
    const sheet = SS.getSheetByName("Pengaturan_Web");
    if (!sheet) return { success: false, msg: "Sheet Pengaturan_Web tidak ditemukan!" };
    // Kolom: B=Nama_Klinik, C=Deskripsi_Sistem, D=Alamat_Klinik, E=Kontak_WA, F=Email_Klinik
    sheet.getRange(2, 2, 1, 4).setValues([[d.namaKlinik, d.deskripsi, d.alamat, d.kontak]]);
    return { success: true };
  } catch(e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

// --- REKAM MEDIS & JADWAL ---
// Sheet RekamMedis: ID_Rekam | ID_Pasien | Tanggal | Keluhan | Diagnosa | Tindakan | Catatan_Fisio | ID_Terapis
function simpanRekamMedisBackend(d) {
  try {
    const rmSheet = SS.getSheetByName("RekamMedis");
    if (!rmSheet) return "Error: Sheet RekamMedis tidak ditemukan!";

    const now = new Date();
    const tglSekarang = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const nextRmId = "RM-" + (1000 + rmSheet.getLastRow());

    rmSheet.appendRow([
      nextRmId,       // ID_Rekam
      d.idPasien,     // ID_Pasien
      tglSekarang,    // Tanggal
      d.keluhan,      // Keluhan
      d.diagnosa,     // Diagnosa
      d.tindakan,     // Tindakan
      d.catatan,      // Catatan_Fisio
      d.terapis       // ID_Terapis (nama terapis)
    ]);

    // Simpan ke sheet Jadwal jika ada tanggal & jam berikutnya
    // Sheet Jadwal: ID_Jadwal | ID_Pasien | Tanggal_Terapi | Jam | ID_Terapis | Status | Status_Pengajuan | ...
    if (d.tglNext && d.jamNext) {
      const jadwalSheet = SS.getSheetByName("Jadwal");
      if (jadwalSheet) {
        const nextJadwalId = "JDW-" + (1000 + jadwalSheet.getLastRow());
        jadwalSheet.appendRow([
          nextJadwalId,   // ID_Jadwal (Col 1)
          d.idPasien,     // ID_Pasien (Col 2)
          d.tglNext,      // Tanggal_Terapi (Col 3)
          d.jamNext,      // Jam (Col 4)
          d.terapis || "",// ID_Terapis (Col 5)
          "Terjadwal",    // Status (Col 6)
          "",             // Status_Pengajuan (Col 7)
          "",             // Request_Tgl_Baru (Col 8)
          ""              // Request_Jam_Baru (Col 9)
        ]);
      }
    }

    return "Rekam medis berhasil disimpan!";
  } catch(e) {
    return "Error: " + e.message;
  }
}

// --- AMBIL INFO KLINIK ---
// Sheet Pengaturan_Web baris 2: ID_Konten | Nama_Klinik | Deskripsi_Sistem | Alamat_Klinik | Kontak_WA | Email_Klinik
function getInfoKlinik() {
  try {
    const sheet = SS.getSheetByName("Pengaturan_Web");
    if (!sheet) return {};
    const row = sheet.getRange(2, 1, 1, 6).getValues()[0];
    return {
      namaKlinik: row[1] || "Klinik Berkah Sehat",
      deskripsi: row[2] || "",
      alamat: row[3] || "",
      kontak: row[4] || "",
      email: row[5] || ""
    };
  } catch(e) {
    return {};
  }
}

// --- DASHBOARD DATA ---
function getDashboardData() {
  // Auto-create Pasien headers if missing
  try {
    const pSheet = SS.getSheetByName("Pasien");
    if (pSheet && pSheet.getLastColumn() < 12) {
      pSheet.getRange(1, 11).setValue("Jenis_Terapi");
      pSheet.getRange(1, 12).setValue("Terapis");
    }
  } catch(e) {}

  const pasien = getDataFromSheet("Pasien");
  const jadwal = getDataFromSheet("Jadwal");
  const rekamMedis = getDataFromSheet("RekamMedis");
  const users = getDataFromSheet("Users");

  const now = new Date();
  const todayA = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const todayB = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const jadwalHariIni = jadwal.filter(function(j) {
    const t = String(j.Tanggal_Terapi).trim();
    return t === todayA || t === todayB;
  });
  const selesai = jadwalHariIni.filter(function(j) { return j.Status === "Selesai"; }).length;
  const antrian = jadwalHariIni.filter(function(j) { return j.Status === "Terjadwal"; }).length;

  const pasienMap = {};
  pasien.forEach(function(p) { pasienMap[String(p.ID_Pasien)] = p; });

  const lastTerapi = {};
  rekamMedis.forEach(function(rm) {
    lastTerapi[String(rm.ID_Pasien)] = { tindakan: rm.Tindakan || "-", terapis: rm.ID_Terapis || "-", tanggal: rm.Tanggal };
  });

  const pasienList = pasien.map(function(p) {
    const lt = lastTerapi[String(p.ID_Pasien)];
    return {
      ID_Pasien: p.ID_Pasien, Nama: p.Nama, NIK: p.NIK, JK: p.JK, No_WA: p.No_WA,
      Alamat: p.Alamat, Pekerjaan: p.Pekerjaan, Gol_Darah: p.Gol_Darah, TTL: p.TTL,
      jenisTerapi: lt ? lt.tindakan : (p.Jenis_Terapi || p["Jenis Terapi"] || p.jenisTerapi || "-"),
      terapis: lt ? lt.terapis : (p.Terapis || p.terapis || "-")
    };
  });

  const jadwalList = jadwalHariIni.map(function(j) {
    const p = pasienMap[String(j.ID_Pasien)];
    const lt = lastTerapi[String(j.ID_Pasien)];
    return {
      ID_Jadwal: j.ID_Jadwal, ID_Pasien: j.ID_Pasien,
      Nama: p ? p.Nama : j.ID_Pasien, Jam: j.Jam, Status: j.Status,
      jenisTerapi: lt ? lt.tindakan : "-", terapis: lt ? lt.terapis : "-"
    };
  }).sort(function(a, b) { return String(a.Jam).localeCompare(String(b.Jam)); });

  const recentActivity = rekamMedis.slice(-5).reverse().map(function(rm) {
    const p = pasienMap[String(rm.ID_Pasien)];
    return { type: "rekam_medis", nama: p ? p.Nama : rm.ID_Pasien, tindakan: rm.Tindakan, terapis: rm.ID_Terapis, tanggal: rm.Tanggal };
  });

  const terapis = users.filter(function(u) { return u.Role === "Terapis"; });

  return {
    totalPasien: pasien.length, sesiHariIni: jadwalHariIni.length,
    antrianHariIni: antrian, selesaiHariIni: selesai,
    pasien: pasienList, jadwalHariIni: jadwalList,
    recentActivity: recentActivity, totalTerapis: terapis.length, totalRM: rekamMedis.length
  };
}

// --- LAPORAN DATA ---
function getLaporanData() {
  const pasien = getDataFromSheet("Pasien");
  const rekamMedis = getDataFromSheet("RekamMedis");
  const jadwal = getDataFromSheet("Jadwal");

  const now = new Date();
  const todayA = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const todayB = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const jadwalHariIni = jadwal.filter(function(j) {
    const t = String(j.Tanggal_Terapi).trim();
    return t === todayA || t === todayB;
  });
  const selesaiHariIni = jadwalHariIni.filter(function(j) { return j.Status === "Selesai"; }).length;
  const antrianHariIni = jadwalHariIni.filter(function(j) { return j.Status === "Terjadwal"; }).length;

  const terapiCount = {};
  rekamMedis.forEach(function(rm) {
    const t = rm.Tindakan || "Lainnya";
    terapiCount[t] = (terapiCount[t] || 0) + 1;
  });

  const selesaiAll = jadwal.filter(function(j) { return j.Status === "Selesai"; }).length;
  const rasio = jadwal.length > 0 ? Math.round((selesaiAll / jadwal.length) * 1000) / 10 : 0;

  let maxCount = 0, terfavorit = "-";
  for (const key in terapiCount) {
    if (terapiCount[key] > maxCount) { maxCount = terapiCount[key]; terfavorit = key; }
  }

  return {
    totalPasien: pasien.length, totalRM: rekamMedis.length,
    totalJadwal: jadwal.length, selesaiAll: selesaiAll,
    selesaiHariIni: selesaiHariIni, antrianHariIni: antrianHariIni,
    rasioKehadiran: rasio, terapiBreakdown: terapiCount, terfavorit: terfavorit
  };
}

// --- TERAPIS DATA ---
function getTerapisData() {
  const users = getDataFromSheet("Users");
  const rekamMedis = getDataFromSheet("RekamMedis");

  const terapis = users.filter(function(u) { return u.Role === "Terapis"; });

  return terapis.map(function(t) {
    const rmCount = rekamMedis.filter(function(rm) { return rm.ID_Terapis === t.Username; }).length;
    return { ID_User: t.ID_User, Username: t.Username, Role: t.Role, totalSesi: rmCount };
  });
}

// --- ANTRIAN DATA ---
function getAntrianData() {
  const jadwal = getDataFromSheet("Jadwal");
  const pasien = getDataFromSheet("Pasien");
  const rekamMedis = getDataFromSheet("RekamMedis");

  const now = new Date();
  const todayA = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const todayB = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const pasienMap = {};
  pasien.forEach(function(p) { pasienMap[String(p.ID_Pasien)] = p; });
  const lastTerapi = {};
  rekamMedis.forEach(function(rm) {
    lastTerapi[String(rm.ID_Pasien)] = { tindakan: rm.Tindakan, terapis: rm.ID_Terapis };
  });

  const antrianHariIni = jadwal.filter(function(j) {
    const t = String(j.Tanggal_Terapi).trim();
    return t === todayA || t === todayB;
  }).map(function(j) {
    const p = pasienMap[String(j.ID_Pasien)];
    const lt = lastTerapi[String(j.ID_Pasien)];
    return {
      ID_Jadwal: j.ID_Jadwal, ID_Pasien: j.ID_Pasien,
      Nama: p ? p.Nama : j.ID_Pasien, Jam: j.Jam, Status: j.Status,
      jenisTerapi: lt ? lt.tindakan : "-", terapis: lt ? lt.terapis : "-"
    };
  }).sort(function(a, b) { return String(a.Jam).localeCompare(String(b.Jam)); });

  return {
    antrian: antrianHariIni, total: antrianHariIni.length,
    menunggu: antrianHariIni.filter(function(a) { return a.Status === "Terjadwal"; }).length,
    selesai: antrianHariIni.filter(function(a) { return a.Status === "Selesai"; }).length
  };
}

// --- UPDATE STATUS JADWAL ---
function updateJadwalStatus(idJadwal, status) {
  try {
    const sheet = SS.getSheetByName("Jadwal");
    if (!sheet) return { success: false, msg: "Sheet Jadwal tidak ditemukan!" };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idJadwal)) {
        sheet.getRange(i + 1, 6).setValue(status); // Status berada di kolom 6 (F)
        return { success: true, msg: "Status antrian berhasil diperbarui." };
      }
    }
    return { success: false, msg: "ID Jadwal tidak ditemukan." };
  } catch(e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

// --- TAMBAH JADWAL LANGSUNG ---
function tambahJadwalBackend(d) {
  try {
    const jadwalSheet = SS.getSheetByName("Jadwal");
    if (!jadwalSheet) return { success: false, msg: "Sheet Jadwal tidak ditemukan!" };
    const nextId = "JDW-" + (1000 + jadwalSheet.getLastRow());
    jadwalSheet.appendRow([nextId, d.idPasien, d.tanggal, d.jam, "", "Terjadwal", "", "", ""]); // Kolom 5 kosong, Kolom 6 status
    return { success: true, msg: "Jadwal berhasil ditambahkan.", id: nextId };
  } catch(e) {
    return { success: false, msg: "Error: " + e.message };
  }
}

// --- API WEB ENDPOINT (CORS & VERCEL SUPPORT) ---
function doPost(e) {
  try {
    // Membaca data kiriman JSON dari luar
    const requestData = JSON.parse(e.postData.contents);
    const fnName = requestData.fnName;
    const args = requestData.args || [];
    let result;

    // Routing pemanggilan fungsi berdasarkan nama fungsi
    if (fnName === "checkLogin") {
      result = checkLogin(args[0], args[1]);
    } else if (fnName === "getPasienPublicData") {
      result = getPasienPublicData(args[0]);
    } else if (fnName === "getManagementData") {
      result = getManagementData();
    } else if (fnName === "hapusData") {
      result = hapusData(args[0], args[1]);
    } else if (fnName === "tambahPasienBackend") {
      result = tambahPasienBackend(args[0]);
    } else if (fnName === "tambahStaffBackend") {
      result = tambahStaffBackend(args[0]);
    } else if (fnName === "updateBerandaBackend") {
      result = updateBerandaBackend(args[0]);
    } else if (fnName === "simpanRekamMedisBackend") {
      result = simpanRekamMedisBackend(args[0]);
    } else if (fnName === "getDataFromSheet") {
      result = getDataFromSheet(args[0]);
    } else if (fnName === "getInfoKlinik") {
      result = getInfoKlinik();
    } else if (fnName === "getDashboardData") {
      result = getDashboardData();
    } else if (fnName === "getLaporanData") {
      result = getLaporanData();
    } else if (fnName === "getTerapisData") {
      result = getTerapisData();
    } else if (fnName === "getAntrianData") {
      result = getAntrianData();
    } else if (fnName === "updateJadwalStatus") {
      result = updateJadwalStatus(args[0], args[1]);
    } else if (fnName === "tambahJadwalBackend") {
      result = tambahJadwalBackend(args[0]);
    } else {
      throw new Error("Fungsi '" + fnName + "' tidak ditemukan di backend.");
    }

    // Mengembalikan hasil sukses
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    // Mengembalikan hasil error jika terjadi kegagalan
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
