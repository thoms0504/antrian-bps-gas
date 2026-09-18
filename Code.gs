// ============================================================
// PST BPS ANTRIAN — BACKUP SYSTEM (Google Apps Script)
// Database: Google Sheets
// ============================================================
// KONFIGURASI: Isi SPREADSHEET_ID dengan ID spreadsheet Anda
// Dapatkan ID dari URL spreadsheet:
//   https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
// ============================================================

const SPREADSHEET_ID = "{EDIT ID SPREADSHEET ANDA}-tc";
const SHEET_QUEUE    = "Queue";
const SHEET_CONFIG   = "Config";
const SHEET_USERS    = "Users";
const SHEET_JADWAL_BULANAN = "JadwalBulanan"; // roster bulanan: siapa jadi Petugas/Resepsionis bulan ini
const SHEET_JADWAL_HARIAN  = "JadwalHarian";  // dari roster bulanan, siapa yg on-duty hari ini (dipilih resepsionis)
const MAX_PETUGAS_PER_BULAN = 3;
const MAX_RESEPSIONIS_PER_BULAN = 1;

// ─── MAIN ROUTER ──────────────────────────────────────────
function doGet(e) {
  const page = (e.parameter.page || "kiosk").toLowerCase();
  let template;
  if (page === "form")           template = HtmlService.createTemplateFromFile("Form");
  else if (page === "dashboard") template = HtmlService.createTemplateFromFile("Dashboard");
  else if (page === "tiket")     {
    template = HtmlService.createTemplateFromFile("Tiket");
    template.entryId = e.parameter.id || "";
  }
  else                           template = HtmlService.createTemplateFromFile("Kiosk");

  return template.evaluate()
    .setTitle("PST BPS — Sistem Antrian")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ─── HELPERS ─────────────────────────────────────────────
function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) sheet.appendRow(headers);
  }
  return sheet;
}

// Format tanggal sebagai string "yyyy-MM-dd" — SELALU gunakan fungsi ini
function todayStr() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd");
}

// Format waktu sebagai string "yyyy-MM-dd HH:mm:ss"
function nowStr() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
}

function generateId() {
  return Utilities.getUuid().replace(/-/g, "").substring(0, 12).toUpperCase();
}

// ─── KONVERSI TANGGAL DARI SHEETS ───────────────────────
// Google Sheets kadang otomatis mengkonversi string tanggal menjadi Date object
// Gunakan fungsi ini untuk membaca kolom tanggal dengan aman
function safeDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM-dd");
  }
  return String(val).trim().substring(0, 10); // ambil 10 karakter pertama "yyyy-MM-dd"
}

// Sama seperti safeDateStr, tapi untuk kolom "Periode" (format "yyyy-MM", tanpa hari).
// Google Sheets kadang mengkonversi string "2026-08" menjadi objek Date (1 Agustus 2026),
// sehingga perbandingan String(val) langsung akan selalu gagal cocok. Gunakan fungsi ini
// setiap kali membaca/mencocokkan kolom Periode di sheet JadwalBulanan.
function safePeriodeStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM");
  }
  return String(val).trim().substring(0, 7);
}

// ─── INISIALISASI SPREADSHEET ────────────────────────────
function setupSpreadsheet() {
  getOrCreateSheet(SHEET_QUEUE, [
    "ID","Ticket","Prefix","Number","NamaPengunjung","Tujuan","Detail",
    "InstansiAlamat","JumlahOrang","Status","Petugas","ServiceType",
    "ServiceDetail","Rating","RatingComment","Timestamp","Tanggal"
  ]);
  const configSheet = getOrCreateSheet(SHEET_CONFIG, ["Key","Value"]);
  if (configSheet.getLastRow() <= 1) {
    configSheet.appendRow(["KANTOR_NAMA", "BPS Provinsi Lampung"]);
    configSheet.appendRow(["KANTOR_KODE", "BPS-LAMPUNG"]);
    configSheet.appendRow(["COUNTER_P", "0"]);
    configSheet.appendRow(["COUNTER_DATE", todayStr()]);
  }
  getOrCreateSheet(SHEET_USERS, ["ID","Nama","Username","Password","Role","IsActive"]);
  getOrCreateSheet(SHEET_JADWAL_BULANAN, ["ID","Periode","PegawaiID","Nama","Role","IsActive"]);
  ensureUsersExtraColumns_();         // Users + kolom FotoUrl (foto profil)
  ensureJadwalHarianExtraColumns_(); // JadwalHarian + kolom IsDutyModeOn (Mode Bertugas)
  ensurePingSheet_();                // PingNotification (fitur Ping/Ready)
  ensureServiceLogSheet_();          // QueueServiceLog (fitur Multi-Petugas)
  ensureQueueExtraColumns();         // Queue + kolom CallSeq/MulaiLayan/Selesai/Email/NoHP
  return { ok: true, message: "Spreadsheet berhasil diinisialisasi." };
}

// Format periode bulan sebagai "yyyy-MM" — dipakai untuk rotasi jadwal bulanan
function currentPeriode() {
  return Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM");
}

// ─── PASTIKAN KOLOM TAMBAHAN ADA DI SHEET QUEUE ──────────
// Menambahkan kolom CallSeq, MulaiLayanTimestamp, SelesaiTimestamp
// jika belum ada (aman dijalankan berkali-kali & di spreadsheet lama).
function ensureQueueExtraColumns() {
  const sheet = getOrCreateSheet(SHEET_QUEUE, [
    "ID","Ticket","Prefix","Number","NamaPengunjung","Tujuan","Detail",
    "InstansiAlamat","JumlahOrang","Status","Petugas","ServiceType",
    "ServiceDetail","Rating","RatingComment","Timestamp","Tanggal"
  ]);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0].map(String);
  const need = ["CallSeq", "MulaiLayanTimestamp", "SelesaiTimestamp", "Email", "NoHP"];
  const toAdd = need.filter(function(h) { return headers.indexOf(h) === -1; });
  if (toAdd.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, toAdd.length).setValues([toAdd]);
  }
  return sheet;
}

// Ambil index kolom (0-based) berdasarkan nama header, -1 jika tidak ada
function getHeaderIndex(headers, name) {
  return headers.indexOf(name);
}

// ─── AMBIL URL WEB APP RESMI ─────────────────────────────
function getAppUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    return { ok: true, url: url };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── CONFIG HELPERS ──────────────────────────────────────
function getConfig(key) {
  const sheet = getSheet(SHEET_CONFIG);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return String(data[i][1]);
  }
  return null;
}

function setConfig(key, value) {
  const sheet = getSheet(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) { sheet.getRange(i + 1, 2).setValue(value); return; }
  }
  sheet.appendRow([key, value]);
}

// ─── GENERATE NOMOR TIKET (PER JENIS LAYANAN: A/B/C/D/U) ─
// Setiap prefix (A, B, C, D, U) punya counter & reset harian sendiri-sendiri.
// Contoh: A-001, A-002 (Perpustakaan), lalu B-001 (Konsultasi Statistik), dst.
function getNextTicketNumber(prefix) {
  const p = String(prefix || "U").toUpperCase();
  const today = todayStr();
  const counterKey = "COUNTER_" + p;
  const dateKey = "COUNTER_DATE_" + p;
  const counterDate = safeDateStr(getConfig(dateKey));
  if (counterDate !== today) {
    setConfig(counterKey, "0");
    setConfig(dateKey, today);
  }
  const next = parseInt(getConfig(counterKey) || "0") + 1;
  setConfig(counterKey, String(next));
  return next;
}

// ─── REGISTRASI PENGUNJUNG ───────────────────────────────
function submitRegistration(formData) {
  try {
    // ── Tentukan prefix nomor antrian ──
    // Kunjungan PST: prefix mengikuti Jenis Layanan PST yang dipilih (A/B/C/D).
    // Kunjungan RAPAT / LAINNYA (tanpa sub-layanan): pakai prefix "U" (Umum).
    let prefix = "U";
    if (formData.purpose === "PST" && formData.pstSubType) {
      prefix = String(formData.pstSubType).toUpperCase();
    }

    const number = getNextTicketNumber(prefix);
    const ticket = prefix + "-" + String(number).padStart(3, "0"); // contoh: A-001, B-012
    const id = generateId();
    // ensureQueueExtraColumns() menjamin urutan kolom: ...Tanggal, CallSeq, MulaiLayanTimestamp,
    // SelesaiTimestamp, Email, NoHP — dipakai juga sebagai getOrCreateSheet untuk Queue.
    const sheet = ensureQueueExtraColumns();

    // ⚠️ PENTING: Simpan tanggal sebagai string agar tidak dikonversi oleh Google Sheets
    const tanggalStr = todayStr();   // "2026-08-28"
    const timestampStr = nowStr();   // "2026-08-28 19:30:00"

    sheet.appendRow([
      id,
      ticket,
      prefix,
      number,
      formData.visitorName || "",
      formData.purpose || "PST",
      formData.visitDetail || "",
      formData.addressInstansi || "",
      parseInt(formData.visitorCount) || 1,
      "WAITING",
      "",
      formData.pstSubType || "",    // Kode sub-layanan PST (A/B/C/D)
      formData.pstSubLabel || "",   // Label sub-layanan PST
      "",
      "",
      timestampStr,
      tanggalStr,
      "", "", "",                   // CallSeq, MulaiLayanTimestamp, SelesaiTimestamp (diisi belakangan)
      formData.email || "",         // Email (opsional)
      formData.phone || ""          // NoHP (opsional)
    ]);

    return {
      ok: true,
      ticket: ticket,
      ticketNumber: number,
      id: id,
      kantorNama: getConfig("KANTOR_NAMA") || "BPS"
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── AMBIL DATA ANTRIAN HARI INI ─────────────────────────
function getQueueToday() {
  try {
    const sheet = ensureQueueExtraColumns();
    if (!sheet || sheet.getLastRow() <= 1) return { ok: true, entries: [] };

    const today = todayStr(); // "2026-08-28"
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(String);
    const idxCallSeq   = getHeaderIndex(headers, "CallSeq");
    const idxMulaiTs   = getHeaderIndex(headers, "MulaiLayanTimestamp");
    const idxSelesaiTs = getHeaderIndex(headers, "SelesaiTimestamp");
    const idxEmail     = getHeaderIndex(headers, "Email");
    const idxNoHP      = getHeaderIndex(headers, "NoHP");
    const entries = [];
    // Baca sekali saja seluruh log kolaborasi (multi-petugas) agar tidak N+1 request ke sheet
    const serviceLogMap = getServiceLogMapAll();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Skip baris kosong
      if (!row[0] || String(row[0]).trim() === "") continue;

      // ⚠️ FIX: Konversi kolom Tanggal (index 16) dengan aman
      // Google Sheets sering otomatis mengubah string "2026-08-28" menjadi Date object
      const rowDate = safeDateStr(row[16]);

      // Hanya tampilkan antrian hari ini
      if (rowDate !== today) continue;

      const entryId = String(row[0] || "");
      const collaborators = (serviceLogMap[entryId] || []).map(function(c) { return c.petugasNama; });

      entries.push({
        id:              entryId,
        ticket:          String(row[1]  || ""),
        prefix:          String(row[2]  || "P"),
        number:          Number(row[3]  || 0),
        visitorName:     String(row[4]  || ""),
        purpose:         String(row[5]  || ""),
        visitDetail:     String(row[6]  || ""),
        addressInstansi: String(row[7]  || ""),
        visitorCount:    Number(row[8]  || 1),
        status:          String(row[9]  || "WAITING"),
        petugas:         String(row[10] || ""),
        serviceType:     String(row[11] || ""),
        serviceDetail:   String(row[12] || ""),
        rating:          row[13] ? Number(row[13]) : null,
        ratingComment:   String(row[14] || ""),
        timestamp:       String(row[15] || ""),
        callSeq:         idxCallSeq   >= 0 ? Number(row[idxCallSeq] || 0) : 0,
        mulaiLayanTs:    idxMulaiTs   >= 0 ? String(row[idxMulaiTs] || "") : "",
        selesaiTs:       idxSelesaiTs >= 0 ? String(row[idxSelesaiTs] || "") : "",
        email:           idxEmail     >= 0 ? String(row[idxEmail] || "") : "",
        noHp:            idxNoHP      >= 0 ? String(row[idxNoHP] || "") : "",
        collaborators:   collaborators,
      });
    }

    // Urutkan: IN_SERVICE > WAITING > DONE > CANCELLED
    const order = { "IN_SERVICE": 0, "WAITING": 1, "DONE": 2, "CANCELLED": 3 };
    entries.sort((a, b) => {
      const diff = (order[a.status] ?? 9) - (order[b.status] ?? 9);
      return diff !== 0 ? diff : a.number - b.number;
    });

    return { ok: true, entries: entries };
  } catch (e) {
    return { ok: false, error: e.message, entries: [] };
  }
}

// ─── UPDATE STATUS ANTRIAN ───────────────────────────────
function updateQueueStatus(entryId, newStatus, petugasName, serviceType, serviceDetail) {
  try {
    const sheet = ensureQueueExtraColumns();
    if (!sheet) return { ok: false, error: "Sheet tidak ditemukan." };
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(String);
    const idxCallSeq   = getHeaderIndex(headers, "CallSeq");
    const idxMulaiTs   = getHeaderIndex(headers, "MulaiLayanTimestamp");
    const idxSelesaiTs = getHeaderIndex(headers, "SelesaiTimestamp");

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(entryId)) {
        sheet.getRange(i + 1, 10).setValue(newStatus);
        if (petugasName)   sheet.getRange(i + 1, 11).setValue(petugasName);
        if (serviceType)   sheet.getRange(i + 1, 12).setValue(serviceType);
        if (serviceDetail) sheet.getRange(i + 1, 13).setValue(serviceDetail);

        // Catat waktu mulai dilayani (sekali saja, & set CallSeq=1 agar kiosk memicu pop-up)
        if (newStatus === "IN_SERVICE") {
          if (idxMulaiTs >= 0 && !data[i][idxMulaiTs]) {
            sheet.getRange(i + 1, idxMulaiTs + 1).setValue(nowStr());
          }
          if (idxCallSeq >= 0) {
            const cur = parseInt(data[i][idxCallSeq]) || 0;
            sheet.getRange(i + 1, idxCallSeq + 1).setValue(cur + 1);
          }
        }
        // Catat waktu selesai
        if (newStatus === "DONE" && idxSelesaiTs >= 0) {
          sheet.getRange(i + 1, idxSelesaiTs + 1).setValue(nowStr());
        }
        return { ok: true };
      }
    }
    return { ok: false, error: "Entry tidak ditemukan." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── PANGGIL ULANG (untuk pop-up & suara di Kiosk) ──────
// Menaikkan CallSeq tanpa mengubah status; dipakai tombol "Panggil Ulang" di dashboard.
function recallQueue(entryId) {
  try {
    const sheet = ensureQueueExtraColumns();
    if (!sheet) return { ok: false, error: "Sheet tidak ditemukan." };
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(String);
    const idxCallSeq = getHeaderIndex(headers, "CallSeq");
    if (idxCallSeq < 0) return { ok: false, error: "Kolom CallSeq tidak ditemukan." };

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(entryId)) {
        if (String(data[i][9]) !== "IN_SERVICE") {
          return { ok: false, error: "Hanya antrian yang sedang dilayani yang bisa dipanggil ulang." };
        }
        const cur = parseInt(data[i][idxCallSeq]) || 0;
        sheet.getRange(i + 1, idxCallSeq + 1).setValue(cur + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: "Entry tidak ditemukan." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── SIMPAN RATING ───────────────────────────────────────
function submitRating(entryId, rating, comment) {
  try {
    const sheet = getSheet(SHEET_QUEUE);
    if (!sheet) return { ok: false, error: "Sheet tidak ditemukan." };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(entryId)) {
        sheet.getRange(i + 1, 14).setValue(rating);
        sheet.getRange(i + 1, 15).setValue(comment || "");
        return { ok: true };
      }
    }
    return { ok: false, error: "Entry tidak ditemukan." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── SIMPAN RATING + CATAT KE PetugasRating ─────────────
// petugasNamaOverride (opsional): kalau diisi, rating dicatat atas nama petugas ini —
// tidak harus sama dengan kolom Petugas yang tercatat di tiket. Berguna untuk:
//  - Kunjungan Rapat/Lainnya yang "Petugas"-nya adalah Resepsionis/Admin sendiri, tapi
//    rating ingin diarahkan ke Petugas PST yang sebenarnya membantu.
//  - Tiket dengan lebih dari satu petugas (kolaborasi/pendamping) — pengunjung atau staf
//    dapat memilih siapa di antara mereka yang dinilai.
//  - Resepsionis menunjuk petugas tertentu untuk menerima rating saat mengisikan penilaian
//    atas nama pengunjung.
// Kalau dikosongkan, perilaku lama tetap berlaku: pakai kolom Petugas dari tiket.
function submitRatingWithPetugas(entryId, rating, comment, petugasNamaOverride) {
  try {
    const qSheet = getSheet(SHEET_QUEUE);
    if (!qSheet) return { ok: false, error: "Sheet tidak ditemukan." };
    const data = qSheet.getDataRange().getValues();
    let petugasNama = "";
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(entryId)) {
        qSheet.getRange(i + 1, 14).setValue(rating);
        qSheet.getRange(i + 1, 15).setValue(comment || "");
        petugasNama = String(data[i][10] || "");
        found = true;
        break;
      }
    }
    if (!found) return { ok: false, error: "Entry tidak ditemukan." };

    const petugasUntukRating = String(petugasNamaOverride || "").trim() || petugasNama;
    if (petugasUntukRating) {
      const prSheet = getOrCreateSheet("PetugasRating",
        ["ID","EntryID","PetugasNama","Score","Comment","Tanggal"]);
      prSheet.appendRow(["PR" + Date.now(), entryId, petugasUntukRating, rating, comment || "", todayStr()]);
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── FOTO PROFIL (opsional, berisi URL gambar) ───────────
// Menambahkan kolom FotoUrl ke sheet Users kalau belum ada (aman dijalankan berkali-kali
// & di spreadsheet lama). Nilainya berupa URL gambar langsung (mis. tautan gambar publik
// dari Google Drive/Photos) — sistem ini tidak menangani unggah berkas biner, jadi pegawai/
// Admin cukup menempelkan link gambar yang sudah ada.
function ensureUsersExtraColumns_() {
  const sheet = getOrCreateSheet(SHEET_USERS, ["ID","Nama","Username","Password","Role","IsActive"]);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0].map(String);
  const need = ["FotoUrl"];
  const toAdd = need.filter(function(h) { return headers.indexOf(h) === -1; });
  if (toAdd.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, toAdd.length).setValues([toAdd]);
  }
  return sheet;
}

// Peta cepat PegawaiID -> FotoUrl, dipakai untuk memperkaya data petugas/resepsionis yang
// ditampilkan di Kiosk & Dashboard tanpa perlu request terpisah per orang.
function getFotoUrlMap_() {
  const map = {};
  try {
    const sheet = ensureUsersExtraColumns_();
    if (!sheet || sheet.getLastRow() <= 1) return map;
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxFoto = getHeaderIndex(headers, "FotoUrl");
    if (idxFoto < 0) return map;
    for (let i = 1; i < rows.length; i++) {
      const id = String(rows[i][0] || "");
      if (id) map[id] = String(rows[i][idxFoto] || "");
    }
  } catch (e) { /* abaikan, kembalikan map kosong */ }
  return map;
}

// ─── AUTENTIKASI PETUGAS (LOGIN VIA USERNAME + PASSWORD) ─
// Password kini benar-benar divalidasi (dibandingkan dengan kolom Password di sheet Users),
// supaya fitur "Edit Profil / Ganti Password" berarti. Ada fallback darurat kalau sheet Users
// gagal dibaca, supaya admin/superadmin tidak pernah terkunci total.
function loginPetugas(username, password) {
  try {
    const rawUser = String(username || "").trim();
    const rawPass = String(password || "").trim();
    const inputUser = rawUser.toLowerCase();

    if (!inputUser) return { ok: false, error: "Username tidak boleh kosong." };
    if (!rawPass)   return { ok: false, error: "Password tidak boleh kosong." };

    // 1. Cek di Sheet Users (sumber utama, termasuk superadmin/admin/resepsionis)
    try {
      const sheet = ensureUsersExtraColumns_();
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0].map(String);
        const idxFoto = getHeaderIndex(headers, "FotoUrl");
        const cleanInput = inputUser.replace(/[^a-z0-9]/g, "");

        for (let i = 1; i < data.length; i++) {
          const rowId       = String(data[i][0] || "");
          const rowNama     = String(data[i][1] || "").trim();
          const rowUsername = String(data[i][2] || "").trim();
          const rowPassword = String(data[i][3] || "");
          const rowRole     = String(data[i][4] || "PETUGAS");
          const rowActive   = data[i][5];
          const rowFoto     = idxFoto >= 0 ? String(data[i][idxFoto] || "") : "";

          if (rowActive === false) continue;

          const cleanNama = rowNama.toLowerCase().replace(/[^a-z0-9]/g, "");
          const cleanUser = rowUsername.toLowerCase().replace(/[^a-z0-9]/g, "");
          const isMatchUser = rowUsername.toLowerCase() === inputUser || rowNama.toLowerCase() === inputUser ||
                               cleanUser === cleanInput || cleanNama === cleanInput;

          if (isMatchUser) {
            if (rowPassword && rowPassword !== rawPass) {
              return { ok: false, error: "Username atau password salah." };
            }
            return {
              ok: true,
              user: {
                id:       rowId || "USR_PETUGAS",
                nama:     rowNama || rawUser,
                username: rowUsername || rawUser,
                role:     rowRole,
                fotoUrl:  rowFoto
              }
            };
          }
        }
      }
    } catch (errSheet) {
      // Sheet Users gagal dibaca — lanjut ke jaring pengaman darurat di bawah.
    }

    // 2. Jaring Pengaman Darurat — hanya aktif kalau sheet Users benar2 tidak bisa dibaca sama sekali
    //    (misalnya belum pernah setupAllPegawaiData()). Menjaga admin tidak pernah terkunci total.
    try {
      const sheetCheck = getSheet(SHEET_USERS);
      if (!sheetCheck || sheetCheck.getLastRow() <= 1) {
        if (inputUser.includes("superadmin") || inputUser.includes("super")) {
          return { ok: true, user: { id: "USR000", nama: "Super Admin BPS", username: "superadmin", role: "SUPERADMIN" } };
        }
        if (inputUser.includes("admin")) {
          return { ok: true, user: { id: "USR000B", nama: "Admin BPS Lampung", username: "admin", role: "ADMIN" } };
        }
        if (inputUser.includes("resep")) {
          return { ok: true, user: { id: "USR000C", nama: "Resepsionis PST", username: "resepsionis", role: "RESEPSIONIS" } };
        }
      }
    } catch (e2) { /* abaikan */ }

    // 3. Username tidak ditemukan di manapun
    return { ok: false, error: "Username tidak ditemukan. Hubungi admin untuk didaftarkan." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── INFO KANTOR ─────────────────────────────────────────
function getKantorInfo() {
  return {
    nama: getConfig("KANTOR_NAMA") || "BPS",
    kode: getConfig("KANTOR_KODE") || "BPS-LAMPUNG",
  };
}

// ════════════════════════════════════════════════════════
// FITUR 3 & 4 — JADWAL BULANAN (ADMIN PILIH & ROTASI PETUGAS/RESEPSIONIS)
// ════════════════════════════════════════════════════════

// Daftar seluruh pegawai (untuk dropdown pemilihan admin & tab Master Pegawai)
function getEmployeeList() {
  try {
    const sheet = ensureUsersExtraColumns_();
    if (!sheet || sheet.getLastRow() <= 1) return { ok: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxFoto = getHeaderIndex(headers, "FotoUrl");
    const data = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      data.push({
        id:       String(rows[i][0] || ""),
        nama:     String(rows[i][1] || ""),
        username: String(rows[i][2] || ""),
        role:     String(rows[i][4] || "PETUGAS"),
        isActive: rows[i][5] !== false,
        fotoUrl:  idxFoto >= 0 ? String(rows[i][idxFoto] || "") : "",
      });
    }
    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, error: e.message, data: [] };
  }
}

// ─── MASTER PEGAWAI: TAMBAH / EDIT (Admin) ───────────────
function addEmployee(nama, username, password, role, fotoUrl) {
  try {
    nama = String(nama || "").trim();
    username = String(username || "").trim();
    if (!nama || !username) return { ok: false, error: "Nama dan username wajib diisi." };

    const sheet = ensureUsersExtraColumns_();
    const rows = sheet.getDataRange().getValues();
    const uname = username.toLowerCase();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2] || "").trim().toLowerCase() === uname) {
        return { ok: false, error: "Username sudah digunakan pegawai lain." };
      }
    }
    const id = "USR_" + Utilities.getUuid().substring(0, 8).toUpperCase();
    const finalRole = role || "PEGAWAI";
    const finalFoto = String(fotoUrl || "").trim();
    sheet.appendRow([id, nama, username, password || "bps2025", finalRole, true, finalFoto]);
    return { ok: true, employee: { id: id, nama: nama, username: username, role: finalRole, isActive: true, fotoUrl: finalFoto } };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function updateEmployeeById(id, nama, username, password, role, isActive, fotoUrl) {
  try {
    const sheet = ensureUsersExtraColumns_();
    if (!sheet) return { ok: false, error: "Sheet Users tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxFoto = getHeaderIndex(headers, "FotoUrl");
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) {
        if (nama)     sheet.getRange(i + 1, 2).setValue(nama);
        if (username) sheet.getRange(i + 1, 3).setValue(username);
        if (password) sheet.getRange(i + 1, 4).setValue(password);
        if (role)     sheet.getRange(i + 1, 5).setValue(role);
        if (typeof isActive === "boolean") sheet.getRange(i + 1, 6).setValue(isActive);
        if (typeof fotoUrl === "string" && idxFoto >= 0) sheet.getRange(i + 1, idxFoto + 1).setValue(fotoUrl.trim());
        return {
          ok: true,
          employee: {
            id: id,
            nama: nama || String(rows[i][1] || ""),
            username: username || String(rows[i][2] || ""),
            role: role || String(rows[i][4] || "PETUGAS"),
            isActive: typeof isActive === "boolean" ? isActive : (rows[i][5] !== false),
            fotoUrl: typeof fotoUrl === "string" ? fotoUrl.trim() : (idxFoto >= 0 ? String(rows[i][idxFoto] || "") : "")
          }
        };
      }
    }
    return { ok: false, error: "Pegawai tidak ditemukan." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Ambil roster bulan tertentu (default bulan berjalan). periode format "yyyy-MM".
function getJadwalBulanan(periode) {
  try {
    const p = periode || currentPeriode();
    const sheet = getOrCreateSheet(SHEET_JADWAL_BULANAN, ["ID","Periode","PegawaiID","Nama","Role","IsActive"]);
    const rows = sheet.getDataRange().getValues();
    const petugas = [];
    let resepsionis = null;
    for (let i = 1; i < rows.length; i++) {
      if (safePeriodeStr(rows[i][1]) !== p) continue;
      if (rows[i][5] === false) continue;
      const item = { id: String(rows[i][2] || ""), nama: String(rows[i][3] || "") };
      if (String(rows[i][4]) === "PETUGAS") petugas.push(item);
      else if (String(rows[i][4]) === "RESEPSIONIS") resepsionis = item;
    }
    return { ok: true, periode: p, petugas: petugas, resepsionis: resepsionis };
  } catch (e) {
    return { ok: false, error: e.message, petugas: [], resepsionis: null };
  }
}

// Simpan roster bulanan. petugasList = [{id,nama}, ...] maks 3. resepsionis = {id,nama} atau null.
function saveJadwalBulanan(periode, petugasList, resepsionis) {
  try {
    const p = periode || currentPeriode();
    petugasList = petugasList || [];
    if (petugasList.length > MAX_PETUGAS_PER_BULAN) {
      return { ok: false, error: "Maksimal " + MAX_PETUGAS_PER_BULAN + " petugas per bulan." };
    }
    const sheet = getOrCreateSheet(SHEET_JADWAL_BULANAN, ["ID","Periode","PegawaiID","Nama","Role","IsActive"]);
    const rows = sheet.getDataRange().getValues();

    // Hapus baris lama untuk periode ini (replace total)
    for (let i = rows.length - 1; i >= 1; i--) {
      if (safePeriodeStr(rows[i][1]) === p) sheet.deleteRow(i + 1);
    }

    // Tulis baris baru — kolom Periode dipaksa format Teks ("@") agar TIDAK
    // otomatis dikonversi Google Sheets menjadi objek Tanggal (yang bikin
    // pencocokan periode di getJadwalBulanan selalu gagal).
    function appendRoleRow(pegawaiId, nama, role) {
      sheet.appendRow(["JB" + Utilities.getUuid().substring(0, 8), p, pegawaiId, nama, role, true]);
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 2).setNumberFormat("@").setValue(p);
    }
    petugasList.forEach(function(pt) { appendRoleRow(pt.id, pt.nama, "PETUGAS"); });
    if (resepsionis && resepsionis.id) appendRoleRow(resepsionis.id, resepsionis.nama, "RESEPSIONIS");

    return { ok: true, message: "Jadwal bulanan periode " + p + " berhasil disimpan." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════
// FITUR 5 & 6 — JADWAL HARIAN (RESEPSIONIS PILIH PETUGAS ON-DUTY HARI INI)
// ════════════════════════════════════════════════════════

function getJadwalHarian(tanggal) {
  try {
    const t = tanggal || todayStr();
    const sheet = ensureJadwalHarianExtraColumns_();
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxDuty = getHeaderIndex(headers, "IsDutyModeOn");
    const data = [];
    for (let i = 1; i < rows.length; i++) {
      if (safeDateStr(rows[i][1]) !== t) continue;
      data.push({
        id: String(rows[i][2] || ""),
        nama: String(rows[i][3] || ""),
        // Kolom baru: default TRUE untuk baris lama yang belum punya nilai (kompatibel mundur)
        isDutyModeOn: idxDuty >= 0 ? (rows[i][idxDuty] !== false) : true
      });
    }
    return { ok: true, tanggal: t, data: data };
  } catch (e) {
    return { ok: false, error: e.message, data: [] };
  }
}

// Ambil jadwal harian SATU BULAN PENUH sekaligus (untuk tampilan kalender) — hemat request.
// Return: { ok, periode, data: { "2026-08-01": [{id,nama}, ...], "2026-08-02": [...], ... } }
function getJadwalHarianBulan(periode) {
  try {
    const p = periode || currentPeriode();
    const sheet = getOrCreateSheet(SHEET_JADWAL_HARIAN, ["ID","Tanggal","PegawaiID","Nama","DipilihOleh","Timestamp"]);
    const rows = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < rows.length; i++) {
      const tgl = safeDateStr(rows[i][1]);
      if (tgl.substring(0, 7) !== p) continue;
      if (!map[tgl]) map[tgl] = [];
      map[tgl].push({ id: String(rows[i][2] || ""), nama: String(rows[i][3] || "") });
    }
    return { ok: true, periode: p, data: map };
  } catch (e) {
    return { ok: false, error: e.message, data: {} };
  }
}

// petugasIds = array id pegawai yang on-duty hari itu (bisa dari roster bulanan atau tambahan ad-hoc resepsionis)
function saveJadwalHarian(tanggal, petugasList, dipilihOleh) {
  try {
    const t = tanggal || todayStr();
    petugasList = petugasList || [];
    const sheet = ensureJadwalHarianExtraColumns_();
    const rows = sheet.getDataRange().getValues();

    for (let i = rows.length - 1; i >= 1; i--) {
      if (safeDateStr(rows[i][1]) === t) sheet.deleteRow(i + 1);
    }
    petugasList.forEach(function(pt) {
      // IsDutyModeOn default TRUE saat baru dijadwalkan; petugas bisa toggle sendiri belakangan.
      sheet.appendRow(["JH" + Utilities.getUuid().substring(0, 8), t, pt.id, pt.nama, dipilihOleh || "", nowStr(), true]);
    });
    return { ok: true, message: "Jadwal harian " + t + " berhasil disimpan." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Simpan Jadwal Harian untuk BANYAK tanggal sekaligus dalam satu eksekusi (dipakai fitur
// Impor Jadwal Harian dari Excel), supaya Admin/Resepsionis tidak perlu mengulang proses
// simpan satu tanggal per satu tanggal secara manual. Membaca & menulis sheet hanya sekali,
// jadi jauh lebih cepat dan aman dari resiko tumpang-tindih dibanding memanggil
// saveJadwalHarian() berkali-kali secara paralel dari klien.
// dataByDate: { "yyyy-MM-dd": [{id, nama}, ...], ... }
function saveJadwalHarianBulk(dataByDate, dipilihOleh) {
  try {
    dataByDate = dataByDate || {};
    const targetDates = Object.keys(dataByDate);
    if (targetDates.length === 0) return { ok: false, error: "Tidak ada data untuk disimpan." };

    const sheet = ensureJadwalHarianExtraColumns_();
    const rows = sheet.getDataRange().getValues();
    const targetSet = {};
    targetDates.forEach(function(d) { targetSet[d] = true; });

    // Hapus semua baris lama untuk tanggal-tanggal target (dari bawah ke atas agar index tidak bergeser)
    for (let i = rows.length - 1; i >= 1; i--) {
      const tgl = safeDateStr(rows[i][1]);
      if (targetSet[tgl]) sheet.deleteRow(i + 1);
    }

    // Tambahkan baris baru untuk seluruh tanggal sekaligus
    let totalRows = 0;
    targetDates.forEach(function(tgl) {
      (dataByDate[tgl] || []).forEach(function(pt) {
        sheet.appendRow(["JH" + Utilities.getUuid().substring(0, 8), tgl, pt.id || "", pt.nama || "", dipilihOleh || "", nowStr(), true]);
        totalRows++;
      });
    });

    return {
      ok: true,
      message: "Berhasil menyimpan jadwal harian untuk " + targetDates.length + " tanggal (" + totalRows + " baris penugasan petugas).",
      dates: targetDates.length, rows: totalRows
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Petugas yang tersedia hari ini = irisan roster bulanan (Role=PETUGAS) dengan pilihan harian resepsionis.
// Dipakai resepsionis/admin saat menugaskan "Mulai Layani", dan petugas sendiri untuk cek apakah dirinya bertugas.
function getPetugasTersediaHariIni(tanggal) {
  try {
    const t = tanggal || todayStr();
    const periode = t.substring(0, 7); // yyyy-MM dari tanggal
    const bulanan = getJadwalBulanan(periode);
    const harian = getJadwalHarian(t);
    if (!bulanan.ok) return { ok: false, error: bulanan.error, data: [], resepsionis: null };
    // Hanya petugas yang dijadwalkan HARI INI dan tidak sedang menonaktifkan Mode Bertugas-nya
    const harianAktif = (harian.data || []).filter(function(x) { return x.isDutyModeOn !== false; });
    const idsHarian = harianAktif.map(function(x) { return x.id; });
    // Hanya petugas yang ADA di roster bulanan DAN dipilih resepsionis hari ini DAN aktif
    const tersedia = (bulanan.petugas || []).filter(function(p) { return idsHarian.indexOf(p.id) !== -1; });
    // Tambahan ad-hoc: kalau resepsionis pilih ID yang tidak ada di roster bulanan (penugasan darurat), tetap ikutkan
    harianAktif.forEach(function(h) {
      if (!tersedia.some(function(x) { return x.id === h.id; })) tersedia.push(h);
    });

    // Perkaya dengan Foto Profil (kalau ada) — dipakai tampilan Kiosk & Dashboard supaya
    // menampilkan avatar, bukan cuma nama teks.
    const fotoMap = getFotoUrlMap_();
    tersedia.forEach(function(p) { p.fotoUrl = fotoMap[p.id] || ""; });

    // Resepsionis yang bertugas bulan ini (dari Jadwal Bulanan) — resepsionis tidak melalui
    // toggle Mode Bertugas/Jadwal Harian seperti Petugas, jadi statusnya "bertugas" selama
    // masih tercatat sebagai Resepsionis periode berjalan.
    let resepsionisData = null;
    if (bulanan.resepsionis) {
      resepsionisData = {
        id: bulanan.resepsionis.id,
        nama: bulanan.resepsionis.nama,
        fotoUrl: fotoMap[bulanan.resepsionis.id] || ""
      };
    }

    return { ok: true, tanggal: t, data: tersedia, resepsionis: resepsionisData };
  } catch (e) {
    return { ok: false, error: e.message, data: [], resepsionis: null };
  }
}

// ════════════════════════════════════════════════════════
// FITUR 2 & 7 — REKAP HARIAN, DASHBOARD INSIGHT, REKAP RATING
// ════════════════════════════════════════════════════════

function getDailyRecap(tanggal) {
  try {
    const t = tanggal || todayStr();
    const sheet = ensureQueueExtraColumns();
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, tanggal: t, total: 0, byStatus: {}, byPrefix: {}, avgRating: 0, ratingCount: 0 };
    }
    const rows = sheet.getDataRange().getValues();
    let total = 0;
    const byStatus = { WAITING: 0, IN_SERVICE: 0, DONE: 0, CANCELLED: 0 };
    const byPrefix = {};
    let ratingSum = 0, ratingCount = 0;

    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (safeDateStr(rows[i][16]) !== t) continue;
      total++;
      const status = String(rows[i][9] || "WAITING");
      byStatus[status] = (byStatus[status] || 0) + 1;
      const prefix = String(rows[i][2] || "U");
      byPrefix[prefix] = (byPrefix[prefix] || 0) + 1;
      const rating = rows[i][13];
      if (rating) { ratingSum += Number(rating); ratingCount++; }
    }
    return {
      ok: true, tanggal: t, total: total, byStatus: byStatus, byPrefix: byPrefix,
      avgRating: ratingCount > 0 ? Math.round((ratingSum / ratingCount) * 10) / 10 : 0,
      ratingCount: ratingCount
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Insight ringkas: tren N hari terakhir, top petugas, rata2 waktu tunggu & waktu layanan (menit)
function getDashboardInsight(days) {
  try {
    const n = days || 7;
    const sheet = ensureQueueExtraColumns();
    if (!sheet || sheet.getLastRow() <= 1) {
      return { ok: true, trend: [], topPetugas: [], avgWaitMin: 0, avgServiceMin: 0 };
    }
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxMulai   = getHeaderIndex(headers, "MulaiLayanTimestamp");
    const idxSelesai = getHeaderIndex(headers, "SelesaiTimestamp");

    // Siapkan tanggal-tanggal N hari terakhir
    const dateList = [];
    for (let d = n - 1; d >= 0; d--) {
      const dt = new Date();
      dt.setDate(dt.getDate() - d);
      dateList.push(Utilities.formatDate(dt, "Asia/Jakarta", "yyyy-MM-dd"));
    }
    const trendMap = {};
    dateList.forEach(function(d) { trendMap[d] = 0; });

    const petugasMap = {}; // nama -> {count, ratingSum, ratingCount}
    let waitSumMs = 0, waitCount = 0;
    let serviceSumMs = 0, serviceCount = 0;

    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      const tgl = safeDateStr(rows[i][16]);
      if (trendMap.hasOwnProperty(tgl)) trendMap[tgl]++;

      const petugasNama = String(rows[i][10] || "");
      if (petugasNama && tgl === todayStr()) {
        if (!petugasMap[petugasNama]) petugasMap[petugasNama] = { nama: petugasNama, count: 0, ratingSum: 0, ratingCount: 0 };
        petugasMap[petugasNama].count++;
        const rating = rows[i][13];
        if (rating) { petugasMap[petugasNama].ratingSum += Number(rating); petugasMap[petugasNama].ratingCount++; }
      }

      // Waktu tunggu = MulaiLayan - Timestamp registrasi (hanya hari ini, biar relevan)
      if (tgl === todayStr() && idxMulai >= 0 && rows[i][idxMulai] && rows[i][15]) {
        const regTs = new Date(String(rows[i][15]).replace(" ", "T"));
        const mulaiTs = new Date(String(rows[i][idxMulai]).replace(" ", "T"));
        const diff = mulaiTs - regTs;
        if (!isNaN(diff) && diff >= 0) { waitSumMs += diff; waitCount++; }
      }
      // Waktu layanan = Selesai - MulaiLayan (hanya hari ini)
      if (tgl === todayStr() && idxSelesai >= 0 && idxMulai >= 0 && rows[i][idxSelesai] && rows[i][idxMulai]) {
        const mulaiTs = new Date(String(rows[i][idxMulai]).replace(" ", "T"));
        const selesaiTs = new Date(String(rows[i][idxSelesai]).replace(" ", "T"));
        const diff = selesaiTs - mulaiTs;
        if (!isNaN(diff) && diff >= 0) { serviceSumMs += diff; serviceCount++; }
      }
    }

    const trend = dateList.map(function(d) { return { tanggal: d, total: trendMap[d] }; });
    const topPetugas = Object.values(petugasMap).map(function(p) {
      return {
        nama: p.nama, count: p.count,
        avgRating: p.ratingCount > 0 ? Math.round((p.ratingSum / p.ratingCount) * 10) / 10 : null
      };
    }).sort(function(a, b) { return b.count - a.count; }).slice(0, 5);

    return {
      ok: true,
      trend: trend,
      topPetugas: topPetugas,
      avgWaitMin: waitCount > 0 ? Math.round(waitSumMs / waitCount / 60000) : 0,
      avgServiceMin: serviceCount > 0 ? Math.round(serviceSumMs / serviceCount / 60000) : 0
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Rekap rating hari ini per petugas + daftar komentar
function getTodayRatingRecap() {
  try {
    const sheet = getSheet(SHEET_QUEUE);
    if (!sheet || sheet.getLastRow() <= 1) return { ok: true, perPetugas: [], list: [] };
    const rows = sheet.getDataRange().getValues();
    const today = todayStr();
    const map = {};
    const list = [];
    for (let i = 1; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      if (safeDateStr(rows[i][16]) !== today) continue;
      const rating = rows[i][13];
      if (!rating) continue;
      const petugas = String(rows[i][10] || "(tanpa nama)");
      if (!map[petugas]) map[petugas] = { nama: petugas, sum: 0, count: 0 };
      map[petugas].sum += Number(rating);
      map[petugas].count++;
      list.push({
        ticket: String(rows[i][1] || ""),
        visitorName: String(rows[i][4] || ""),
        petugas: petugas,
        rating: Number(rating),
        comment: String(rows[i][14] || "")
      });
    }
    const perPetugas = Object.values(map).map(function(m) {
      return { nama: m.nama, count: m.count, avg: Math.round((m.sum / m.count) * 10) / 10 };
    }).sort(function(a, b) { return b.avg - a.avg; });
    return { ok: true, perPetugas: perPetugas, list: list };
  } catch (e) {
    return { ok: false, error: e.message, perPetugas: [], list: [] };
  }
}

// ════════════════════════════════════════════════════════
// FITUR 8 — EDIT / UPDATE PROFIL (ADMIN / RESEPSIONIS / PETUGAS)
// ════════════════════════════════════════════════════════

function updateProfile(currentUsername, newNama, newUsername, newPassword, newFotoUrl) {
  try {
    const sheet = ensureUsersExtraColumns_();
    if (!sheet || sheet.getLastRow() <= 1) return { ok: false, error: "Data pegawai belum di-setup." };
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxFoto = getHeaderIndex(headers, "FotoUrl");
    const cur = String(currentUsername || "").trim().toLowerCase();

    for (let i = 1; i < rows.length; i++) {
      const rowUsername = String(rows[i][2] || "").trim().toLowerCase();
      const rowNama     = String(rows[i][1] || "").trim().toLowerCase();
      if (rowUsername === cur || rowNama === cur) {
        if (newNama)     sheet.getRange(i + 1, 2).setValue(newNama);
        if (newUsername) sheet.getRange(i + 1, 3).setValue(newUsername);
        if (newPassword) sheet.getRange(i + 1, 4).setValue(newPassword);
        if (typeof newFotoUrl === "string" && idxFoto >= 0) sheet.getRange(i + 1, idxFoto + 1).setValue(newFotoUrl.trim());
        return {
          ok: true,
          user: {
            id:       String(rows[i][0] || ""),
            nama:     newNama || String(rows[i][1] || ""),
            username: newUsername || String(rows[i][2] || ""),
            role:     String(rows[i][4] || "PETUGAS"),
            fotoUrl:  typeof newFotoUrl === "string" ? newFotoUrl.trim() : (idxFoto >= 0 ? String(rows[i][idxFoto] || "") : "")
          }
        };
      }
    }
    return { ok: false, error: "Akun tidak ditemukan di data pegawai." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── AMBIL SUMMARY RATING PETUGAS ────────────────────────
function getPetugasRatingSummary() {
  try {
    const sheet = getSheet("PetugasRating");
    if (!sheet || sheet.getLastRow() <= 1) return { ok: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < rows.length; i++) {
      const nama  = String(rows[i][2] || "");
      const score = parseInt(rows[i][3]) || 0;
      if (!nama) continue;
      if (!map[nama]) map[nama] = { nama: nama, total: 0, count: 0 };
      map[nama].total += score;
      map[nama].count += 1;
    }
    const data = Object.values(map).map(m => ({
      nama: m.nama,
      count: m.count,
      avg: m.count > 0 ? (m.total / m.count).toFixed(1) : "0"
    }));
    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, error: e.message, data: [] };
  }
}

// ════════════════════════════════════════════════════════
// FITUR 9 — PELACAKAN TIKET LIVE UNTUK PENGUNJUNG (?page=tiket)
// ════════════════════════════════════════════════════════

// Ambil detail 1 tiket + posisi antrian & estimasi menit tunggu, untuk halaman Tiket.html
// (dipoll berkala oleh pengunjung dari perangkatnya sendiri).
function getTicketLive(entryId) {
  try {
    const sheet = ensureQueueExtraColumns();
    if (!sheet || sheet.getLastRow() <= 1) return { ok: false, error: "Tiket tidak ditemukan." };
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxEmail = getHeaderIndex(headers, "Email");
    const idxNoHP  = getHeaderIndex(headers, "NoHP");

    let found = null, foundIdx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(entryId)) { found = rows[i]; foundIdx = i; break; }
    }
    if (!found) return { ok: false, error: "Tiket tidak ditemukan. Pastikan tautan yang dibuka masih benar." };

    const entry = {
      id:              String(found[0]  || ""),
      ticket:          String(found[1]  || ""),
      prefix:          String(found[2]  || ""),
      number:          Number(found[3]  || 0),
      visitorName:     String(found[4]  || ""),
      purpose:         String(found[5]  || ""),
      visitDetail:     String(found[6]  || ""),
      addressInstansi: String(found[7]  || ""),
      visitorCount:    Number(found[8]  || 1),
      status:          String(found[9]  || "WAITING"),
      petugas:         String(found[10] || ""),
      serviceType:     String(found[11] || ""),
      serviceDetail:   String(found[12] || ""),
      rating:          found[13] ? Number(found[13]) : null,
      ratingComment:   String(found[14] || ""),
      tanggal:         safeDateStr(found[16]),
      email:           idxEmail >= 0 ? String(found[idxEmail] || "") : "",
      noHp:            idxNoHP  >= 0 ? String(found[idxNoHP]  || "") : "",
      collaborators:   (getServiceLogMapAll()[String(found[0] || "")] || []).map(function(c) { return c.petugasNama; }),
    };

    // Hitung berapa antrian yang masih di depan (prefix sama, sama-sama WAITING, nomor lebih kecil)
    let antrianDiDepan = 0;
    if (entry.status === "WAITING") {
      for (let i = 1; i < rows.length; i++) {
        if (i === foundIdx) continue;
        if (String(rows[i][2]) !== entry.prefix) continue;
        if (String(rows[i][9]) !== "WAITING") continue;
        if (safeDateStr(rows[i][16]) !== entry.tanggal) continue;
        if (Number(rows[i][3] || 0) < entry.number) antrianDiDepan++;
      }
    }
    // Estimasi kasar: 10 menit per antrian yang masih di depan (kategori sama)
    const estimasiMenit = antrianDiDepan * 10;

    return { ok: true, entry: entry, antrianDiDepan: antrianDiDepan, estimasiMenit: estimasiMenit };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════
// FITUR 10 — PING / READY (RESEPSIONIS ⇄ PETUGAS)
// ════════════════════════════════════════════════════════
// Catatan: Google Apps Script tidak mendukung push/websocket, jadi notifikasi
// ini bekerja via polling (lihat Dashboard.html, setiap ±4 detik) — bukan real-time
// murni, tapi cukup untuk kebutuhan kantor kecil.

const SHEET_PING = "PingNotification";

function ensurePingSheet_() {
  return getOrCreateSheet(SHEET_PING,
    ["ID","EntryID","EntryTicket","VisitorName","FromNama","ToPetugasID","ToPetugasNama","Status","CreatedAt","RespondedAt"]);
}

// Dipanggil Resepsionis/Admin: beritahu petugas tertentu bahwa ada tamu menunggu.
function sendPing(entryId, entryTicket, visitorName, fromNama, toPetugasId, toPetugasNama) {
  try {
    if (!toPetugasId && !toPetugasNama) return { ok: false, error: "Pilih petugas yang akan di-ping." };
    const sheet = ensurePingSheet_();
    const id = "PING" + Utilities.getUuid().substring(0, 8).toUpperCase();
    sheet.appendRow([id, entryId, entryTicket || "", visitorName || "", fromNama || "",
      toPetugasId || "", toPetugasNama || "", "PENDING", nowStr(), ""]);
    return { ok: true, id: id };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Dipoll Petugas: ping yang masih PENDING dan ditujukan untuknya.
function getPendingPingsForPetugas(petugasId, petugasNama) {
  try {
    const sheet = ensurePingSheet_();
    if (!sheet || sheet.getLastRow() <= 1) return { ok: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    const data = [];
    const namaLower = String(petugasNama || "").toLowerCase();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]) !== "PENDING") continue;
      const toId = String(rows[i][5] || "");
      const toNamaLower = String(rows[i][6] || "").toLowerCase();
      if (toId === String(petugasId) || (namaLower && toNamaLower === namaLower)) {
        data.push({
          id: String(rows[i][0]), entryId: String(rows[i][1]), ticket: String(rows[i][2]),
          visitorName: String(rows[i][3]), fromNama: String(rows[i][4]), createdAt: String(rows[i][8])
        });
      }
    }
    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, error: e.message, data: [] };
  }
}

// Dipanggil Petugas saat klik "Saya Siap Melayani".
function respondPingReady(pingId, petugasNama) {
  try {
    const sheet = ensurePingSheet_();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(pingId)) {
        sheet.getRange(i + 1, 8).setValue("READY");
        sheet.getRange(i + 1, 10).setValue(nowStr());
        return { ok: true };
      }
    }
    return { ok: false, error: "Ping tidak ditemukan (mungkin sudah kedaluwarsa)." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Dipoll Resepsionis/Admin: ping yang dikirimnya dan sudah dijawab READY oleh petugas.
function getReadyPingsForResepsionis(fromNama) {
  try {
    const sheet = ensurePingSheet_();
    if (!sheet || sheet.getLastRow() <= 1) return { ok: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    const data = [];
    const fromLower = String(fromNama || "").toLowerCase();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][7]) !== "READY") continue;
      if (String(rows[i][4] || "").toLowerCase() !== fromLower) continue;
      data.push({
        id: String(rows[i][0]), ticket: String(rows[i][2]), visitorName: String(rows[i][3]),
        toPetugasNama: String(rows[i][6]), respondedAt: String(rows[i][9])
      });
    }
    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, error: e.message, data: [] };
  }
}

// ════════════════════════════════════════════════════════
// FITUR 11 — MULTI-PETUGAS (KOLABORASI PENDAMPING PER TIKET)
// ════════════════════════════════════════════════════════
// Petugas utama tetap tercatat di kolom Queue.Petugas seperti sebelumnya.
// Petugas tambahan (pendamping) dicatat di sheet terpisah agar 1 tiket bisa
// dilayani lebih dari 1 orang tanpa mengubah struktur kolom Queue yang sudah ada.

const SHEET_SERVICELOG = "QueueServiceLog";

function ensureServiceLogSheet_() {
  return getOrCreateSheet(SHEET_SERVICELOG, ["ID","EntryID","PetugasID","PetugasNama","PeranLog","Timestamp"]);
}

// Baca seluruh log sekali, dikelompokkan per EntryID — dipakai getQueueToday agar hemat request.
function getServiceLogMapAll() {
  const sheet = ensureServiceLogSheet_();
  const map = {};
  if (!sheet || sheet.getLastRow() <= 1) return map;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const eid = String(rows[i][1] || "");
    if (!eid) continue;
    if (!map[eid]) map[eid] = [];
    map[eid].push({ petugasId: String(rows[i][2] || ""), petugasNama: String(rows[i][3] || ""), peran: String(rows[i][4] || "") });
  }
  return map;
}

function getServiceLogForEntry(entryId) {
  try {
    const map = getServiceLogMapAll();
    return { ok: true, data: map[String(entryId)] || [] };
  } catch (e) {
    return { ok: false, error: e.message, data: [] };
  }
}

// Tambahkan petugas pendamping untuk tiket yang sedang/sudah dilayani.
function joinServiceLog(entryId, petugasId, petugasNama) {
  try {
    if (!petugasNama) return { ok: false, error: "Pilih petugas pendamping terlebih dahulu." };
    const sheet = ensureServiceLogSheet_();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === String(entryId) &&
          (String(rows[i][2]) === String(petugasId) || String(rows[i][3]).toLowerCase() === String(petugasNama).toLowerCase())) {
        return { ok: false, error: "Petugas tersebut sudah tercatat sebagai pendamping tiket ini." };
      }
    }
    const id = "SL" + Utilities.getUuid().substring(0, 8).toUpperCase();
    sheet.appendRow([id, entryId, petugasId || "", petugasNama, "PENDAMPING", nowStr()]);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════
// FITUR 12 — MODE BERTUGAS MANDIRI (TOGGLE PETUGAS)
// ════════════════════════════════════════════════════════
// Melengkapi (bukan menggantikan) Jadwal Bulanan+Harian: petugas yang sudah
// dijadwalkan tetap dapat menandai dirinya "non-aktif sementara" (mis. sedang
// istirahat/rapat) tanpa perlu Resepsionis mengubah Jadwal Harian.

function ensureJadwalHarianExtraColumns_() {
  const sheet = getOrCreateSheet(SHEET_JADWAL_HARIAN, ["ID","Tanggal","PegawaiID","Nama","DipilihOleh","Timestamp"]);
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0].map(String);
  const need = ["IsDutyModeOn"];
  const toAdd = need.filter(function(h) { return headers.indexOf(h) === -1; });
  if (toAdd.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, toAdd.length).setValues([toAdd]);
  }
  return sheet;
}

// Petugas klik toggle "Mode Bertugas" di Dashboard → ubah IsDutyModeOn untuk baris jadwal hariannya.
function setDutyModeToggle(tanggal, petugasId, isOn) {
  try {
    const t = tanggal || todayStr();
    const sheet = ensureJadwalHarianExtraColumns_();
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(String);
    const idxTgl  = getHeaderIndex(headers, "Tanggal");
    const idxPeg  = getHeaderIndex(headers, "PegawaiID");
    const idxDuty = getHeaderIndex(headers, "IsDutyModeOn");
    for (let i = 1; i < rows.length; i++) {
      if (safeDateStr(rows[i][idxTgl]) === t && String(rows[i][idxPeg]) === String(petugasId)) {
        sheet.getRange(i + 1, idxDuty + 1).setValue(!!isOn);
        return { ok: true, isOn: !!isOn };
      }
    }
    return { ok: false, error: "Anda tidak terjadwal pada tanggal ini, sehingga status Mode Bertugas tidak dapat diubah." };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// Status toggle petugas untuk hari ini — dipanggil saat login/tab Antrian dimuat.
function getMyDutyModeToday(petugasId) {
  try {
    const t = todayStr();
    const harian = getJadwalHarian(t);
    const mine = (harian.data || []).find(function(x) { return x.id === petugasId; });
    if (!mine) return { ok: true, scheduled: false, isOn: false };
    return { ok: true, scheduled: true, isOn: mine.isDutyModeOn !== false };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ════════════════════════════════════════════════════════
// FITUR 13 — SINKRONISASI MASTER PEGAWAI DARI WAWAI API
// ════════════════════════════════════════════════════════
// Skema respons Wawai API (terverifikasi langsung ke endpoint, bukan asumsi):
//   { "success": true, "total": <n>, "raw_data": [
//       { "id":.., "nama":.., "nip_lama":.., "nip_baru":.., "satker":..,
//         "golongan":.., "nama_jabatan":.., "pendidikan_terakhir":..,
//         "updated_at":.., "nama_tanpa_gelar":.. }, ... ] }
//
// ⚠️ CATATAN KEAMANAN: URL di bawah ini menyertakan API key. Karena Code.gs berjalan
// di server (Apps Script), key ini TIDAK pernah terkirim ke browser pengunjung — tapi
// tetap akan ikut ter-commit kalau file ini diunggah ke repo publik / dibagikan ke orang
// lain. Kalau itu terjadi, jalankan setWawaiApiUrl(url) SEKALI dari editor Apps Script
// untuk memindahkan URL+key ke Script Properties (lebih aman), lalu hapus/kosongkan
// nilai default di bawah sebelum membagikan kodenya.
const WAWAI_API_URL_DEFAULT = "https://wawai.statapps.dev/api/public/pegawai?api_key=ak_prod_9f8d7c6b5a4e3d2c1b0a9f8e7d6c5b4a3f2e1d0c8b9a1c2d3e4f5a6b7";

function getWawaiApiUrl_() {
  const stored = PropertiesService.getScriptProperties().getProperty("WAWAI_API_URL");
  return stored || WAWAI_API_URL_DEFAULT;
}

// Jalankan SEKALI dari editor Apps Script untuk menyimpan URL+API key di Script
// Properties (bukan di source code) — opsional, tapi disarankan sebelum membagikan
// project ini ke orang lain atau meng-commit ke repo publik.
function setWawaiApiUrl(url) {
  PropertiesService.getScriptProperties().setProperty("WAWAI_API_URL", url);
  return { ok: true, message: "URL Wawai API tersimpan di Script Properties." };
}

// Tarik seluruh data pegawai dari Wawai API dan gabungkan ke sheet Users.
// - Pegawai yang SUDAH ada (dicocokkan via NIP) → nama/username diperbarui, Password/Role/IsActive TIDAK disentuh.
// - Pegawai baru (NIP belum ada) → ditambahkan dengan Role=PEGAWAI (netral, belum berfungsi
//   di sistem antrian), Password default "bps2025". Admin perlu mengedit Role-nya menjadi
//   PETUGAS/RESEPSIONIS lewat tab Master Pegawai agar dapat dijadwalkan & melayani antrian.
// - Kalau username hasil generate bentrok dengan akun lain yang BUKAN dari sinkronisasi ini
//   (misalnya superadmin/admin/resepsionis atau seed manual), baris tsb dilewati agar akun
//   yang sudah dipakai untuk login tidak tertimpa.
function syncPegawaiFromWawai() {
  try {
    const url = getWawaiApiUrl_();
    const resp = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    const statusCode = resp.getResponseCode();
    if (statusCode < 200 || statusCode >= 300) {
      return { ok: false, error: "Wawai API mengembalikan status HTTP " + statusCode + "." };
    }

    let json;
    try {
      json = JSON.parse(resp.getContentText());
    } catch (parseErr) {
      return { ok: false, error: "Respons Wawai API bukan JSON yang valid." };
    }
    if (json.success === false) {
      return { ok: false, error: "Wawai API menolak permintaan (cek API key)." };
    }
    const list = json.raw_data;
    if (!Array.isArray(list)) {
      return { ok: false, error: "Format data Wawai API tidak dikenali (field 'raw_data' tidak ditemukan)." };
    }

    const sheet = getOrCreateSheet(SHEET_USERS, ["ID","Nama","Username","Password","Role","IsActive"]);
    const rows = sheet.getDataRange().getValues();
    const existingByUsername = {}; // username/email(lower) -> nomor baris sheet (1-based)
    const existingByNip = {};      // nip_baru/nip_lama -> nomor baris sheet
    for (let i = 1; i < rows.length; i++) {
      const uname = String(rows[i][2] || "").trim().toLowerCase();
      if (uname) existingByUsername[uname] = i + 1;
      const id = String(rows[i][0] || "");
      if (id.indexOf("WAW_") === 0) existingByNip[id.substring(4)] = i + 1;
    }

    // Akun login pegawai hasil sinkronisasi Wawai API TIDAK memakai username buatan dari nama,
    // melainkan:
    //   - Email/Username : "<nip_lama>@bps.go.id"
    //   - Password        : 8 digit pertama dari nip_baru
    // Kedua NIP wajib ada agar kredensial bisa dibentuk; kalau salah satu kosong, baris dilewati
    // (dihitung terpisah sebagai "incomplete" agar kelihatan di pesan hasil sinkronisasi).
    let added = 0, updated = 0, skipped = 0, incomplete = 0;

    list.forEach(function(p) {
      const nipLama = String(p.nip_lama || "").trim();
      const nipBaru = String(p.nip_baru || "").trim();
      const namaLengkap = String(p.nama || "").trim();
      const nipKey = nipBaru || nipLama; // dipakai sebagai kunci pencocokan ID "WAW_<nipKey>"

      if (!namaLengkap || !nipKey) { skipped++; return; }
      if (!nipLama || !nipBaru) { incomplete++; return; } // tidak bisa bentuk email+password lengkap

      const email = nipLama + "@bps.go.id";
      const passwordFromNip = nipBaru.substring(0, 8);

      if (existingByNip[nipKey]) {
        // Sudah pernah disinkronkan sebelumnya → perbarui nama & email (password TIDAK disentuh,
        // supaya kalau pegawai sudah mengganti password sendiri lewat Profil, tidak tertimpa balik).
        const rowNum = existingByNip[nipKey];
        sheet.getRange(rowNum, 2).setValue(namaLengkap);
        const currentUsername = String(rows[rowNum - 1][2] || "").trim().toLowerCase();
        if (currentUsername !== email.toLowerCase() && !existingByUsername[email.toLowerCase()]) {
          sheet.getRange(rowNum, 3).setValue(email);
          existingByUsername[email.toLowerCase()] = rowNum;
        }
        updated++;
      } else if (existingByUsername[email.toLowerCase()]) {
        // Email sudah dipakai akun lain (mis. dibuat manual/seed) — jangan ditimpa.
        skipped++;
      } else {
        const id = "WAW_" + nipKey;
        // Role default PEGAWAI (bukan langsung PETUGAS): data dari Wawai API mencakup SELURUH
        // pegawai lintas jabatan (Pranata Komputer, Analis Anggaran, Kepala Satker, dst), bukan
        // hanya yang bertugas di loket PST. Admin/Superadmin perlu meninjau & mengedit Role
        // pegawai yang relevan menjadi PETUGAS/RESEPSIONIS lewat tab Master Pegawai sebelum
        // mereka bisa dijadwalkan & melayani antrian.
        sheet.appendRow([id, namaLengkap, email, passwordFromNip, "PEGAWAI", true]);
        const newRow = sheet.getLastRow();
        existingByUsername[email.toLowerCase()] = newRow;
        existingByNip[nipKey] = newRow;
        added++;
      }
    });

    return {
      ok: true,
      message: "Sinkronisasi selesai: " + added + " pegawai baru, " + updated + " diperbarui, " +
        skipped + " dilewati, " + incomplete + " dilewati karena NIP tidak lengkap " +
        "(total data dari Wawai API: " + list.length + "). Login pegawai baru: email " +
        "<NIP Lama>@bps.go.id, password 8 digit pertama NIP Baru.",
      added: added, updated: updated, skipped: skipped, incomplete: incomplete, total: list.length
    };
  } catch (e) {
    return { ok: false, error: "Gagal menghubungi Wawai API: " + e.message };
  }
}

// ─── SETUP SELURUH DATA PEGAWAI BPS LAMPUNG ──────────────
// Jalankan fungsi ini SATU KALI di GAS Editor untuk mengisi data awal
function setupAllPegawaiData() {
  setupSpreadsheet();

  setConfig("KANTOR_NAMA", "BPS Provinsi Lampung");
  setConfig("KANTOR_KODE", "LAMPUNG");

  // ── Reset dan isi ulang sheet Users ──
  const userSheet = getSheet(SHEET_USERS);
  userSheet.clearContents();
  userSheet.appendRow(["ID","Nama","Username","Password","Role","IsActive"]);

  // Helper: buat username dari nama (huruf kecil, tanpa gelar, dipisah titik)
  function makeUsername(nama) {
    const clean = nama.split(",")[0].toLowerCase()
      .replace(/[^a-z0-9\s]/g, "")
      .trim()
      .replace(/\s+/g, ".");
    return clean || "user";
  }

  // Data pegawai: [ID, Nama, Password, Role, IsActive] — username akan dibuat otomatis
  // Baris dengan 6 elemen: [ID, Nama, Username, Password, Role, IsActive] — username sudah custom
  const rawPegawai = [
    // ── SUPERADMIN & ADMIN (6 elemen = username custom) ──
    ["USR000",  "Super Admin BPS",    "superadmin", "superadmin", "SUPERADMIN",  true],
    ["USR000B", "Admin BPS Lampung",  "admin",      "admin123",   "ADMIN",       true],
    ["USR000C", "Resepsionis PST",    "resepsionis","bps2025",    "RESEPSIONIS", true],

    // ── DAFTAR LENGKAP PEGAWAI BPS LAMPUNG (5 elemen = username auto) ──
    ["USR001", "A. Yoga Gumbira, SST", "bps2025", "PETUGAS", true],
    ["USR002", "A.Fikri Pazanudin, SST", "bps2025", "PETUGAS", true],
    ["USR003", "Aan Andriatno, A.Md, S.Tr.Stat", "bps2025", "PETUGAS", true],
    ["USR004", "Aan Ardiansyah. AS, SST., M.T.", "bps2025", "PETUGAS", true],
    ["USR005", "Aatina Izzati Penta Harnowati, SE, M.S.Ak.", "bps2025", "PETUGAS", true],
    ["USR006", "Abdul Kadir Jailani", "bps2025", "PETUGAS", true],
    ["USR007", "Abdul Kadir, SE, MM", "bps2025", "PETUGAS", true],
    ["USR008", "Abizar Alghifari, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR009", "Ade Fitriansyah, SST", "bps2025", "PETUGAS", true],
    ["USR010", "Ade Irma Rilyani, A.Md.Kom.", "bps2025", "PETUGAS", true],
    ["USR011", "Adi Irawan, S.E.", "bps2025", "PETUGAS", true],
    ["USR012", "Aditya Anggit Pradika, S.Tr.Stat", "bps2025", "PETUGAS", true],
    ["USR013", "Afifah Dheany Putri, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR014", "Agistin Nafta", "bps2025", "PETUGAS", true],
    ["USR015", "Agung Harry Sakti, A.Md.", "bps2025", "PETUGAS", true],
    ["USR016", "Agus Prasetyo, A.Md.", "bps2025", "PETUGAS", true],
    ["USR017", "Agus Suroso", "bps2025", "PETUGAS", true],
    ["USR018", "Agus Susilo", "bps2025", "PETUGAS", true],
    ["USR019", "Agustini", "bps2025", "PETUGAS", true],
    ["USR020", "Aguswana Saputra", "bps2025", "PETUGAS", true],
    ["USR021", "Ahlul Karom, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR022", "Ahmad Fauzi, SE, MM", "bps2025", "PETUGAS", true],
    ["USR023", "Ahmad Febri Hutama Wiriansyah, SST", "bps2025", "PETUGAS", true],
    ["USR024", "Ahmad Janizar, S.IP", "bps2025", "PETUGAS", true],
    ["USR025", "Ahmad Mutawally, SST", "bps2025", "PETUGAS", true],
    ["USR026", "Ahmad Naufal Suprayogi, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR027", "Ahmad Oki Volda Arlanda, A.Md", "bps2025", "PETUGAS", true],
    ["USR028", "Ahmad Rifjayansyah, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR029", "Ahmad Rifki Febrianto, SST, M.EKK", "bps2025", "PETUGAS", true],
    ["USR030", "Aisyah Azizah Nur Rahmah, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR031", "Aisyah Chairani Putri, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR032", "Aisyah Fahma Ardiana, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR033", "Aisyah Salsabila, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR034", "Ajeng Roro Pangestu, SST., M.E.K.K.", "bps2025", "PETUGAS", true],
    ["USR035", "Ajid Hajiji, SST, M.Si.", "bps2025", "PETUGAS", true],
    ["USR036", "Akhino Yogi Pranata, SST", "bps2025", "PETUGAS", true],
    ["USR037", "Akhmad Riadi", "bps2025", "PETUGAS", true],
    ["USR038", "Akwilla Bahkti Sitindaon, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR039", "Alberto Maradona", "bps2025", "PETUGAS", true],
    ["USR040", "Aldi Rizqul Umam, SST", "bps2025", "PETUGAS", true],
    ["USR041", "Alevareza Dipta Adani, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR042", "Ambo Upek, SE", "bps2025", "PETUGAS", true],
    ["USR043", "Ambriyanto, S.E.", "bps2025", "PETUGAS", true],
    ["USR044", "Amri Kusuma, S.E.", "bps2025", "PETUGAS", true],
    ["USR045", "Ana Destriana, S.Si.", "bps2025", "PETUGAS", true],
    ["USR046", "Andi Candra", "bps2025", "PETUGAS", true],
    ["USR047", "Andi Stiawan, SP", "bps2025", "PETUGAS", true],
    ["USR048", "Andi Yusman", "bps2025", "PETUGAS", true],
    ["USR049", "Andika Nur Budiharso, S.E.", "bps2025", "PETUGAS", true],
    ["USR050", "Andrawina Susanto, S.Si", "bps2025", "PETUGAS", true],
    ["USR051", "Andriazi, SE., MM", "bps2025", "PETUGAS", true],
    ["USR052", "Andries Siregar, SP", "bps2025", "PETUGAS", true],
    ["USR053", "Angga Jati Batara, SST", "bps2025", "PETUGAS", true],
    ["USR054", "Anggi Budi Pratiwi, S.Stat.", "bps2025", "PETUGAS", true],
    ["USR055", "Anggia Wulandari Harahap, S.Psi", "bps2025", "PETUGAS", true],
    ["USR056", "Anisa Ramadhani, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR057", "Anita Desmarini, SST", "bps2025", "PETUGAS", true],
    ["USR058", "Anne Oktavia Andriyani, A.Md.", "bps2025", "PETUGAS", true],
    ["USR059", "Annisa Arum Adiza, A.Md.Kb.N.", "bps2025", "PETUGAS", true],
    ["USR060", "Annisa Fauziatul Mardiyah, SST", "bps2025", "PETUGAS", true],
    ["USR061", "Aprilia Puspita Sari, SST", "bps2025", "PETUGAS", true],
    ["USR062", "Ardelia Raras Nisreyasa, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR063", "Ardhan Setyo Nugroho, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR064", "Ardi Yansyah", "bps2025", "PETUGAS", true],
    ["USR065", "Ari Arnantyo, SE", "bps2025", "PETUGAS", true],
    ["USR066", "Ari Rusmasari, SST, M.Si.", "bps2025", "PETUGAS", true],
    ["USR067", "Aria Jaya Atmaja, S.E.", "bps2025", "PETUGAS", true],
    ["USR068", "Arief Novriansyah, S.M.", "bps2025", "PETUGAS", true],
    ["USR069", "Arief Rahmanda Al-Mursyid, SST", "bps2025", "PETUGAS", true],
    ["USR070", "Arif Eko Handoko, S.E., M.M.", "bps2025", "PETUGAS", true],
    ["USR071", "Arif Rahman Maulana, S.Si, M.M.", "bps2025", "PETUGAS", true],
    ["USR072", "Arifin Jafar, SST, M.T.", "bps2025", "PETUGAS", true],
    ["USR073", "Arip Purwanto, S.M.", "bps2025", "PETUGAS", true],
    ["USR074", "Aris Munandar", "bps2025", "PETUGAS", true],
    ["USR075", "Aris Priasetya Utama, S.Si.", "bps2025", "PETUGAS", true],
    ["USR076", "Ariyanto", "bps2025", "PETUGAS", true],
    ["USR077", "Ariyanto, SST., M.M.", "bps2025", "PETUGAS", true],
    ["USR078", "Arlien Harlikah Pisananti, SST", "bps2025", "PETUGAS", true],
    ["USR079", "Aroki Sutanto Mandala Yudha, A.Md.", "bps2025", "PETUGAS", true],
    ["USR080", "Arum Pratiwi, SST", "bps2025", "PETUGAS", true],
    ["USR081", "Arum Purbowati, SST", "bps2025", "PETUGAS", true],
    ["USR082", "Asita Sekar Asri, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR083", "Asri Kurniawan, S.M.", "bps2025", "PETUGAS", true],
    ["USR084", "Astryati Marganingsih, SST", "bps2025", "PETUGAS", true],
    ["USR085", "Atik Heriyandani, SE", "bps2025", "PETUGAS", true],
    ["USR086", "Auliya Jamiatus Saufi, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR087", "Austin Abdul Aziz Riano Junior, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR088", "Ayu Setianingsih, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR089", "Ayu Winarni, S.Si", "bps2025", "PETUGAS", true],
    ["USR090", "Ayu Yudestia Pratiwi, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR091", "Azelea Juliastita, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR092", "Azmi Zulfani Putri, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR093", "Bagas Ashari, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR094", "Bangsawan, SST, M.Si.", "bps2025", "PETUGAS", true],
    ["USR095", "Bayu Juniardi, SE", "bps2025", "PETUGAS", true],
    ["USR096", "Bayu Rahmawan, A.Md", "bps2025", "PETUGAS", true],
    ["USR097", "Belinda Yena Putri, A.Md.Kom", "bps2025", "PETUGAS", true],
    ["USR098", "Beni Siswanto, A.Md.", "bps2025", "PETUGAS", true],
    ["USR099", "Bima Ghafara, A.Md", "bps2025", "PETUGAS", true],
    ["USR100", "Bima Prastya Handawi, S.I.Kom.", "bps2025", "PETUGAS", true],
    ["USR101", "Bimo Ade Budiman Fikri, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR102", "Cahyo Novi Tri Darto, S.E.", "bps2025", "PETUGAS", true],
    ["USR103", "Carissa Rabbanirmala, SST", "bps2025", "PETUGAS", true],
    ["USR104", "Chintia Arisandi Hidayat, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR105", "Citra Tantri Widyaningtyas, S.Stat.", "bps2025", "PETUGAS", true],
    ["USR106", "Clara Tridiana, SST, M.S.E", "bps2025", "PETUGAS", true],
    ["USR107", "Dadang Supriatna", "bps2025", "PETUGAS", true],
    ["USR108", "Dama Yanti", "bps2025", "PETUGAS", true],
    ["USR109", "Darmansyah", "bps2025", "PETUGAS", true],
    ["USR110", "Darul Ambardi, SE", "bps2025", "PETUGAS", true],
    ["USR111", "Darwin Amri, S.P.", "bps2025", "PETUGAS", true],
    ["USR112", "Deanis Camelia Anugrah Putri, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR113", "Dechi Yulpratiwi, SST, M.E.", "bps2025", "PETUGAS", true],
    ["USR114", "Dede Apriyawan, SST", "bps2025", "PETUGAS", true],
    ["USR115", "Dedi Triatmoko, A.Md.", "bps2025", "PETUGAS", true],
    ["USR116", "Defika, SE", "bps2025", "PETUGAS", true],
    ["USR117", "Defki Suherman", "bps2025", "PETUGAS", true],
    ["USR118", "Dekka Safutra, A.Md.", "bps2025", "PETUGAS", true],
    ["USR119", "Della Krisna Prayuga, S.M.", "bps2025", "PETUGAS", true],
    ["USR120", "Delvi Rutania Prama, S.Si., M.Stat.", "bps2025", "PETUGAS", true],
    ["USR121", "Delvira Cindy Rosmilda, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR122", "Deni Setiawan, A.Md", "bps2025", "PETUGAS", true],
    ["USR123", "Desliyani Tri Wandita, SST, M.E.", "bps2025", "PETUGAS", true],
    ["USR124", "Desyana Khotiah", "bps2025", "PETUGAS", true],
    ["USR125", "Devi Novanti, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR126", "Devi Pratiwi, SST", "bps2025", "PETUGAS", true],
    ["USR127", "Devie Triana Sari, S.Si, M.P.W.K.", "bps2025", "PETUGAS", true],
    ["USR128", "Dewi Ratih Ranggawuni, S.ST", "bps2025", "PETUGAS", true],
    ["USR129", "Dewi Wahyuningsih, SST., M.Si.", "bps2025", "PETUGAS", true],
    ["USR130", "Dewi Yuliana S., S.T.", "bps2025", "PETUGAS", true],
    ["USR131", "Dhimas Reza Destian", "bps2025", "PETUGAS", true],
    ["USR132", "Dhyantanu Harsa, SST., MM", "bps2025", "PETUGAS", true],
    ["USR133", "Diah Hadianing Putri, S.Si", "bps2025", "PETUGAS", true],
    ["USR134", "Diah Septia Ningrum, A.Md.Kom", "bps2025", "PETUGAS", true],
    ["USR135", "Dian Wuryandari Syafiatin, S.Si., M.M.", "bps2025", "PETUGAS", true],
    ["USR136", "Dicky Suharto", "bps2025", "PETUGAS", true],
    ["USR137", "Dina Anggraena, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR138", "Dini Alfitri Zahra, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR139", "Dini Arianindy, S.Tr.Stat", "bps2025", "PETUGAS", true],
    ["USR140", "Dinny Pravitasari, SST, M.S.E.", "bps2025", "PETUGAS", true],
    ["USR141", "Dira Annisa Nasution, SST", "bps2025", "PETUGAS", true],
    ["USR142", "Dita Christina Simorangkir, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR143", "Dita Regi Kartika, S.Tr.Stat", "bps2025", "PETUGAS", true],
    ["USR144", "Diza Potenza, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR145", "Dr. Ahmadriswan Nasution, S.Si, M.T", "bps2025", "PETUGAS", true],
    ["USR146", "Dr. Hady Suryono, M.Si.", "bps2025", "PETUGAS", true],
    ["USR147", "Drisnaf Swastyardi, S.Si, MSE, MA.", "bps2025", "PETUGAS", true],
    ["USR148", "Dwi Ambarningrum, S.A.N., M.M.", "bps2025", "PETUGAS", true],
    ["USR149", "Dwi Astuti, SH", "bps2025", "PETUGAS", true],
    ["USR150", "Dwi Erna Susilowati, S.M.", "bps2025", "PETUGAS", true],
    ["USR151", "Dwi Hendro Siswo Purnomo, SST", "bps2025", "PETUGAS", true],
    ["USR152", "Dwi Octharianty Fauziah, A.Md.", "bps2025", "PETUGAS", true],
    ["USR153", "Ebi Heriyanto, S.M.", "bps2025", "PETUGAS", true],
    ["USR154", "Edi Ristanto, A.Md.", "bps2025", "PETUGAS", true],
    ["USR155", "Eduan Mirza", "bps2025", "PETUGAS", true],
    ["USR156", "Edy Santoso", "bps2025", "PETUGAS", true],
    ["USR157", "Eka Pratiwi Arnaz, SST", "bps2025", "PETUGAS", true],
    ["USR158", "Eka Yusda, SE", "bps2025", "PETUGAS", true],
    ["USR159", "Eklesia Valentia, A.Md.Kb.N.", "bps2025", "PETUGAS", true],
    ["USR160", "Eko Guswanto", "bps2025", "PETUGAS", true],
    ["USR161", "Eko Purnomo, SST., MM", "bps2025", "PETUGAS", true],
    ["USR162", "Eko Rahmat", "bps2025", "PETUGAS", true],
    ["USR163", "Eko Teguh Widodo, SST., M.Sc.", "bps2025", "PETUGAS", true],
    ["USR164", "Eko Wardiyanto", "bps2025", "PETUGAS", true],
    ["USR165", "Elok Nurfatika Gunawan, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR166", "Elpa Heryanti, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR167", "Elva Dwi Kristiani, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR168", "Erfan Kurniansah", "bps2025", "PETUGAS", true],
    ["USR169", "Erika Haryulistiani Saksono, S.E., M.E.", "bps2025", "PETUGAS", true],
    ["USR170", "Erika Santi, ST, M.Si", "bps2025", "PETUGAS", true],
    ["USR171", "Erisa Nur Imama, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR172", "Erizal, S.M.", "bps2025", "PETUGAS", true],
    ["USR173", "Erma Fitriana, SST", "bps2025", "PETUGAS", true],
    ["USR174", "Erry Roma Nazaruddin, S.M.", "bps2025", "PETUGAS", true],
    ["USR175", "Ervina Yolanda, SST", "bps2025", "PETUGAS", true],
    ["USR176", "Erwan Jafrilda, S.Sos", "bps2025", "PETUGAS", true],
    ["USR177", "Erwandi, S.Si, M.Si", "bps2025", "PETUGAS", true],
    ["USR178", "Erwansyah Yusup", "bps2025", "PETUGAS", true],
    ["USR179", "Esa Anindika Sari, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR180", "Eva Nirwana, SST", "bps2025", "PETUGAS", true],
    ["USR181", "Evie Ermawati, SST, M.M.", "bps2025", "PETUGAS", true],
    ["USR182", "Fadheel Wisnu Utomo, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR183", "Fahmi Muhammad Sahal, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR184", "Fahroni Agustarita, SE", "bps2025", "PETUGAS", true],
    ["USR185", "Faiz Alwan Alharis, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR186", "Fajar Berlian, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR187", "Fajar Mustaqim, S.Si., M.Si.", "bps2025", "PETUGAS", true],
    ["USR188", "Fajriana Fadhlul Fikri, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR189", "Fani Dwi Putra Rahmadhany, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR190", "Fanisa Dwita Hanggarani, SST", "bps2025", "PETUGAS", true],
    ["USR191", "Fara Yulia Fransiska, S.Tr.Stat", "bps2025", "PETUGAS", true],
    ["USR192", "Faradilla Anastasya, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR193", "Farda Zayana Majid, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR194", "Farid Indrajaya, SE", "bps2025", "PETUGAS", true],
    ["USR195", "Faza Nur Fuadina, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR196", "Fazani", "bps2025", "PETUGAS", true],
    ["USR197", "Febbial Qausar", "bps2025", "PETUGAS", true],
    ["USR198", "Febiyana Qomariyah, SST., M.M.", "bps2025", "PETUGAS", true],
    ["USR199", "Febriana Dwi Jayanti, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR200", "Febriana Susiwi, SKM, M.Si", "bps2025", "PETUGAS", true],
    ["USR201", "Febyola Ayu Martha, A.Md", "bps2025", "PETUGAS", true],
    ["USR202", "Ferry", "bps2025", "PETUGAS", true],
    ["USR203", "Ferry Irawan", "bps2025", "PETUGAS", true],
    ["USR204", "Fidya Alhayu, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR205", "Firman Ardhiana Putra, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR206", "Firmansyah, A.Md.", "bps2025", "PETUGAS", true],
    ["USR207", "Firna Novi Anggoro, SH, M.H.", "bps2025", "PETUGAS", true],
    ["USR208", "Fithriyah, SST", "bps2025", "PETUGAS", true],
    ["USR209", "Fitra Saputra, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR210", "Fitri Nurjanah, S.E., M.M.", "bps2025", "PETUGAS", true],
    ["USR211", "Fransiska Suryani, S.M.", "bps2025", "PETUGAS", true],
    ["USR212", "Frida Jubilate Hutabarat, SST", "bps2025", "PETUGAS", true],
    ["USR213", "Ganes Angga Witata, SE., M.Kom.", "bps2025", "PETUGAS", true],
    ["USR214", "Ghessani Nurina Khaq, SST", "bps2025", "PETUGAS", true],
    ["USR215", "Gigih Agus Susiyanto, SST, M.E.K.K", "bps2025", "PETUGAS", true],
    ["USR216", "Gilang Abdul Jabbar, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR217", "Gindo Nainggolan, A.Md.", "bps2025", "PETUGAS", true],
    ["USR218", "Gun Gun Nugraha, S.Si, M.S.E", "bps2025", "PETUGAS", true],
    ["USR219", "Gunawan Catur Prasetyo, SST, M.Si.", "bps2025", "PETUGAS", true],
    ["USR220", "Guruh Dewa Prataba, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR221", "Gusmalia Sari, SE", "bps2025", "PETUGAS", true],
    ["USR222", "Habni Hamara Azmatiy, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR223", "Hamzah Saifudin, SST", "bps2025", "PETUGAS", true],
    ["USR224", "Hana Fitianingrum, S.Si.", "bps2025", "PETUGAS", true],
    ["USR225", "Hanik Devianingrum, SST", "bps2025", "PETUGAS", true],
    ["USR226", "Hardianty, S.Si", "bps2025", "PETUGAS", true],
    ["USR227", "Haris Sugara, S.E.", "bps2025", "PETUGAS", true],
    ["USR228", "Hariyanto", "bps2025", "PETUGAS", true],
    ["USR229", "Haston Handono", "bps2025", "PETUGAS", true],
    ["USR230", "Hendarwanto, SE", "bps2025", "PETUGAS", true],
    ["USR231", "Hendrit Lolly Pradikta, S.M.", "bps2025", "PETUGAS", true],
    ["USR232", "Henida Widyatama, SST", "bps2025", "PETUGAS", true],
    ["USR233", "Henny Surya Indraswari, SST., M.Si", "bps2025", "PETUGAS", true],
    ["USR234", "Heri Fitriyanto", "bps2025", "PETUGAS", true],
    ["USR235", "Herman", "bps2025", "PETUGAS", true],
    ["USR236", "Hermansyah", "bps2025", "PETUGAS", true],
    ["USR237", "Hermanto", "bps2025", "PETUGAS", true],
    ["USR238", "Hermawan Prasetyo, SST, MT.", "bps2025", "PETUGAS", true],
    ["USR239", "Herny Wahyuni, SE, M.M.", "bps2025", "PETUGAS", true],
    ["USR240", "Heru Wijayanto", "bps2025", "PETUGAS", true],
    ["USR241", "Hesti Ayuningtiyas, SST", "bps2025", "PETUGAS", true],
    ["USR242", "Hetty Rahayu, A.Md", "bps2025", "PETUGAS", true],
    ["USR243", "Id Arter, SE.", "bps2025", "PETUGAS", true],
    ["USR244", "Ihsandi Vidi Ul Awal, SST", "bps2025", "PETUGAS", true],
    ["USR245", "Ihza Mahendra, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR246", "Iin Aryani, A.Md", "bps2025", "PETUGAS", true],
    ["USR247", "Ika Nurfadhilah, S.M.", "bps2025", "PETUGAS", true],
    ["USR248", "Ika Wahyuni Novianti, A.Md", "bps2025", "PETUGAS", true],
    ["USR249", "Ikhsan, SE", "bps2025", "PETUGAS", true],
    ["USR250", "Imam Sujono, S.Tr.Stat", "bps2025", "PETUGAS", true],
    ["USR251", "Imanullah Achmad, SE", "bps2025", "PETUGAS", true],
    ["USR252", "Imas Gita Anakis", "bps2025", "PETUGAS", true],
    ["USR253", "Imron", "bps2025", "PETUGAS", true],
    ["USR254", "Imron Rosyadi, S.E.", "bps2025", "PETUGAS", true],
    ["USR255", "Inah Angelia Saragih, S.Si", "bps2025", "PETUGAS", true],
    ["USR256", "Indra Gunawan", "bps2025", "PETUGAS", true],
    ["USR257", "Indra Kurniawan, S.Sos.", "bps2025", "PETUGAS", true],
    ["USR258", "Indra, S.Si.", "bps2025", "PETUGAS", true],
    ["USR259", "Indri Puspitasari, S.Stat.", "bps2025", "PETUGAS", true],
    ["USR260", "Intan Eka Debora Silaban, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR261", "Iqbal Maulana, A.Md.Stat.", "bps2025", "PETUGAS", true],
    ["USR262", "Ir. Herum Fajarwati, M.M", "bps2025", "PETUGAS", true],
    ["USR263", "Ir. Sri Rezkie Desmawati, ME", "bps2025", "PETUGAS", true],
    ["USR264", "Irvan Patuan Marsahala Simamora, SST", "bps2025", "PETUGAS", true],
    ["USR265", "Irwan, S.E.", "bps2025", "PETUGAS", true],
    ["USR266", "Ismail Yusuf, SE", "bps2025", "PETUGAS", true],
    ["USR267", "Isna Syafira Khalif Ilma, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR268", "Isnaeni Nur Khasanah, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR269", "Istiqlal", "bps2025", "PETUGAS", true],
    ["USR270", "Istu Indah Setyaningsih, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR271", "Jantika Ayu Ramadhani, A.Md.Kom", "bps2025", "PETUGAS", true],
    ["USR272", "Jemali", "bps2025", "PETUGAS", true],
    ["USR273", "Johan Darmawan, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR274", "Johan Irawan, SE.", "bps2025", "PETUGAS", true],
    ["USR275", "John Fernando Silaban, SE", "bps2025", "PETUGAS", true],
    ["USR276", "John Knedi, S.Si, MM", "bps2025", "PETUGAS", true],
    ["USR277", "Joni Sampurno, A.Md, S.M.", "bps2025", "PETUGAS", true],
    ["USR278", "Joseph Gabriel Napitupulu, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR279", "Jua Mahardhika, SST, M.M.", "bps2025", "PETUGAS", true],
    ["USR280", "Juliadi", "bps2025", "PETUGAS", true],
    ["USR281", "Jusman, S.M.", "bps2025", "PETUGAS", true],
    ["USR282", "K. Nurika Damayanti, SST., M.Stat.", "bps2025", "PETUGAS", true],
    ["USR283", "Kaisar Samudra, S.E.", "bps2025", "PETUGAS", true],
    ["USR284", "Khalida Hanum, S.Si.", "bps2025", "PETUGAS", true],
    ["USR285", "Lailatus Syifa Aulianisa, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR286", "Lakoni Husin, A.Md", "bps2025", "PETUGAS", true],
    ["USR287", "Lara Ayu Cahyaningtyas, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR288", "Lenna Nalurita, SP", "bps2025", "PETUGAS", true],
    ["USR289", "Lia Gustina, SST, M.E.K.K", "bps2025", "PETUGAS", true],
    ["USR290", "Lisa Erlinawati, S.Pd.", "bps2025", "PETUGAS", true],
    ["USR291", "Lisawati, S. I. Kom., M.I.kom", "bps2025", "PETUGAS", true],
    ["USR292", "Lissa Octavia Wardana, S.Tr.Stat.", "bps2025", "PETUGAS", true],
    ["USR293", "Ludi Indra, S.H.", "bps2025", "PETUGAS", true],
    ["USR294", "Lukman Hadiansyah, S.M.", "bps2025", "PETUGAS", true],
    ["USR295", "Luluk Tanjung Wahyuningsih, SE", "bps2025", "PETUGAS", true],
    ["USR296", "Lulut Ariyadi", "bps2025", "PETUGAS", true],
    ["USR297", "Lusi Oktari, S.Si., M.M", "bps2025", "PETUGAS", true],
    ["USR298", "Luxmaning Hutaki Widiastari, S.Psi., M.Psi.T.", "bps2025", "PETUGAS", true],
    ["USR299", "M. Ali Rizky", "bps2025", "PETUGAS", true],
    ["USR300", "M. Meiru Panca Rezki, S.Tr.Stat.", "bps2025", "PETUGAS", true]
  ];

  rawPegawai.forEach(function(r) {
    if (r.length === 6) {
      // Sudah punya username custom (superadmin, admin, resepsionis)
      userSheet.appendRow(r);
    } else {
      // Generate username otomatis dari nama
      userSheet.appendRow([r[0], r[1], makeUsername(r[1]), r[2], r[3], r[4]]);
    }
  });

  // ── Reset PetugasRating ──
  const prSheet = getOrCreateSheet("PetugasRating",
    ["ID","EntryID","PetugasNama","Score","Comment","Tanggal"]);
  prSheet.clearContents();
  prSheet.appendRow(["ID","EntryID","PetugasNama","Score","Comment","Tanggal"]);

  // ── Reset Queue dan isi data dummy hari ini ──
  const qSheet = getSheet(SHEET_QUEUE);
  qSheet.clearContents();
  qSheet.appendRow([
    "ID","Ticket","Prefix","Number","NamaPengunjung","Tujuan","Detail",
    "InstansiAlamat","JumlahOrang","Status","Petugas","ServiceType",
    "ServiceDetail","Rating","RatingComment","Timestamp","Tanggal"
  ]);

  const today = todayStr();

  // Data dummy antrian hari ini (format tiket baru: prefix mengikuti jenis layanan A/B/C/D, RAPAT/LAINNYA pakai "U")
  const dummyQueue = [
    ["DUM001","B-001","B",1,"Thomson Sibuea","PST","Konsultasi PDRB Tahun 2025",
     "Universitas Lampung",1,"DONE","Ahlul Karom, S.Tr.Stat.","B - Konsultasi Statistik",
     "",5,"Pelayanan sangat memuaskan!",
     today+" 08:15:00", today],
    ["DUM002","C-001","C",1,"Siti Nurhaliza","PST","Data Inflasi 2024",
     "Dinas Perindustrian Lampung",2,"DONE","Dewi Wahyuningsih, SST., M.Si.","C - Produk Statistik Berbayar",
     "",4,"Mantap, cepat dan ramah",
     today+" 08:45:00", today],
    ["DUM003","D-001","D",1,"Budi Santoso","PST","Rekomendasi survei mandiri",
     "PT Sinar Mas Lampung",1,"DONE","Ahmad Rifki Febrianto, SST, M.EKK","D - Rekomendasi Statistik",
     "","","",
     today+" 09:10:00", today],
    ["DUM004","B-002","B",2,"Ratna Dewi","PST","Konsultasi data kemiskinan",
     "Pemerintah Kota Metro",2,"IN_SERVICE","Ahlul Karom, S.Tr.Stat.","","",
     "","",
     today+" 09:30:00", today],
    ["DUM005","A-001","A",1,"Hendra Gunawan","PST","Data kependudukan 2020",
     "BPS Kabupaten Pesawaran",1,"WAITING","","","","","",
     today+" 09:45:00", today],
    ["DUM006","U-001","U",1,"Annisa Rahmawati","RAPAT","Koordinasi survei pertanian",
     "Dinas Pertanian Provinsi Lampung",3,"WAITING","","","","","",
     today+" 10:00:00", today],
    ["DUM007","A-002","A",2,"Farhan Maulana","PST","Tabel IO Lampung",
     "Bappeda Provinsi Lampung",1,"WAITING","","","","","",
     today+" 10:15:00", today],
    ["DUM008","U-002","U",2,"Yuni Astuti","LAINNYA","Pengambilan dokumen statistik",
     "Kanwil Kemenkeu Lampung",1,"WAITING","","","","","",
     today+" 10:30:00", today],
  ];

  dummyQueue.forEach(function(r) { qSheet.appendRow(r); });

  // Set counter awal per-prefix agar tidak bentrok dengan data dummy di atas
  setConfig("COUNTER_A", "2"); setConfig("COUNTER_DATE_A", today);
  setConfig("COUNTER_B", "2"); setConfig("COUNTER_DATE_B", today);
  setConfig("COUNTER_C", "1"); setConfig("COUNTER_DATE_C", today);
  setConfig("COUNTER_D", "1"); setConfig("COUNTER_DATE_D", today);
  setConfig("COUNTER_U", "2"); setConfig("COUNTER_DATE_U", today);

  // Tambah data rating
  prSheet.appendRow(["PR001","DUM001","Ahlul Karom, S.Tr.Stat.",5,"Pelayanan sangat memuaskan!", today]);
  prSheet.appendRow(["PR002","DUM002","Dewi Wahyuningsih, SST., M.Si.",4,"Mantap, cepat dan ramah", today]);

  return { ok: true, message: "Berhasil! Data " + rawPegawai.length + " pegawai + antrian dummy sudah dimasukkan ke Google Sheets." };
}

// Alias agar bisa dijalankan dengan nama lain
function setupDummyData() {
  return setupAllPegawaiData();
}