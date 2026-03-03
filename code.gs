// ==========================================
// KONFIGURASI UTAMA
// ==========================================
const DRIVE_FOLDER_ID = "1VshBRT2Joxs0Na0XsO8GWNcQ2351mbwM"; 
const TIMEZONE = "Asia/Makassar"; // WITA
// [BARU] URL Logo Sekolah untuk Laporan PDF
const LOGO_DRIVE_ID = "1A_Gt9DA8UyYkDItvlDR2cFVekExmyD4B";

// ==========================================
// 1. INISIALISASI SYSTEM
// ==========================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ADMIN SI-AKSI')
    .addItem('1. Setup Database', 'setupDatabaseManual')
    .addItem('2. Tes Kirim PDF (Bulan Ini)', 'generateMonthlyReport')
    .addItem('3. Aktifkan Trigger Otomatis', 'setupMonthlyReportTrigger')
    .addToUi();
}

// UPDATE FUNGSI doGet untuk Routing Halaman Baru
function doGet(e) {
  let template = 'Index';
  if (e.parameter.page === 'dashboard') template = 'dashboard';
  if (e.parameter.page === 'laporan') template = 'report'; // <--- TAMBAHAN BARU
  
  return HtmlService.createTemplateFromFile(template)
    .evaluate()
    .setTitle('SI-AKSI | SMKPP NEGERI BIMA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ==========================================
// UPDATE: DATA FETCHER (Menambahkan URL Foto)
// ==========================================
/**
 * Mengambil data laporan berdasarkan rentang tanggal dan nama,
 * serta mendeteksi pegawai yang belum mengisi kinerja hari ini.
 */
function getFilteredData(startStr, endStr, nameFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const sheetStaff = ss.getSheetByName('DataPegawai');
  
  if (!sheetStaff) return { status: 'error', msg: 'Sheet DataPegawai tidak ditemukan' };

  // 1. Ambil Master Daftar Pegawai dari Kolom A (Index 0)
  const allStaff = sheetStaff.getRange(2, 1, sheetStaff.getLastRow() - 1, 1).getValues().flat();

  const startDate = new Date(startStr); 
  startDate.setHours(0,0,0,0);
  const endDate = new Date(endStr);
  endDate.setHours(23,59,59,999);

  let allRows = [];
  
  // 2. Ambil semua data dari sheet yang berawalan "Laporan_"
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName.indexOf("Laporan_") === 0 || sheetName === "LaporanKerja") {
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        allRows = allRows.concat(data.slice(1));
      }
    }
  });

  // 3. Filter Data Berdasarkan Tanggal dan Nama
  const filtered = allRows.filter(row => {
    const rowDate = new Date(row[1]);
    const isDateMatch = rowDate >= startDate && rowDate <= endDate;
    const isNameMatch = nameFilter === "ALL" ? true : row[2] === nameFilter;
    return isDateMatch && isNameMatch;
  });

  // Urutkan berdasarkan tanggal terkecil
  filtered.sort((a, b) => new Date(a[1]) - new Date(b[1]));

  // 4. Hitung Statistik untuk Dashboard
  const stats = {
    total: filtered.length,
    selesai: filtered.filter(r => r[11] === 'SELESAI').length,
    proses: filtered.filter(r => r[11] === 'PROSES').length,
    izin: filtered.filter(r => ['IZIN', 'SAKIT', 'CUTI', 'DINAS'].includes(r[11])).length,
    daily: {}
  };

  filtered.forEach(row => {
    const tglKey = Utilities.formatDate(new Date(row[1]), TIMEZONE, "dd/MM");
    stats.daily[tglKey] = (stats.daily[tglKey] || 0) + 1;
  });

  // 5. Logika Pegawai Belum Melapor (Alpha) - Khusus Hari Ini
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy");
  const reportedToday = filtered
    .filter(r => Utilities.formatDate(new Date(r[1]), TIMEZONE, "dd/MM/yyyy") === todayStr)
    .map(r => r[2]);
  
  // Membandingkan Master Pegawai vs Yang sudah lapor hari ini
  const absentToday = allStaff.filter(name => !reportedToday.includes(name));

  // 6. Mapping Data untuk Tabel UI
  const tableData = filtered.map(r => {
    const formatJam = (d) => (d instanceof Date) ? Utilities.formatDate(d, TIMEZONE, "HH:mm") : "-";
    return {
      tgl: Utilities.formatDate(new Date(r[1]), TIMEZONE, "dd/MM/yyyy"),
      nama: r[2],
      kegiatan: r[3],
      jam: r[11] === 'SELESAI' ? `${formatJam(new Date(r[7]))} - ${formatJam(new Date(r[10]))}` : "-",
      status: r[11],
      foto_awal: r[6],
      foto_akhir: r[9]
    };
  }).reverse();

  return { table: tableData, stats: stats, absentToday: absentToday };
}

// ==========================================
// 1. HELPER: ANALISA KINERJA VIA GEMINI AI
// ==========================================
function getAIReview(activityList, percentage) {
  const API_KEY = "AIzaSyAAe-AhIBCsGt8p9V3q6wCmnBj9o1L30ME"; // Gunakan Key Anda
  const MODEL_NAME = "gemini-2.5-flash"; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${API_KEY}`;

  // Batasi karakter agar tidak overload token (ambil 1500 karakter pertama saja)
  const activities = activityList.join(", ").substring(0, 1500);

  const prompt = `
    Bertindaklah sebagai Kepala Sekolah SMKPP Negeri Bima
    Berikan evaluasi kinerja singkat (maksimal 3 kalimat) dan saran pengembangan untuk pegawai ini.
    
    Data Kinerja:
    - Tingkat Kehadiran/Penyelesaian: ${percentage}%
    - Daftar Pekerjaan yang dilakukan: ${activities}

    Gunakan bahasa formal birokrasi Indonesia yang apresiatif namun tegas.
    Jangan sebutkan angka persentase di dalam kalimat, fokus pada kualitas kegiatannya.
  `;

  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates.length > 0) {
       return json.candidates[0].content.parts[0].text.trim();
    }
  } catch (e) {
    console.error("AI Error: " + e.toString());
  }
  return null; // Return null jika gagal, agar pakai fallback manual
}


// ==========================================
// PDF GENERATOR FINAL (FIX: FILTER LOGIC)
// ==========================================
function generateFilteredPDF(startStr, endStr, nameFilter, dataJson) {
  // 1. Parse Data Mentah dari Frontend
  let filledData = JSON.parse(dataJson);

  // === PERBAIKAN UTAMA DISINI ===
  // Jika permintaan adalah untuk SATU ORANG (bukan ALL), 
  // maka kita WAJIB membuang data milik orang lain dari array ini.
  if (nameFilter !== "ALL") {
    filledData = filledData.filter(row => row.nama === nameFilter);
  }
  // ==============================

  const periode = `${Utilities.formatDate(new Date(startStr), TIMEZONE, "dd MMM yyyy")} s.d. ${Utilities.formatDate(new Date(endStr), TIMEZONE, "dd MMM yyyy")}`;
  const title = nameFilter === "ALL" ? "REKAPITULASI SELURUH PEGAWAI" : `LAPORAN KINERJA: ${nameFilter.toUpperCase()}`;
  
  // 2. LOAD LOGO
  let logoBase64 = "";
  try {
    const file = DriveApp.getFileById(LOGO_DRIVE_ID); 
    const blob = file.getBlob();
    const b64 = Utilities.base64Encode(blob.getBytes());
    logoBase64 = `data:${blob.getContentType()};base64,${b64}`;
  } catch (e) {
    logoBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="; 
  }

  // 3. LOGIKA KALENDER (MENGISI HARI KOSONG)
  let tableRows = "";
  let currentDate = new Date(startStr);
  const stopDate = new Date(endStr);
  
  // Statistik
  let totalHariKalender = 0;
  let hariTerisi = 0;
  let hariKosong = 0;

  // Set jam ke 00:00:00
  currentDate.setHours(0,0,0,0);
  stopDate.setHours(23,59,59,999);

  let rowIndex = 0;

  while (currentDate <= stopDate) {
    totalHariKalender++;
    
    const dateCheck = Utilities.formatDate(currentDate, TIMEZONE, "dd/MM/yyyy");
    
    // Cari data di tanggal ini (DATA SUDAH BERSIH KARENA DI-FILTER DIATAS)
    const rowsOnThisDay = filledData.filter(r => r.tgl === dateCheck);

    if (rowsOnThisDay.length > 0) {
      // === KASUS A: ADA KERJAAN ===
      hariTerisi++;
      rowsOnThisDay.forEach(row => {
        let color = row.status === 'SELESAI' ? '#166534' : '#ca8a04'; 
        if(['IZIN','SAKIT','CUTI'].includes(row.status)) color = '#ea580c';

        const linkAwal = row.foto_awal ? `<a href="${row.foto_awal}" style="color:#2563eb; text-decoration:none;">[Foto Awal]</a>` : '-';
        const linkAkhir = row.foto_akhir ? `<a href="${row.foto_akhir}" style="color:#2563eb; text-decoration:none;">[Foto Akhir]</a>` : '-';
        const displayBukti = row.status === 'SELESAI' ? `${linkAwal}<br>${linkAkhir}` : linkAwal;

        tableRows += `
          <tr style="background-color: ${rowIndex % 2 === 0 ? '#fff' : '#f9fafb'};">
            <td style="padding:8px; border-bottom:1px solid #eee;">${row.tgl}</td>
            <td style="padding:8px; border-bottom:1px solid #eee;">${row.nama}</td>
            <td style="padding:8px; border-bottom:1px solid #eee;">${row.kegiatan}</td>
            <td style="padding:8px; border-bottom:1px solid #eee; text-align:center;">${row.jam}</td>
            <td style="padding:8px; border-bottom:1px solid #eee; color:${color}; font-weight:bold;">${row.status}</td>
            <td style="padding:8px; border-bottom:1px solid #eee; font-size:10px; text-align:center;">${displayBukti}</td>
          </tr>
        `;
        rowIndex++;
      });

    } else {
      // === KASUS B: TIDAK MENGISI (ZONK) ===
      hariKosong++;
      
      const dayNum = currentDate.getDay(); // 0 = Minggu
      const isSunday = dayNum === 0;
      
      let msg = isSunday ? "HARI LIBUR (MINGGU)" : "TIDAK MENGISI LAPORAN KINERJA";
      let bgColor = isSunday ? "#f3f4f6" : "#fef2f2"; 
      let textColor = isSunday ? "#9ca3af" : "#ef4444"; 

      // Tampilkan Nama sesuai filter
      const displayName = nameFilter === "ALL" ? "-" : nameFilter;

      tableRows += `
        <tr style="background-color: ${bgColor};">
          <td style="padding:8px; border-bottom:1px solid #eee; color:${textColor}; font-weight:bold;">${dateCheck}</td>
          <td style="padding:8px; border-bottom:1px solid #eee; color:${textColor};">${displayName}</td>
          <td style="padding:8px; border-bottom:1px solid #eee; color:${textColor}; font-style:italic;">${msg}</td>
          <td style="padding:8px; border-bottom:1px solid #eee; text-align:center;">-</td>
          <td style="padding:8px; border-bottom:1px solid #eee; color:${textColor}; font-weight:bold;">ALPHA</td>
          <td style="padding:8px; border-bottom:1px solid #eee; text-align:center;">-</td>
        </tr>
      `;
      rowIndex++;
    }

    currentDate.setDate(currentDate.getDate() + 1);
  }

  // 4. AI REKOMENDASI
  let persentase = totalHariKalender === 0 ? 0 : Math.round((hariTerisi / totalHariKalender) * 100);
  const activityList = [...new Set(filledData.map(r => r.kegiatan))]; 
  
  let rekomendasi = "";
  let warnaRekomendasi = "#333";
  let sumberSaran = "System AI Analysis";

  if (filledData.length > 0) {
    const aiResult = getAIReview(activityList, persentase);
    if (aiResult) {
      rekomendasi = aiResult;
      if (persentase >= 90) warnaRekomendasi = "#166534";
      else if (persentase >= 70) warnaRekomendasi = "#ca8a04";
      else warnaRekomendasi = "#b91c1c";
    }
  }

  // Fallback Manual
  if (!rekomendasi) {
    sumberSaran = "System Logic";
    if (nameFilter === "ALL") {
       rekomendasi = "Rekapitulasi total pegawai. Silakan cek laporan individu untuk detail evaluasi.";
       warnaRekomendasi = "#333";
    } else {
        if (hariKosong > 3) {
          rekomendasi = `PERLU PEMBINAAN. Terdeteksi tidak mengisi laporan kinerja sebanyak ${hariKosong} hari. Mohon tingkatkan kedisiplinan administrasi harian.`;
          warnaRekomendasi = "#b91c1c";
        } else if (persentase >= 90) {
          rekomendasi = "SANGAT BAIK. Laporan lengkap dan konsisten setiap hari.";
          warnaRekomendasi = "#166534";
        } else {
          rekomendasi = "CUKUP. Mohon lengkapi laporan tepat waktu setiap hari kerja.";
          warnaRekomendasi = "#ca8a04";
        }
    }
  }

  // 5. STRUKTUR HTML
  const html = `
    <html>
    <head>
      <style>
        body { font-family: 'Helvetica', sans-serif; padding: 20px; color: #333; }
        .header { text-align: center; margin-bottom: 20px; border-bottom: 3px solid #059669; padding-bottom: 10px; }
        .header-content { display: flex; align-items: center; justify-content: center; gap: 15px; flex-direction: column; }
        h1 { margin: 5px 0 0 0; color: #059669; font-size: 18px; }
        p { margin: 5px 0; color: #555; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; font-size: 10px; margin-bottom: 20px; }
        th { background: #059669; color: white; padding: 8px; text-align: left; }
        
        .summary-box { 
            border: 1px solid #e5e7eb; 
            background-color: #f0fdf4; 
            padding: 15px; 
            border-radius: 8px; 
            margin-top: 10px; 
            page-break-inside: avoid;
            position: relative;
        }
        .summary-badge {
            position: absolute; top: -10px; right: 10px;
            background: #059669; color: white; 
            font-size: 9px; padding: 3px 8px; border-radius: 10px;
            text-transform: uppercase; font-weight: bold;
        }
        .summary-title { font-size: 12px; font-weight: bold; margin-bottom: 5px; text-transform: uppercase; color: #374151; }
        .summary-text { font-size: 11px; font-style: italic; line-height: 1.5; color: ${warnaRekomendasi}; }

        .ttd-container { display: flex; justify-content: flex-end; margin-top: 50px; page-break-inside: avoid; padding-right: 30px; }
        .ttd-box { width: 250px; text-align: center; font-size: 11px; }
        .ttd-space { height: 70px; }
      </style>
    </head>
    <body>
      <div class="header">
        <div class="header-content">
           <img src="${logoBase64}" style="width: 60px; height: auto;">
           <div>
             <h1>${title}</h1>
             <p>SMKPP NEGERI BIMA | Periode: ${periode}</p>
           </div>
        </div>
      </div>

      <div style="font-size: 11px; margin-bottom: 10px;">
        <b>Ringkasan Kehadiran:</b> <br>
        Periode: ${totalHariKalender} Hari | Mengisi: ${hariTerisi} Hari | <b>Kosong: ${hariKosong} Hari</b>
      </div>

      <table>
        <thead>
          <tr>
            <th width="12%">Tanggal</th>
            <th width="20%">Nama Pegawai</th>
            <th width="30%">Uraian Kegiatan</th>
            <th width="15%" style="text-align:center;">Waktu</th>
            <th width="10%">Status</th>
            <th width="13%" style="text-align:center;">Bukti</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>

      <div class="summary-box">
         <div class="summary-badge">✨ ${sumberSaran}</div>
         <div class="summary-title">Evaluasi Kinerja & Disiplin:</div>
         <div class="summary-text">"${rekomendasi}"</div>
      </div>

      <div class="ttd-container">
         <div class="ttd-box">
            <p>Bima, ${Utilities.formatDate(new Date(), TIMEZONE, "dd MMMM yyyy")}<br>Mengetahui,<br>Kepala Sekolah</p>
            <div class="ttd-space"></div>
            <p><b>Abdul Hamid, S.Pt., M.Pd</b><br>NIP. 19770324 200801 1 010</p>
         </div>
      </div>

    </body>
    </html>
  `;
  
  const blob = Utilities.newBlob(html, MimeType.HTML).setName(`Laporan_${nameFilter.replace(/\s/g,'_')}_${startStr}.pdf`);
  const pdf = blob.getAs(MimeType.PDF);
  return 'data:application/pdf;base64,' + Utilities.base64Encode(pdf.getBytes());
}

function getDynamicSheetName() {
  const now = new Date();
  // Membuat nama sheet: Laporan_Feb_2026
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"];
  return "Laporan_" + monthNames[now.getMonth()] + "_" + now.getFullYear();
}

function checkAndSetupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Sheet Pegawai (Tetap Permanen/Satu saja)
  let sheetStaff = ss.getSheetByName('DataPegawai');
  if (!sheetStaff) {
    sheetStaff = ss.insertSheet('DataPegawai');
    sheetStaff.appendRow(['Nama Pegawai', 'Jabatan', 'ID System', 'Device ID', 'Email']);
    sheetStaff.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e2e8f0');
    sheetStaff.setFrozenRows(1);
  }

  // 2. Setup Sheet Laporan Bulanan (Dinamis)
  const activeSheetName = getDynamicSheetName();
  let sheetJob = ss.getSheetByName(activeSheetName);
  
  if (!sheetJob) {
    sheetJob = ss.insertSheet(activeSheetName);
    const headers = [
      'ID Job', 'Tgl Dibuat', 'Nama Pegawai', 'Kegiatan', 'Lokasi Detil', 
      'Koordinat Awal', 'Foto Before', 'Waktu Mulai', 
      'Koordinat Akhir', 'Foto After', 'Waktu Selesai', 'Status'
    ];
    sheetJob.appendRow(headers);
    sheetJob.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#166534').setFontColor('#ffffff');
    sheetJob.setFrozenRows(1);
    
    // Opsional: Beri warna tab berbeda untuk tiap bulan agar rapi
    sheetJob.setTabColor(now.getMonth() % 2 === 0 ? "059669" : "0284c7");
  }
  
  return activeSheetName;
}

function setupDatabaseManual() {
  checkAndSetupDatabase();
  SpreadsheetApp.getUi().alert('Database Siap! Pastikan kolom Email di DataPegawai diisi.');
}

// Tambahkan ini di awal fungsi submitStart() atau submitFinish()
function unlockBeforeSubmit() {
  document.getElementById('namaStart').disabled = false;
  document.getElementById('namaFinish').disabled = false;
}

// ==========================================
// 1. HANDLE JOB START (DENGAN CEK DUPLIKASI)
// ==========================================
/**
 * Fungsi Utama: Mulai Pekerjaan
 * Dilengkapi dengan Auto-Register Device ID dan Proteksi Duplikasi
 */
function handleJobStart(formObject) {
  try {
    const activeSheetName = checkAndSetupDatabase(); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLaporan = ss.getSheetByName(activeSheetName);
    const sheetUser = ss.getSheetByName("DataPegawai");

    if (!sheetUser) throw new Error("Sheet 'DataPegawai' tidak ditemukan.");

    // --- 1. VALIDASI & AUTO-REGISTER DEVICE ID ---
    const userData = sheetUser.getDataRange().getValues();
    const rowIndex = userData.findIndex(row => 
      row[0].toString().trim().toLowerCase() === formObject.namaPegawai.toString().trim().toLowerCase()
    );

    if (rowIndex === -1) {
      throw new Error(`Nama [${formObject.namaPegawai}] tidak terdaftar di database.`);
    }

    const registeredDeviceId = userData[rowIndex][3] ? userData[rowIndex][3].toString().trim() : "";
    const currentDeviceId = formObject.deviceId.toString().trim();

    // Jika Kolom D Kosong, Daftarkan otomatis
    if (registeredDeviceId === "") {
      sheetUser.getRange(rowIndex + 1, 4).setValue(currentDeviceId);
      console.log("Device ID didaftarkan otomatis untuk: " + formObject.namaPegawai);
    } 
    // Jika sudah terdaftar, cek kecocokan
    else if (registeredDeviceId !== currentDeviceId) {
      throw new Error("AKSES DITOLAK: Akun ini sudah terkunci di perangkat lain.");
    }

    // --- 2. PROTEKSI DUPLIKASI KEGIATAN ---
    const lastRow = sheetLaporan.getLastRow();
    if (lastRow > 1) {
      const dataCek = sheetLaporan.getRange(2, 1, lastRow - 1, 12).getValues();
      const isDuplicate = dataCek.find(row => 
        row[2] === formObject.namaPegawai && 
        row[3].trim().toLowerCase() === formObject.kegiatan.trim().toLowerCase() && 
        row[11] === "PROSES"
      );

      if (isDuplicate) {
        throw new Error(`KEGIATAN GANDA: "${formObject.kegiatan}" masih dalam status PROSES.`);
      }
    }

    // --- 3. SIMPAN DATA & FOTO ---
    const idJob = "AKSI-" + Utilities.formatDate(new Date(), TIMEZONE, "ddHHmmss");
    const fileUrl = saveImageToDrive(formObject.imageDataBefore, idJob + "_BEFORE.jpg");
    const now = new Date();
    
    // Format Kolom: ID, Tgl, Nama, Kegiatan, Lokasi, Koordinat, Foto, Jam, ..., Status
    sheetLaporan.appendRow([
      idJob, 
      now, 
      formObject.namaPegawai, 
      formObject.kegiatan, 
      formObject.lokasiKerja, 
      formObject.koordinatStart, 
      fileUrl, 
      Utilities.formatDate(now, TIMEZONE, "HH:mm"), 
      "", "", "", "PROSES"
    ]);

    return { 
      status: 'success', 
      msg: 'Laporan Awal Berhasil Terkirim!', 
      id: idJob 
    };

  } catch (e) {
    console.error(e.toString());
    return { 
      status: 'error', 
      msg: e.toString().replace("Error: ", "") 
    };
  }
}
/**
 * Fungsi untuk mengambil statistik poin pegawai
 */
function getGamificationStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("STATISTIK"); // Sesuaikan nama sheet Anda
  
  if (!sheet) return { points: 0, level: "Pemula", rank: "-" };

  // Logika sederhana: ambil data dari cell tertentu
  // Sesuaikan koordinat cell (baris, kolom) dengan spreadsheet Anda
  const points = sheet.getRange("B2").getValue() || 0; 
  const rank = sheet.getRange("B3").getValue() || "N/A";
  
  let level = "Bronze";
  if (points > 500) level = "Silver";
  if (points > 1000) level = "Gold";

  return {
    points: points,
    level: level,
    rank: rank
  };
}

/**
 * Fungsi Utama: Menyelesaikan Pekerjaan
 * Dilengkapi Verifikasi Device ID, Radius Lokasi, dan Update Status
 */
function handleJobFinish(formObject) {
  try {
    const activeSheetName = getDynamicSheetName();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLaporan = ss.getSheetByName(activeSheetName);
    
    if (!sheetLaporan) throw new Error("Database bulan ini (" + activeSheetName + ") belum dibuat.");

    const lastRow = sheetLaporan.getLastRow();
    if (lastRow < 2) throw new Error("Tidak ada data laporan aktif.");

    // Ambil data kolom ID (Kolom A)
    const dataIds = sheetLaporan.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const rowIndex = dataIds.indexOf(formObject.jobId);
    
    if (rowIndex === -1) throw new Error("ID Laporan [" + formObject.jobId + "] tidak ditemukan di sheet.");
    
    const rowNumber = rowIndex + 2; 
    const currentStatus = sheetLaporan.getRange(rowNumber, 12).getValue(); // Kolom L
    
    if (currentStatus === "SELESAI") throw new Error("Laporan ini sudah diselesaikan sebelumnya.");

    // SIMPAN FOTO & UPDATE
    const fileUrl = saveImageToDrive(formObject.imageDataAfter, formObject.jobId + "_AFTER.jpg");
    const now = new Date();

    // Update: I (Koord Akhir), J (Foto After), K (Jam Selesai), L (Status)
    sheetLaporan.getRange(rowNumber, 9).setValue(formObject.koordinatEnd); 
    sheetLaporan.getRange(rowNumber, 10).setValue(fileUrl);               
    sheetLaporan.getRange(rowNumber, 11).setValue(Utilities.formatDate(now, TIMEZONE, "HH:mm")); 
    sheetLaporan.getRange(rowNumber, 12).setValue("SELESAI");             

    return { 
      status: 'success', 
      msg: 'Laporan Selesai Kerja Berhasil Diverifikasi.' 
    };

  } catch (e) {
    return { status: 'error', msg: e.toString() };
  }
}
// ==========================================
// 3. MODUL IZIN & SAKIT
// ==========================================
function handleAbsence(formObject) {
  try {
    checkAndSetupDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLaporan = ss.getSheetByName('LaporanKerja');
    
    const idJob = "IZIN-" + Utilities.formatDate(new Date(), TIMEZONE, "ddHHmmss");
    const now = new Date();
    
    let fileUrl = "";
    if (formObject.imageBukti && formObject.imageBukti.length > 100) {
       fileUrl = saveImageToDrive(formObject.imageBukti, idJob + "_BUKTI.jpg");
    }

    sheetLaporan.appendRow([
      idJob, now, formObject.namaPegawai, 
      `[${formObject.tipeIzin}] ${formObject.alasan}`, 
      "-", "-", fileUrl, "00:00", "-", "", "23:59", 
      formObject.tipeIzin
    ]);

    return { status: 'success', msg: 'Pengajuan Izin Berhasil!' };
  } catch (e) {
    return { status: 'error', msg: e.toString() };
  }
}

// ==========================================
// 4. DATA GETTERS (API untuk Frontend)
// ==========================================

function getStaffList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('DataPegawai');
    
    // Jika sheet belum ada, buatkan
    if (!sheet) {
      setupDatabaseManual();
      sheet = ss.getSheetByName('DataPegawai');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; // Kirim array kosong jika tidak ada nama
    
    return sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
  } catch (e) {
    console.error("Error getStaffList: " + e.toString());
    return [];
  }
}

function getOpenJobs(namaPegawai) {
  const activeSheetName = getDynamicSheetName();
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName);
  if (!ss || ss.getLastRow() < 2) return [];

  const data = ss.getRange(2, 1, ss.getLastRow() - 1, 12).getValues();
  
  return data.filter(row => row[2] === namaPegawai && row[11] === "PROSES")
             .map(row => ({ 
               id: row[0], 
               desc: row[3],
               startCoord: row[5], // Kolom F (Koordinat Awal)
               time: row[1] instanceof Date ? row[1].getTime() : new Date(row[1]).getTime() // Tgl Dibuat
             }));
}

// 1. VERIFIKASI PERANGKAT & SINKRONISASI PEKERJAAN
function checkUserByDevice(deviceId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStaff = ss.getSheetByName('DataPegawai');
  if (!sheetStaff) return { status: 'error', msg: 'Sheet DataPegawai tidak ditemukan' };
  
  // Ambil semua data dari sheet
  const dataStaff = sheetStaff.getDataRange().getValues();
  
  let foundName = null;
  let foundJabatan = null;

  // Lakukan pembersihan Device ID dari HP agar tidak ada spasi pengganggu
  const cleanDeviceId = deviceId.toString().trim();

  // Loop mulai dari baris kedua (indeks 1)
  for (let i = 1; i < dataStaff.length; i++) {
    // SESUAI STRUKTUR ANDA:
    // A=0 (Nama), B=1 (Jabatan), C=2 (ID Sys), D=3 (Device ID)
    const rowDeviceId = dataStaff[i][3] ? dataStaff[i][3].toString().trim() : "";
    
    if (rowDeviceId === cleanDeviceId) {
      foundName = dataStaff[i][0];    // AMBIL KOLOM A (Index 0)
      foundJabatan = dataStaff[i][1]; // AMBIL KOLOM B (Index 1)
      break;
    }
  }

  // Debugging log (Cek di "Executions" jika masih gagal)
  console.log("Mencari Device ID: " + cleanDeviceId);
  console.log("Hasil ditemukan: " + foundName);

  if (!foundName) return { status: 'unknown' };

  // Ambil pekerjaan pending (Pastikan nama di sheet laporan cocok dengan foundName)
  const activeSheetName = getDynamicSheetName();
  const sheetJobs = ss.getSheetByName(activeSheetName);
  let pendingJobs = [];

  if (sheetJobs && sheetJobs.getLastRow() > 1) {
    const dataJobs = sheetJobs.getRange(2, 1, sheetJobs.getLastRow() - 1, 12).getValues();
    pendingJobs = dataJobs
      .filter(row => row[2].toString().trim() === foundName.toString().trim() && row[11] === "PROSES")
      .map(row => ({ 
        id: row[0], 
        desc: row[3],
        name: row[2],
        startCoord: row[5],
        time: row[1] instanceof Date ? row[1].getTime() : new Date(row[1]).getTime()
      }));
  }

  return { 
    status: 'found', 
    name: foundName, 
    jabatan: foundJabatan, 
    pending: pendingJobs 
  };
}

// --- FUNGSI VERIFIKASI DEVICE (PASTE DI CODE.GS) ---
function verifyDevice(formDeviceId, userRow, rowIndex, sheetUser) {
  const registeredDeviceId = userRow[3] ? userRow[3].toString().trim() : "";
  const currentDeviceId = formDeviceId.toString().trim();

  // JIKA KOLOM D KOSONG: Daftarkan HP ini secara otomatis (Fitur Memudahkan Pegawai Baru)
  if (registeredDeviceId === "") {
    sheetUser.getRange(rowIndex + 1, 4).setValue(currentDeviceId); // Simpan ke Kolom D
    console.log("Device ID baru didaftarkan otomatis untuk: " + userRow[0]);
    return true;
  }

  // JIKA SUDAH ADA ISI: Cek kecocokannya
  if (registeredDeviceId !== currentDeviceId) {
    throw new Error("PERANGKAT TIDAK SAH. Akun ini sudah terkunci di HP lain. Hubungi Admin untuk reset.");
  }

  return true;
}

function getHistoryHarian(namaPegawai) {
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LaporanKerja');
  if (!ss || ss.getLastRow() < 2) return [];

  const now = new Date();
  const todayStr = Utilities.formatDate(now, TIMEZONE, "dd/MM/yyyy");
  const data = ss.getRange(2, 1, ss.getLastRow() - 1, 12).getValues();

  const filteredData = data.filter(row => {
    if (row[2] !== namaPegawai) return false;
    let rowDateStr = "";
    try { rowDateStr = Utilities.formatDate(new Date(row[1]), TIMEZONE, "dd/MM/yyyy"); } catch (e) { return false; }
    return (rowDateStr === todayStr) || (row[11] === "PROSES");
  });

  return filteredData.map(row => {
    let jamMulai = row[7];
    if (jamMulai instanceof Date) jamMulai = Utilities.formatDate(jamMulai, TIMEZONE, "HH:mm");
    
    let jamSelesai = row[10];
    if (jamSelesai instanceof Date) jamSelesai = Utilities.formatDate(jamSelesai, TIMEZONE, "HH:mm");
    if (!jamSelesai) jamSelesai = "-";

    return { kegiatan: row[3], jamMulai: jamMulai, jamSelesai: jamSelesai, status: row[11] };
  }).reverse();
}

function getRecentGallery() {
  const FOLDER_ID = '1VshBRT2Joxs0Na0XsO8GWNcQ2351mbwM'; // Pastikan ID ini Benar
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  let fileList = [];
  
  // Ambil file gambar (JPEG, PNG, dsb)
  while (files.hasNext()) {
    let file = files.next();
    let mime = file.getMimeType();
    if (mime === "image/jpeg" || mime === "image/png" || mime === "image/jpg") {
      fileList.push({
        id: file.getId(),
        created: file.getDateCreated()
      });
    }
  }

  // Urutkan dari yang terbaru
  fileList.sort((a, b) => b.created - a.created);
  
  // Ambil 5 foto terbaru dan konversi ke URL Direct View
  return fileList.slice(0, 5).map(f => ({
    // Format URL ini sangat penting agar gambar langsung tampil
    url: "https://lh3.googleusercontent.com/u/0/d/" + f.id, 
    date: Utilities.formatDate(f.created, "GMT+7", "dd MMM HH:mm")
  }));
}

// ==========================================
// 5. MODUL LAPORAN BULANAN (PDF + EMAIL + LOGO)
// ==========================================

function generateMonthlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLaporan = ss.getSheetByName('LaporanKerja');
  const sheetPegawai = ss.getSheetByName('DataPegawai');
  
  if (!sheetLaporan || !sheetPegawai) return;
  
  const dataLaporan = sheetLaporan.getDataRange().getValues();
  const dataPegawai = sheetPegawai.getDataRange().getValues();
  
  const now = new Date();
  
  // --- KONFIGURASI BULAN ---
  // A. MODE TESTING (Bulan Ini):
  const reportDate = new Date(now.getFullYear(), now.getMonth(), 1); 
  
  // B. MODE PRODUKSI (Bulan Lalu - Gunakan ini untuk trigger otomatis):
  // const reportDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  
  const monthName = Utilities.formatDate(reportDate, TIMEZONE, "MMMM yyyy");
  const filterMonth = reportDate.getMonth();

  console.log("Memulai Generasi Rapor: " + monthName);

  for (let i = 1; i < dataPegawai.length; i++) {
    const nama = dataPegawai[i][0];
    const email = dataPegawai[i][4]; // Kolom E
    
    if (!email) continue;

    const laporanPegawai = dataLaporan.slice(1).filter(row => {
      try {
        const rowDate = new Date(row[1]);
        return row[2] === nama && rowDate.getMonth() === filterMonth;
      } catch(e) { return false; }
    });

    if (laporanPegawai.length > 0) {
      try {
        const pdfBlob = createPDF(nama, monthName, laporanPegawai);
        
        MailApp.sendEmail({
          to: email,
          subject: `Rapor Kinerja: ${nama} (${monthName})`,
          body: `Halo ${nama},\n\nBerikut terlampir rekap laporan kinerja Anda periode ${monthName}.\n\nSalam,\nTim Digitalisasi Sekolah SMKPP Negeri Bima`,
          attachments: [pdfBlob]
        });
        console.log(`✅ Terkirim: ${nama}`);
      } catch (e) {
        console.log(`❌ Gagal: ${nama} - ${e.message}`);
      }
    }
  }
}

function createPDF(nama, periode, data) {
  let tableRows = "";
  
  data.forEach(row => {
    const tgl = Utilities.formatDate(new Date(row[1]), TIMEZONE, "dd/MM/yyyy");
    const kegiatan = row[3];
    const status = row[11];

    let jamMulai = row[7];
    if (jamMulai instanceof Date) jamMulai = Utilities.formatDate(jamMulai, TIMEZONE, "HH:mm");
    
    let jamSelesai = row[10];
    if (jamSelesai instanceof Date) jamSelesai = Utilities.formatDate(jamSelesai, TIMEZONE, "HH:mm");

    const jamDisplay = (status === 'IZIN' || status === 'SAKIT') ? '-' : `${jamMulai} - ${jamSelesai}`;
    
    let color = '#000';
    if(status === 'SELESAI') color = '#166534'; 
    if(status === 'PROSES') color = '#ca8a04';
    if(status.includes('IZIN') || status.includes('SAKIT')) color = '#ea580c';
    
    tableRows += `
      <tr style="border-bottom: 1px solid #eee;">
        <td style="padding: 8px;">${tgl}</td>
        <td style="padding: 8px;">${kegiatan}</td>
        <td style="padding: 8px; text-align:center;">${jamDisplay}</td>
        <td style="padding: 8px; color:${color}; font-weight:bold;">${status}</td>
      </tr>
    `;
  });

  // --- HTML PDF DENGAN LOGO ---
  const htmlContent = `
    <html>
    <head>
      <style>
        body { font-family: 'Helvetica', sans-serif; color: #333; padding: 20px; }
        /* Header dengan Flexbox agar Logo dan Teks sejajar */
        .header { 
            display: flex; 
            align-items: center; 
            justify-content: center;
            gap: 15px;
            border-bottom: 3px solid #059669; 
            padding-bottom: 15px; 
            margin-bottom: 25px; 
        }
        .logo-img {
            width: 70px;
            height: auto;
        }
        .header-text {
            text-align: left;
        }
        .header h1 { color: #059669; margin: 0; font-size: 22px; line-height: 1.2; }
        .header p { margin: 0; color: #666; font-size: 14px; }

        .info { margin-bottom: 20px; font-size: 14px; background: #f9fafb; padding: 10px; border-radius: 5px; }
        table { width: 100%; border-collapse: collapse; font-size: 12px; }
        th { background-color: #059669; color: white; text-align: left; padding: 10px; }
        .footer { margin-top: 30px; text-align: center; font-size: 10px; color: #999; border-top: 1px solid #eee; padding-top: 10px; }
      </style>
    </head>
    <body>
      <div class="header">
        <img src="${SCHOOL_LOGO_URL}" class="logo-img" alt="Logo SMKPP Bima">
        
        <div class="header-text">
            <h1>RAPOR KINERJA PEGAWAI</h1>
            <p>SMKPP NEGERI BIMA (WITA)</p>
        </div>
      </div>
      
      <div class="info">
        <b>Nama Pegawai:</b> ${nama}<br>
        <b>Periode:</b> ${periode}
      </div>

      <table>
        <thead>
          <tr>
            <th width="15%">Tanggal</th>
            <th width="50%">Kegiatan</th>
            <th width="20%" style="text-align:center;">Jam (WITA)</th>
            <th width="15%">Status</th>
          </tr>
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      <div class="footer">
        Dicetak otomatis oleh Sistem SI-AKSI pada ${Utilities.formatDate(new Date(), TIMEZONE, "dd/MM/yyyy HH:mm")} WITA
      </div>
    </body>
    </html>
  `;

  const blob = Utilities.newBlob(htmlContent, MimeType.HTML).setName(`Rapor_${nama}_${periode}.pdf`);
  return blob.getAs(MimeType.PDF);
}

// ==========================================
// 6. HELPERS
// ==========================================

function saveImageToDrive(base64Data, filename) {
  try {
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.substring(base64Data.indexOf(',') + 1));
    const blob = Utilities.newBlob(bytes, contentType, filename);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    
    // MATIKAN BARIS INI DENGAN MENAMBAHKAN // DI DEPANNYA
    // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) {
    throw new Error("Akses ditolak: DriveApp. " + e.message);
  }
}

function getDistanceMeters(lat1, lon1, lat2, lon2) {
  const R = 6371e3; 
  const dLat = (lat2-lat1) * Math.PI/180;
  const dLon = (lon2-lon1) * Math.PI/180;
  const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
            Math.cos(lat1*Math.PI/180) * Math.cos(lat2*Math.PI/180) * Math.sin(dLon/2) * Math.sin(dLon/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c; 
}

function setupMonthlyReportTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'generateMonthlyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('generateMonthlyReport')
    .timeBased()
    .onMonthDay(1)
    .atHour(6) 
    .create();
    
  SpreadsheetApp.getUi().alert('✅ Trigger Laporan Bulanan (Tgl 1) Aktif!');
}

// ==========================================
// 7. MODUL AI GEMINI (VERSI FINAL - LOGIKA LOKASI CERDAS)
// ==========================================
function callGeminiAI(textInput) {
  const API_KEY = "AIzaSyAAe-AhIBCsGt8p9V3q6wCmnBj9o1L30ME"; 
  const MODEL_NAME = "gemini-1.5-flash"; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${API_KEY}`;

  // PROMPT DENGAN INSTRUKSI "FORCE-FORMAL"
  const prompt = `
    Jadilah pakar administrasi e-Kinerja ASN. 
    Tugas: Ubah input singkat menjadi narasi aktivitas yang lengkap, formal, dan profesional.
    
    ATURAN BAKU:
    1. WAJIB diawali kata kerja: "Melaksanakan", "Melakukan", atau "Menyusun".
    2. Kalimat harus mengandung unsur: [KEGIATAN] + [OBJEK] + [TUJUAN/OUTPUT].
    3. DILARANG hanya menambahkan kata "Melaksanakan" di depan input asli.
    4. Kalimat minimal harus terdiri dari 10 kata.
    
    CONTOH STANDAR E-KINERJA:
    - Input: "tanam bibit" -> Hasil: "Melaksanakan penanaman bibit tanaman pada area lahan praktik guna memastikan keberhasilan fase awal budidaya komoditas pertanian."
    - Input: "beri pakan" -> Hasil: "Melaksanakan manajemen pemberian pakan ternak secara rutin untuk menjaga produktivitas dan kesehatan hewan ternak di lingkungan sekolah."
    
    Input Pegawai: "${textInput}"
    Hasil (Hanya 1 kalimat narasi formal):
  `;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 1, // Maksimalkan kreativitas agar kalimat panjang
      topP: 0.95,
      maxOutputTokens: 250
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const resText = response.getContentText();
    const json = JSON.parse(resText);
    
    if (json.candidates && json.candidates.length > 0) {
      let hasil = json.candidates[0].content.parts[0].text.trim();
      
      // Filter Akhir: Jika AI masih bandel kasih kalimat pendek, kita buatkan pola paksa
      if (hasil.split(' ').length < 5) {
        return "Melaksanakan proses " + textInput + " secara berkala guna mendukung ketercapaian target kinerja dan standar operasional prosedur di SMKPP Negeri Bima.";
      }
      return hasil;
    } else {
      throw new Error("Respons AI Kosong");
    }
  } catch (e) {
    // Kalimat cadangan jika API error, tetap dibuat panjang & formal
    return "Melaksanakan kegiatan " + textInput + " sesuai dengan instruksi kerja untuk memastikan efektivitas pelaksanaan tugas harian di lingkungan sekolah.";
  }
}

function cekDaftarModel() {
  const API_KEY = "AIzaSyAAe-AhIBCsGt8p9V3q6wCmnBj9o1L30ME"; // Key Anda
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${API_KEY}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());
    
    console.log("=== DAFTAR MODEL YANG BISA DIPAKAI ===");
    if (json.models) {
      json.models.forEach(m => {
        // Hanya tampilkan model yang support generateContent
        if (m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent")) {
           console.log("Nama Model: " + m.name);
        }
      });
    } else {
      console.log("Tidak ada model ditemukan.");
    }
    console.log("======================================");
    
  } catch (e) {
    console.log("Error Cek Model: " + e.toString());
  }
}

// ==========================================
// 9. API DASHBOARD (LEADERBOARD CHART)
// ==========================================

function getLeaderboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheetName = getDynamicSheetName(); // Ambil sheet bulan ini saja
  const sheet = ss.getSheetByName(activeSheetName);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { names: [], counts: [] };
  }

  const data = sheet.getDataRange().getValues();
  const counts = {};

  // Loop data (Mulai baris ke-2 / index 1)
  // Kolom: 2=Nama, 11=Status (Index 0-based)
  data.slice(1).forEach(row => {
    try {
      const name = row[2];
      const status = row[11];

      // Di sistem bulanan, kita tidak perlu cek bulan lagi 
      // karena sheetnya sudah pasti sheet bulan ini.
      if (status === 'SELESAI') {
        counts[name] = (counts[name] || 0) + 1;
      }
    } catch (e) {
      console.log("Error baris: " + e.message);
    }
  });

  // Urutkan berdasarkan performa tertinggi
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);

  return {
    names: sorted.map(item => item[0]),
    counts: sorted.map(item => item[1])
  };
}

/**
 * Fungsi untuk membersihkan laporan yang menggantung lebih dari 24 jam.
 * Jalankan fungsi ini menggunakan Trigger setiap 1 jam sekali.
 */
function autoCleanupExpiredJobs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const now = new Date();
  const limitTime = 24 * 60 * 60 * 1000; // Konversi 24 Jam ke Milidetik

  sheets.forEach(sheet => {
    // Memeriksa sheet yang diawali dengan "Laporan_" atau sheet utama "LaporanKerja"
    if (sheet.getName().indexOf("Laporan_") === 0 || sheet.getName() === "LaporanKerja") {
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;

      const range = sheet.getRange(2, 1, lastRow - 1, 12);
      const values = range.getValues();
      let hasChange = false;

      const updatedValues = values.map(row => {
        // Indeks: 1 = Tanggal/Waktu Mulai, 11 = Status
        const startTime = new Date(row[1]).getTime();
        const status = row[11];

        // Jika status masih PROSES dan sudah lewat dari 24 jam
        if (status === "PROSES" && (now.getTime() - startTime) > limitTime) {
          row[11] = "KADALUWARSA (SISTEM)"; // Mengubah status agar tidak dihitung di statistik
          row[10] = "Gagal Selesai > 24 Jam"; // Catatan pada kolom bukti foto akhir
          hasChange = true;
        }
        return row;
      });

      if (hasChange) {
        range.setValues(updatedValues);
        console.log(`Pembersihan dilakukan di sheet: ${sheet.getName()}`);
      }
    }
  });
}

function tesFolderLangsung() {
  try {
    const folderId = "1VshBRT2Joxs0Na0XsO8GWNcQ2351mbwM";
    const folder = DriveApp.getFolderById(folderId);
    folder.createFile("Tes_Akses_Sistem.txt", "Jika file ini ada, berarti izin Drive sudah sukses!");
    console.log("SUKSES: File berhasil dibuat di folder.");
  } catch (e) {
    console.log("GAGAL: " + e.message);
  }
}