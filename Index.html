/*******************************************************
 * SISTEM ABSENSI TERPADU
 * MTs PSM LOCERET
 * PROJECT BARU
 *
 * TAHAP 1
 * DATABASE + FUNGSI DASAR
 *******************************************************/

const NAMA_DATABASE = 'DATABASE ABSENSI TERPADU MTs PSM LOCERET';

/**
 * ID Spreadsheet database.
 *
 * Jika project Apps Script dibuat langsung dari
 * Google Spreadsheet database, otomatis menggunakan
 * Spreadsheet aktif.
 *
 * Jika project berdiri sendiri, isi ID spreadsheet
 * pada bagian ini.
 */
const DATABASE_ID = '';

/**
 * Nama-nama sheet database
 */
const SHEET = {
  PENGATURAN: 'PENGATURAN',
  SISWA: 'SISWA',
  GURU: 'GURU',
  KELAS: 'KELAS',
  MAPEL: 'MAPEL',
  JADWAL: 'JADWAL',
  ABSENSI_SISWA: 'ABSENSI_SISWA',
  ABSENSI_GURU: 'ABSENSI_GURU',
  JAM_SEKOLAH: 'JAM_SEKOLAH',
  HARI_LIBUR: 'HARI_LIBUR',
  KALENDER_AKADEMIK: 'KALENDER_AKADEMIK',
  USER: 'USER'
};


/*******************************************************
 * HEADER DATABASE
 *******************************************************/

const HEADER = {};

/* PENGATURAN */
HEADER.PENGATURAN = [
  'KUNCI',
  'NILAI',
  'KETERANGAN'
];

/* SISWA */
HEADER.SISWA = [
  'NIS',
  'NAMA',
  'KELAS',
  'JK',
  'TEMPAT_LAHIR',
  'TANGGAL_LAHIR',
  'ALAMAT',
  'NAMA_ORTU',
  'NO_WA_ORTU',
  'STATUS',
  'KETERANGAN'
];

/* GURU */
HEADER.GURU = [
  'ID_GURU',
  'NAMA',
  'NIP',
  'NUPTK',
  'JK',
  'NO_HP',
  'EMAIL',
  'MAPEL',
  'STATUS',
  'KETERANGAN'
];

/* KELAS */
HEADER.KELAS = [
  'ID_KELAS',
  'NAMA_KELAS',
  'TINGKAT',
  'WALI_KELAS',
  'STATUS'
];

/* MAPEL */
HEADER.MAPEL = [
  'ID_MAPEL',
  'NAMA_MAPEL',
  'KETERANGAN',
  'STATUS'
];

/**
 * JADWAL
 *
 * Nomor JAM menjadi dasar timer.
 * Bukan nama mata pelajaran.
 */
HEADER.JADWAL = [
  'ID_JADWAL',
  'HARI',
  'JAM_KE',
  'JAM_MULAI',
  'JAM_SELESAI',
  'KELAS',
  'MAPEL',
  'ID_GURU',
  'NAMA_GURU',
  'STATUS'
];

/* ABSENSI SISWA */
HEADER.ABSENSI_SISWA = [
  'TIMESTAMP',
  'TANGGAL',
  'JAM',
  'NIS',
  'NAMA',
  'KELAS',
  'JK',
  'KEGIATAN',
  'JAM_KE',
  'STATUS',
  'KETERANGAN'
];

/* ABSENSI GURU */
HEADER.ABSENSI_GURU = [
  'TIMESTAMP',
  'TANGGAL',
  'JAM',
  'ID_GURU',
  'NAMA_GURU',
  'JAM_KE',
  'KELAS',
  'MAPEL',
  'STATUS',
  'KETERANGAN'
];

/**
 * JAM_SEKOLAH
 *
 * Digunakan oleh mesin timer sekolah.
 */
HEADER.JAM_SEKOLAH = [
  'ID',
  'HARI',
  'URUTAN',
  'KEGIATAN',
  'JAM_KE',
  'JAM_MULAI',
  'JAM_SELESAI',
  'DURASI_MENIT',
  'NOTIFIKASI',
  'STATUS'
];

/* HARI LIBUR */
HEADER.HARI_LIBUR = [
  'TANGGAL',
  'NAMA_LIBUR',
  'JENIS',
  'KETERANGAN',
  'STATUS'
];

/* KALENDER AKADEMIK */
HEADER.KALENDER_AKADEMIK = [
  'TANGGAL_MULAI',
  'TANGGAL_SELESAI',
  'KEGIATAN',
  'KETERANGAN',
  'STATUS'
];

/* USER */
HEADER.USER = [
  'ID_USER',
  'USERNAME',
  'PASSWORD',
  'NAMA',
  'ROLE',
  'ID_GURU',
  'STATUS'
];


/*******************************************************
 * DATABASE
 *******************************************************/

/**
 * Mengambil spreadsheet database.
 */
function getDatabase_() {

  if (DATABASE_ID && DATABASE_ID.trim() !== '') {
    return SpreadsheetApp.openById(DATABASE_ID);
  }

  const aktif = SpreadsheetApp.getActiveSpreadsheet();

  if (!aktif) {
    throw new Error(
      'Database belum terhubung. Isi DATABASE_ID pada Code.gs.'
    );
  }

  return aktif;
}


/**
 * Membuat / mengambil sheet.
 */
function getSheet_(namaSheet) {

  const ss = getDatabase_();

  let sheet = ss.getSheetByName(namaSheet);

  if (!sheet) {
    sheet = ss.insertSheet(namaSheet);
  }

  return sheet;
}


/**
 * Membuat header jika belum ada.
 */
function siapkanHeader_(namaSheet, header) {

  const sheet = getSheet_(namaSheet);

  if (sheet.getLastRow() === 0) {

    sheet
      .getRange(1, 1, 1, header.length)
      .setValues([header]);

    formatHeader_(sheet, header.length);

  } else {

    const barisPertama = sheet
      .getRange(1, 1, 1, header.length)
      .getValues()[0];

    let kosong = true;

    for (let i = 0; i < barisPertama.length; i++) {
      if (String(barisPertama[i]).trim() !== '') {
        kosong = false;
        break;
      }
    }

    if (kosong) {
      sheet
        .getRange(1, 1, 1, header.length)
        .setValues([header]);

      formatHeader_(sheet, header.length);
    }
  }

  sheet.setFrozenRows(1);

  return sheet;
}


/**
 * Format header.
 */
function formatHeader_(sheet, jumlahKolom) {

  sheet
    .getRange(1, 1, 1, jumlahKolom)
    .setFontWeight('bold');

  sheet
    .getRange(1, 1, 1, jumlahKolom)
    .setHorizontalAlignment('center');

  sheet.autoResizeColumns(1, jumlahKolom);
}


/*******************************************************
 * SETUP DATABASE
 *******************************************************/

/**
 * Jalankan fungsi ini SATU KALI
 * untuk membuat struktur database.
 */
function setupDatabase() {

  const ss = getDatabase_();

  const daftarSheet = [
    SHEET.PENGATURAN,
    SHEET.SISWA,
    SHEET.GURU,
    SHEET.KELAS,
    SHEET.MAPEL,
    SHEET.JADWAL,
    SHEET.ABSENSI_SISWA,
    SHEET.ABSENSI_GURU,
    SHEET.JAM_SEKOLAH,
    SHEET.HARI_LIBUR,
    SHEET.KALENDER_AKADEMIK,
    SHEET.USER
  ];

  daftarSheet.forEach(function(nama) {

    if (!ss.getSheetByName(nama)) {
      ss.insertSheet(nama);
    }

  });


  siapkanHeader_(
    SHEET.PENGATURAN,
    HEADER.PENGATURAN
  );

  siapkanHeader_(
    SHEET.SISWA,
    HEADER.SISWA
  );

  siapkanHeader_(
    SHEET.GURU,
    HEADER.GURU
  );

  siapkanHeader_(
    SHEET.KELAS,
    HEADER.KELAS
  );

  siapkanHeader_(
    SHEET.MAPEL,
    HEADER.MAPEL
  );

  siapkanHeader_(
    SHEET.JADWAL,
    HEADER.JADWAL
  );

  siapkanHeader_(
    SHEET.ABSENSI_SISWA,
    HEADER.ABSENSI_SISWA
  );

  siapkanHeader_(
    SHEET.ABSENSI_GURU,
    HEADER.ABSENSI_GURU
  );

  siapkanHeader_(
    SHEET.JAM_SEKOLAH,
    HEADER.JAM_SEKOLAH
  );

  siapkanHeader_(
    SHEET.HARI_LIBUR,
    HEADER.HARI_LIBUR
  );

  siapkanHeader_(
    SHEET.KALENDER_AKADEMIK,
    HEADER.KALENDER_AKADEMIK
  );

  siapkanHeader_(
    SHEET.USER,
    HEADER.USER
  );


  isiPengaturanAwal_();
  isiJamSekolahAwal_();


  return {
    success: true,
    database: ss.getName(),
    sheets: daftarSheet
  };
}


/*******************************************************
 * PENGATURAN AWAL
 *******************************************************/

function isiPengaturanAwal_() {

  const sheet = getSheet_(SHEET.PENGATURAN);

  if (sheet.getLastRow() > 1) {
    return;
  }

  const data = [

    [
      'NAMA_SEKOLAH',
      'MTs PSM LOCERET',
      'Nama madrasah'
    ],

    [
      'JUDUL_SISTEM',
      'SISTEM ABSENSI TERPADU GURU & SISWA',
      'Judul aplikasi'
    ],

    [
      'JAM_MASUK',
      '07:00',
      'Jam mulai kegiatan sekolah'
    ],

    [
      'DURASI_JAM_PELAJARAN',
      '45',
      'Durasi satu jam pelajaran'
    ],

    [
      'JUMLAH_JAM_SENIN_KAMIS',
      '12',
      'Jumlah jam termasuk Madin'
    ],

    [
      'JUMLAH_JAM_JUMAT',
      '6',
      'Jumlah jam Jumat'
    ],

    [
      'MULAI_EKSTRA_SABTU',
      '07:00',
      'Kegiatan ekstra Sabtu'
    ],

    [
      'SELESAI_EKSTRA_SABTU',
      '08:30',
      'Akhir kegiatan ekstra Sabtu'
    ],

    [
      'DURASI_DHUHUR',
      '15',
      'Durasi waktu Dhuhur sebelum Madin'
    ]

  ];

  sheet
    .getRange(
      2,
      1,
      data.length,
      3
    )
    .setValues(data);
}


/*******************************************************
 * JAM SEKOLAH AWAL
 *******************************************************/

/**
 * Membuat template aturan waktu.
 *
 * Waktu final akan kita isi setelah
 * seluruh jam sekolah dikunci.
 */
function isiJamSekolahAwal_() {

  const sheet = getSheet_(SHEET.JAM_SEKOLAH);

  if (sheet.getLastRow() > 1) {
    return;
  }

  const data = [

    [
      'SJ-001',
      'SENIN-KAMIS',
      1,
      'SHOLAT DHUHA',
      1,
      '',
      '',
      '',
      'YA',
      'AKTIF'
    ],

    [
      'SJ-002',
      'SENIN-KAMIS',
      2,
      'PELAJARAN',
      2,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-003',
      'SENIN-KAMIS',
      3,
      'PELAJARAN',
      3,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-004',
      'SENIN-KAMIS',
      4,
      'PELAJARAN',
      4,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-005',
      'SENIN-KAMIS',
      5,
      'PELAJARAN',
      5,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-006',
      'SENIN-KAMIS',
      6,
      'PELAJARAN',
      6,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-007',
      'SENIN-KAMIS',
      7,
      'PELAJARAN',
      7,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-008',
      'SENIN-KAMIS',
      8,
      'PELAJARAN',
      8,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-009',
      'SENIN-KAMIS',
      9,
      'PELAJARAN',
      9,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-010',
      'SENIN-KAMIS',
      10,
      'PELAJARAN',
      10,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-011',
      'SENIN-KAMIS',
      11,
      'PELAJARAN',
      11,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-012',
      'SENIN-KAMIS',
      12,
      'MADIN',
      12,
      '',
      '',
      '',
      'YA',
      'AKTIF'
    ],

    [
      'SJ-F01',
      'JUMAT',
      1,
      'PELAJARAN',
      1,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-F02',
      'JUMAT',
      2,
      'PELAJARAN',
      2,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-F03',
      'JUMAT',
      3,
      'PELAJARAN',
      3,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-F04',
      'JUMAT',
      4,
      'PELAJARAN',
      4,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-F05',
      'JUMAT',
      5,
      'PELAJARAN',
      5,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-F06',
      'JUMAT',
      6,
      'PELAJARAN',
      6,
      '',
      '',
      45,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-SABTU-01',
      'SABTU',
      0,
      'KEGIATAN EKSTRA',
      0,
      '07:00',
      '08:30',
      90,
      'YA',
      'AKTIF'
    ],

    [
      'SJ-SABTU-02',
      'SABTU',
      1,
      'PELAJARAN',
      1,
      '08:30',
      '',
      45,
      'YA',
      'AKTIF'
    ]

  ];

  sheet
    .getRange(
      2,
      1,
      data.length,
      data[0].length
    )
    .setValues(data);
}


/*******************************************************
 * PING / CEK DATABASE
 *******************************************************/

/**
 * Jalankan untuk memastikan database terhubung.
 */
function ping() {

  const ss = getDatabase_();

  return {
    success: true,
    message: 'DATABASE TERHUBUNG',
    database: ss.getName(),
    spreadsheetId: ss.getId(),
    url: ss.getUrl()
  };
}


/**
 * Mengecek seluruh sheet.
 */
function cekDatabase() {

  const ss = getDatabase_();

  const hasil = {};

  Object.keys(SHEET).forEach(function(kunci) {

    const namaSheet = SHEET[kunci];

    const sheet = ss.getSheetByName(namaSheet);

    hasil[namaSheet] = {
      ada: !!sheet,
      jumlahBaris: sheet
        ? sheet.getLastRow()
        : 0,
      jumlahKolom: sheet
        ? sheet.getLastColumn()
        : 0
    };

  });

  return {
    success: true,
    database: ss.getName(),
    sheets: hasil
  };
}


/*******************************************************
 * WEB APP DASAR
 *******************************************************/
  
function doGet(e) {
  // Endpoint API memakai POST dari scanner GitHub. GET sengaja hanya
  // menampilkan status backend supaya tidak ada absensi yang tersimpan
  // hanya karena URL dibuka atau di-refresh.
  const html = `<!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Backend Absensi MTs PSM Loceret</title></head><body style="font-family:Arial,sans-serif;padding:24px;line-height:1.6"><h2>Backend Absensi Terpadu MTs PSM Loceret aktif</h2><p>Alamat ini adalah endpoint server. Buka scanner GitHub proyek baru untuk melakukan pemindaian QR.</p><p>Absensi tidak dicatat melalui halaman ini.</p></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('Backend Absensi MTs PSM Loceret')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Pintu masuk scanner GitHub baru.
 * POST dilakukan melalui form + iframe agar tidak bergantung pada CORS.
 * Hasil dikirim kembali ke halaman scanner melalui postMessage.
 */
function doPost(e) {
  let hasil;
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const qr = String(p.qr || p.kodeQR || p.nis || '').trim();
    const role = String(p.role || '').trim().toLowerCase();
    const lat = p.latitude;
    const lon = p.longitude;

    // Fail-closed: jangan izinkan endpoint publik menulis absensi sebelum
    // sekolah memasang KODE_SCANNER pada sheet PENGATURAN.
    const pinTersimpan = String(bacaPengaturanTerpadu_('KODE_SCANNER', '') || '').trim();
    const pinMasuk = String(p.pin || '').trim();
    if (!pinTersimpan) {
      hasil = { success: false, message: 'Scanner belum diaktifkan. Admin perlu mengisi KODE_SCANNER pada sheet PENGATURAN.' };
    } else if (!pinMasuk || pinMasuk !== pinTersimpan) {
      hasil = { success: false, message: 'Kode akses scanner salah atau belum diisi.' };
    } else if (!qr) {
      hasil = { success: false, message: 'Data QR kosong. Silakan pindai ulang.' };
    } else {
      const siswaDariQR = /^NIS\s*[:=]/i.test(qr) || /^\d{3,30}$/.test(qr);
      const prosesSebagaiSiswa = role === 'siswa' || (!role && siswaDariQR);
      const prosesSebagaiGuru = role === 'guru' || (!role && !siswaDariQR);

      if (prosesSebagaiSiswa) {
        hasil = prosesScanSiswaTerpadu(qr, lat, lon);
      } else if (prosesSebagaiGuru) {
        hasil = prosesScanGuruTerpadu(qr);
      } else {
        hasil = { success: false, message: 'Jenis QR tidak dikenali.' };
      }
    }
  } catch (err) {
    hasil = { success: false, message: 'Server gagal memproses scan: ' + (err && err.message ? err.message : err) };
  }
  return responsScannerPostMessage_(hasil);
}

/** Mengembalikan hasil POST ke scanner GitHub dalam iframe tersembunyi. */
function responsScannerPostMessage_(hasil) {
  const safeJson = JSON.stringify(hasil || { success: false, message: 'Server tidak mengembalikan hasil.' })
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>(function(){var result=' + safeJson + ';' +
    'try{window.parent.postMessage({source:"MTSPM_ABSENSI_RESULT",result:result},"*");}catch(e){}' +
    'document.body.textContent="Respons absensi telah diproses. Silakan kembali ke scanner.";' +
    '})();</scr' + 'ipt></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle('Hasil Absensi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/****************************************************
 * ABSENSI TERPADU MTs PSM LOCERET
 * MODUL : JADWAL SEKOLAH
 ****************************************************/

const NAMA_SHEET_JADWAL = '06_Jadwal';
const NAMA_SHEET_ABSENSI = '05_Absensi';
const NAMA_SHEET_SISWA = '02_Siswa';


/**
 * Mengambil jadwal sekolah berdasarkan hari
 */
function getJadwalSekolah_(tanggal) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NAMA_SHEET_JADWAL);

  if (!sheet) {
    throw new Error('Sheet 06_Jadwal tidak ditemukan.');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return null;
  }

  const namaHari = [
    'Minggu',
    'Senin',
    'Selasa',
    'Rabu',
    'Kamis',
    'Jumat',
    'Sabtu'
  ];

  const hari = namaHari[new Date(tanggal).getDay()];

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const id = row[0];
    const hariJadwal = String(row[1]).trim();
    const jamMulai = formatJam_(row[2]);
    const jamSelesai = formatJam_(row[3]);
    const jenis = String(row[4]).trim();
    const status = String(row[5]).trim().toUpperCase();

    if (
      hariJadwal.toLowerCase() === hari.toLowerCase() &&
      status === 'AKTIF'
    ) {

      return {
        id: id,
        hari: hari,
        jamMulai: jamMulai,
        jamSelesai: jamSelesai,
        jenis: jenis,
        status: status
      };
    }
  }

  return null;
}


/**
 * Mengubah nilai jam menjadi HH:mm
 */
function formatJam_(nilai) {

  if (nilai instanceof Date) {

    return Utilities.formatDate(
      nilai,
      Session.getScriptTimeZone(),
      'HH:mm'
    );

  }

  if (typeof nilai === 'string') {

    const teks = nilai.trim();

    if (/^\d{1,2}:\d{2}$/.test(teks)) {

      const bagian = teks.split(':');

      return (
        String(parseInt(bagian[0], 10)).padStart(2, '0') +
        ':' +
        bagian[1]
      );
    }
  }

  return '';
}


/**
 * Mengubah HH:mm menjadi menit
 */
function jamKeMenit_(jam) {

  if (!jam) return null;

  const bagian = jam.split(':');

  return (
    parseInt(bagian[0], 10) * 60 +
    parseInt(bagian[1], 10)
  );
}


/**
 * Menghitung keterlambatan
 */
function hitungKeterlambatan_(jamScan, jamMulai) {

  const scanMenit = jamKeMenit_(jamScan);
  const mulaiMenit = jamKeMenit_(jamMulai);

  if (scanMenit === null || mulaiMenit === null) {
    return 0;
  }

  const terlambat = scanMenit - mulaiMenit;

  return terlambat > 0 ? terlambat : 0;
}


/**
 * Menentukan status kehadiran
 */
function tentukanStatusAbsensi_(menitTerlambat) {

  if (menitTerlambat > 0) {
    return 'TERLAMBAT';
  }

  return 'HADIR';
}


/**
 * Proses utama absensi masuk
 */
function prosesMasukPulang_(nis, jenisAbsensi) {

  const lock = LockService.getScriptLock();

  lock.waitLock(10000);

  try {

    nis = String(nis).trim();

    if (!nis) {
      throw new Error('NIS belum diisi.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetSiswa =
      ss.getSheetByName(NAMA_SHEET_SISWA);

    const sheetAbsensi =
      ss.getSheetByName(NAMA_SHEET_ABSENSI);

    if (!sheetSiswa) {
      throw new Error('Sheet 02_Siswa tidak ditemukan.');
    }

    if (!sheetAbsensi) {
      throw new Error('Sheet 05_Absensi tidak ditemukan.');
    }


    /**********************************************
     * CARI SISWA
     **********************************************/

    const dataSiswa =
      sheetSiswa.getDataRange().getValues();

    let siswa = null;

    for (let i = 1; i < dataSiswa.length; i++) {

      const row = dataSiswa[i];

      const nisData =
        String(row[0]).trim();

      if (nisData === nis) {

        siswa = {
          nis: nisData,
          nama: row[1],
          kelas: row[2],
          jk: row[3]
        };

        break;
      }
    }

    if (!siswa) {
      throw new Error(
        'Siswa dengan NIS ' + nis + ' tidak ditemukan.'
      );
    }


    /**********************************************
     * WAKTU SEKARANG
     **********************************************/

    const sekarang = new Date();

    const tanggal =
      Utilities.formatDate(
        sekarang,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );

    const jam =
      Utilities.formatDate(
        sekarang,
        Session.getScriptTimeZone(),
        'HH:mm'
      );


    /**********************************************
     * CEK JADWAL
     **********************************************/

    const jadwal =
      getJadwalSekolah_(sekarang);

    if (!jadwal) {

      return {
        success: false,
        status: 'LIBUR',
        message:
          'Hari ini tidak terdapat jadwal sekolah aktif.'
      };
    }


    /**********************************************
     * HITUNG TERLAMBAT
     **********************************************/

    const terlambat =
      hitungKeterlambatan_(
        jam,
        jadwal.jamMulai
      );

    const status =
      tentukanStatusAbsensi_(
        terlambat
      );


    /**********************************************
     * CEK ABSENSI HARI INI
     **********************************************/

    const dataAbsensi =
      sheetAbsensi.getDataRange().getValues();

    let sudahAbsen = false;

    for (let i = 1; i < dataAbsensi.length; i++) {

      const row = dataAbsensi[i];

      const tanggalData =
        Utilities.formatDate(
          new Date(row[1]),
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );

      const nisData =
        String(row[3]).trim();

      if (
        tanggalData === tanggal &&
        nisData === nis
      ) {

        sudahAbsen = true;
        break;
      }
    }


    if (sudahAbsen) {

      return {
        success: false,
        status: 'SUDAH_ABSEN',
        message:
          siswa.nama +
          ' sudah melakukan absensi hari ini.'
      };
    }


    /**********************************************
     * SIMPAN ABSENSI
     **********************************************/

    sheetAbsensi.appendRow([

      sekarang,              // TIMESTAMP
      tanggal,               // TANGGAL
      jam,                   // JAM MASUK
      siswa.nis,             // NIS
      siswa.nama,            // NAMA
      siswa.kelas,           // KELAS
      siswa.jk,              // JK
      jenisAbsensi || 'MASUK',
      status,                // STATUS
      '',                    // JAM PULANG
      jadwal.jamMulai,       // BATAS MASUK
      terlambat,             // TERLAMBAT MENIT
      jadwal.jamSelesai      // JADWAL PULANG

    ]);


    /**********************************************
     * HASIL
     **********************************************/

    return {

      success: true,

      status: status,

      nis: siswa.nis,

      nama: siswa.nama,

      kelas: siswa.kelas,

      tanggal: tanggal,

      jamMasuk: jam,

      batasMasuk: jadwal.jamMulai,

      jamPulang: jadwal.jamSelesai,

      terlambat: terlambat,

      jenis: jenisAbsensi || 'MASUK',

      message:
        'Absensi berhasil. ' +
        siswa.nama +
        ' - ' +
        status

    };

  } finally {

    lock.releaseLock();

  }
}
/****************************************************
 * MODUL GPS / GEOFENCE
 ****************************************************/

const NAMA_SHEET_SETTING = '01_Setting';


/**
 * Membaca setting berdasarkan parameter
 */
function getSetting_(parameter) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NAMA_SHEET_SETTING);

  if (!sheet) {
    throw new Error('Sheet 01_Setting tidak ditemukan.');
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    const key = String(data[i][0]).trim();

    if (key === parameter) {
      return data[i][1];
    }
  }

  return '';
}


/**
 * Menghitung jarak dua koordinat
 * menggunakan rumus Haversine
 */
function hitungJarakMeter_(lat1, lon1, lat2, lon2) {

  const R = 6371000;

  const dLat =
    (lat2 - lat1) * Math.PI / 180;

  const dLon =
    (lon2 - lon1) * Math.PI / 180;

  const a =
    Math.sin(dLat / 2) *
    Math.sin(dLat / 2) +

    Math.cos(lat1 * Math.PI / 180) *
    Math.cos(lat2 * Math.PI / 180) *

    Math.sin(dLon / 2) *
    Math.sin(dLon / 2);

  const c =
    2 * Math.atan2(
      Math.sqrt(a),
      Math.sqrt(1 - a)
    );

  return R * c;
}


/**
 * Validasi lokasi siswa
 */
function validasiLokasiSekolah_(latitude, longitude) {

  const latSekolah =
    parseFloat(
      getSetting_('LATITUDE_SEKOLAH')
    );

  const lonSekolah =
    parseFloat(
      getSetting_('LONGITUDE_SEKOLAH')
    );

  const radius =
    parseFloat(
      getSetting_('RADIUS_SEKOLAH_METER')
    ) || 150;


  if (
    isNaN(latSekolah) ||
    isNaN(lonSekolah)
  ) {

    return {
      valid: false,
      message:
        'Koordinat sekolah belum diatur.'
    };
  }


  const lat =
    parseFloat(latitude);

  const lon =
    parseFloat(longitude);


  if (
    isNaN(lat) ||
    isNaN(lon)
  ) {

    return {
      valid: false,
      message:
        'Lokasi GPS tidak valid.'
    };
  }


  const jarak =
    hitungJarakMeter_(
      latSekolah,
      lonSekolah,
      lat,
      lon
    );


  return {

    valid: jarak <= radius,

    jarak: Math.round(jarak),

    radius: radius,

    latitude: lat,

    longitude: lon,

    message:
      jarak <= radius
        ? 'Lokasi berada di area sekolah.'
        : 'Anda berada di luar area sekolah.'

  };
}
/****************************************************
 * ABSENSI DARI QR CODE
 ****************************************************/

function prosesAbsensiQR_(nis, latitude, longitude) {

  nis = String(nis).trim();

  if (!nis) {

    return {
      success: false,
      message: 'QR Code tidak berisi NIS.'
    };

  }


  // Validasi GPS
  const lokasi =
    validasiLokasiSekolah_(
      latitude,
      longitude
    );


  if (!lokasi.valid) {

    return {

      success: false,

      type: 'GPS',

      message:
        lokasi.message,

      jarak:
        lokasi.jarak || null,

      radius:
        lokasi.radius || null

    };

  }


  // Proses absensi
  const hasil =
    prosesMasukPulang_(
      nis,
      'QR'
    );


  if (!hasil.success) {

    return {

      ...hasil,

      jarak:
        lokasi.jarak

    };

  }


  return {

    ...hasil,

    jarak:
      lokasi.jarak,

    radius:
      lokasi.radius,

    latitude:
      lokasi.latitude,

    longitude:
      lokasi.longitude

  };

}
/*******************************************************
 * DATABASE SISWA
 * CRUD SISWA
 *******************************************************/

/**
 * Mengambil seluruh data siswa
 */
function getDataSiswa() {

  const sheet = getSheet_(SHEET.SISWA);

  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const data = sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      HEADER.SISWA.length
    )
    .getValues();


  return data.map(function(row, index) {

    return {

      baris: index + 2,

      nis: row[0] || '',

      nama: row[1] || '',

      kelas: row[2] || '',

      jk: row[3] || '',

      tempatLahir: row[4] || '',

      tanggalLahir:
        formatTanggalSiswa_(row[5]),

      alamat: row[6] || '',

      namaOrtu: row[7] || '',

      noWaOrtu: row[8] || '',

      status:
        row[9] || 'AKTIF',

      keterangan:
        row[10] || ''

    };

  });

}


/**
 * Tambah siswa baru
 */
function tambahSiswa(data) {

  if (!data) {
    throw new Error(
      'Data siswa tidak ditemukan.'
    );
  }


  const nis =
    String(data.nis || '').trim();

  const nama =
    String(data.nama || '').trim();


  if (!nis) {
    throw new Error(
      'NIS wajib diisi.'
    );
  }


  if (!nama) {
    throw new Error(
      'Nama siswa wajib diisi.'
    );
  }


  const sheet =
    getSheet_(SHEET.SISWA);


  /***************************************************
   * CEK NIS DUPLIKAT
   ***************************************************/

  const lastRow =
    sheet.getLastRow();


  if (lastRow > 1) {

    const dataNIS =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    for (
      let i = 0;
      i < dataNIS.length;
      i++
    ) {

      const nisLama =
        String(
          dataNIS[i][0]
        ).trim();


      if (
        nisLama === nis
      ) {

        throw new Error(
          'NIS ' +
          nis +
          ' sudah terdaftar.'
        );

      }

    }

  }


  /***************************************************
   * DATA YANG DISIMPAN
   ***************************************************/

  const baris = [

    nis,

    nama,

    String(
      data.kelas || ''
    ).trim(),

    String(
      data.jk || ''
    ).trim(),

    String(
      data.tempatLahir || ''
    ).trim(),

    data.tanggalLahir
      ? new Date(
          data.tanggalLahir
        )
      : '',

    String(
      data.alamat || ''
    ).trim(),

    String(
      data.namaOrtu || ''
    ).trim(),

    String(
      data.noWaOrtu || ''
    ).trim(),

    String(
      data.status || 'AKTIF'
    ).trim(),

    String(
      data.keterangan || ''
    ).trim()

  ];


  sheet.appendRow(baris);


  return {

    success: true,

    message:
      'Data siswa ' +
      nama +
      ' berhasil ditambahkan.'

  };

}


/**
 * Edit siswa
 */
function editSiswa(
  nomorBaris,
  data
) {

  if (!data) {

    throw new Error(
      'Data siswa tidak ditemukan.'
    );

  }


  nomorBaris =
    Number(nomorBaris);


  if (
    !nomorBaris ||
    nomorBaris < 2
  ) {

    throw new Error(
      'Nomor baris siswa tidak valid.'
    );

  }


  const nis =
    String(
      data.nis || ''
    ).trim();


  const nama =
    String(
      data.nama || ''
    ).trim();


  if (!nis) {

    throw new Error(
      'NIS wajib diisi.'
    );

  }


  if (!nama) {

    throw new Error(
      'Nama siswa wajib diisi.'
    );

  }


  const sheet =
    getSheet_(SHEET.SISWA);


  if (
    nomorBaris >
    sheet.getLastRow()
  ) {

    throw new Error(
      'Data siswa tidak ditemukan.'
    );

  }


  /***************************************************
   * CEK NIS DUPLIKAT
   * kecuali baris yang sedang diedit
   ***************************************************/

  const lastRow =
    sheet.getLastRow();


  if (lastRow > 1) {

    const dataNIS =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues();


    for (
      let i = 0;
      i < dataNIS.length;
      i++
    ) {

      const barisSekarang =
        i + 2;


      if (
        barisSekarang ===
        nomorBaris
      ) {

        continue;

      }


      const nisLama =
        String(
          dataNIS[i][0]
        ).trim();


      if (
        nisLama === nis
      ) {

        throw new Error(
          'NIS ' +
          nis +
          ' sudah digunakan siswa lain.'
        );

      }

    }

  }


  /***************************************************
   * UPDATE DATA
   ***************************************************/

  const baris = [

    nis,

    nama,

    String(
      data.kelas || ''
    ).trim(),

    String(
      data.jk || ''
    ).trim(),

    String(
      data.tempatLahir || ''
    ).trim(),

    data.tanggalLahir
      ? new Date(
          data.tanggalLahir
        )
      : '',

    String(
      data.alamat || ''
    ).trim(),

    String(
      data.namaOrtu || ''
    ).trim(),

    String(
      data.noWaOrtu || ''
    ).trim(),

    String(
      data.status || 'AKTIF'
    ).trim(),

    String(
      data.keterangan || ''
    ).trim()

  ];


  sheet
    .getRange(
      nomorBaris,
      1,
      1,
      HEADER.SISWA.length
    )
    .setValues([
      baris
    ]);


  return {

    success: true,

    message:
      'Data siswa ' +
      nama +
      ' berhasil diperbarui.'

  };

}


/**
 * Hapus siswa
 */
function hapusSiswa(
  nomorBaris
) {

  nomorBaris =
    Number(nomorBaris);


  if (
    !nomorBaris ||
    nomorBaris < 2
  ) {

    throw new Error(
      'Nomor baris tidak valid.'
    );

  }


  const sheet =
    getSheet_(SHEET.SISWA);


  if (
    nomorBaris >
    sheet.getLastRow()
  ) {

    throw new Error(
      'Data siswa tidak ditemukan.'
    );

  }


  sheet.deleteRow(
    nomorBaris
  );


  return {

    success: true,

    message:
      'Data siswa berhasil dihapus.'

  };

}


/**
 * Format tanggal untuk ditampilkan
 */
function formatTanggalSiswa_(tanggal) {

  if (!tanggal) {
    return '';
  }


  if (
    Object.prototype.toString
      .call(tanggal) ===
    '[object Date]'
  ) {

    return Utilities
      .formatDate(
        tanggal,
        Session.getScriptTimeZone(),
        'yyyy-MM-dd'
      );

  }


  return String(tanggal);

}
/**
 * ============================================================
 * UPGRADE DATABASE QR CODE
 * Menambahkan KODE_QR tanpa menghapus data yang sudah ada.
 * ============================================================
 */
function upgradeDatabaseQR() {

  const ss = getDatabase_();

  // =========================
  // SISWA
  // =========================
  const sheetSiswa = ss.getSheetByName(SHEET.SISWA);

  if (!sheetSiswa) {
    throw new Error('Sheet SISWA tidak ditemukan.');
  }

  tambahKolomQR_(sheetSiswa, 'KODE_QR');


  // =========================
  // GURU
  // =========================
  const sheetGuru = ss.getSheetByName(SHEET.GURU);

  if (!sheetGuru) {
    throw new Error('Sheet GURU tidak ditemukan.');
  }

  tambahKolomQR_(sheetGuru, 'KODE_QR');


  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'KODE_QR berhasil ditambahkan ke SISWA dan GURU.'
  };
}


/**
 * Menambahkan kolom KODE_QR
 * jika belum tersedia.
 */
function tambahKolomQR_(sheet, namaKolom) {

  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    sheet.getRange(1, 1).setValue(namaKolom);
    return;
  }


  const header = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(value) {
      return String(value).trim().toUpperCase();
    });


  // Jika sudah ada, jangan lakukan apa-apa.
  if (header.indexOf(namaKolom) !== -1) {
    return;
  }


  // Tambahkan di kolom paling kanan.
  sheet
    .getRange(1, lastColumn + 1)
    .setValue(namaKolom)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');


  sheet.setFrozenRows(1);
}
function tesQRDatabaseBaru() {
  const qrCode = 'NIS:121235180009250020';

  let nis = String(qrCode).trim();

  if (nis.toUpperCase().startsWith('NIS:')) {
    nis = nis.substring(4).trim();
  }

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(SHEET.SISWA);

  if (!sheet) {
    throw new Error('Sheet SISWA tidak ditemukan.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error('Database SISWA masih kosong.');
  }

  const data = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // Kolom SISWA:
    // A = NIS
    // B = NAMA
    // C = KELAS
    const nisData = String(row[0] || '').trim();

    if (nisData === nis) {
      const hasil = {
        success: true,
        ditemukan: true,
        nis: nisData,
        nama: String(row[1] || '').trim(),
        kelas: String(row[2] || '').trim(),
        qr: qrCode,
        message: 'QR lama berhasil ditemukan di database baru.'
      };

      Logger.log(JSON.stringify(hasil, null, 2));
      return hasil;
    }
  }

  const hasil = {
    success: false,
    ditemukan: false,
    nis: nis,
    qr: qrCode,
    message: 'NIS ' + nis + ' belum ditemukan di sheet SISWA.'
  };

  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}
function tesKodeQRSiswa() {
  const nisTarget = '121235180009250020';

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(SHEET.SISWA);

  if (!sheet) {
    throw new Error('Sheet SISWA tidak ditemukan.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error('Data SISWA kosong.');
  }

  const header = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(v) {
      return String(v).trim().toUpperCase();
    });

  const kolomNIS = header.indexOf('NIS');
  const kolomQR = header.indexOf('KODE_QR');

  if (kolomNIS === -1) {
    throw new Error('Kolom NIS tidak ditemukan.');
  }

  if (kolomQR === -1) {
    throw new Error('Kolom KODE_QR tidak ditemukan.');
  }

  const data = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const nis = String(data[i][kolomNIS] || '').trim();

    if (nis === nisTarget) {
      const qr = String(data[i][kolomQR] || '').trim();

      const hasil = {
        success: true,
        nis: nis,
        kodeQR: qr,
        sesuai: qr === 'NIS:' + nis,
        message: qr === 'NIS:' + nis
          ? 'KODE_QR sudah sesuai dengan QR lama.'
          : 'KODE_QR belum sesuai atau masih kosong.'
      };

      Logger.log(JSON.stringify(hasil, null, 2));
      return hasil;
    }
  }

  throw new Error('NIS tidak ditemukan: ' + nisTarget);
}
function isiKodeQRLama() {
  const ss = getDatabase_();
  const sheet = ss.getSheetByName(SHEET.SISWA);

  if (!sheet) {
    throw new Error('Sheet SISWA tidak ditemukan.');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 1) {
    throw new Error('Data SISWA masih kosong.');
  }

  const header = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(v) {
      return String(v).trim().toUpperCase();
    });

  const kolomNIS = header.indexOf('NIS');
  const kolomQR = header.indexOf('KODE_QR');

  if (kolomNIS === -1) {
    throw new Error('Kolom NIS tidak ditemukan.');
  }

  if (kolomQR === -1) {
    throw new Error('Kolom KODE_QR tidak ditemukan.');
  }

  const data = sheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getValues();

  let jumlahDiisi = 0;

  for (let i = 0; i < data.length; i++) {
    const nis = String(data[i][kolomNIS] || '').trim();
    const qrLama = String(data[i][kolomQR] || '').trim();

    // Hanya isi jika NIS ada dan KODE_QR masih kosong
    if (nis && !qrLama) {
      data[i][kolomQR] = 'NIS:' + nis;
      jumlahDiisi++;
    }
  }

  sheet
    .getRange(2, 1, data.length, lastColumn)
    .setValues(data);

  const hasil = {
    success: true,
    jumlahDiisi: jumlahDiisi,
    message: jumlahDiisi + ' KODE_QR berhasil diisi dengan format QR lama.'
  };

  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}
function tesCocokkanQR() {
  const qrCode = 'NIS:121235180009240002';

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(SHEET.SISWA);

  if (!sheet) {
    throw new Error('Sheet SISWA tidak ditemukan.');
  }

  const data = sheet.getDataRange().getDisplayValues();
  const header = data[0].map(function(v) {
    return String(v).trim().toUpperCase();
  });

  const kolomNIS = header.indexOf('NIS');
  const kolomNama = header.indexOf('NAMA');
  const kolomKelas = header.indexOf('KELAS');
  const kolomQR = header.indexOf('KODE_QR');

  if ([kolomNIS, kolomNama, kolomKelas, kolomQR].includes(-1)) {
    throw new Error('Ada kolom wajib yang belum ditemukan.');
  }

  for (let i = 1; i < data.length; i++) {
    if (
      data[i][kolomQR].trim() === qrCode &&
      data[i][kolomNIS].trim() === qrCode.substring(4)
    ) {
      const hasil = {
        success: true,
        nis: data[i][kolomNIS],
        nama: data[i][kolomNama],
        kelas: data[i][kolomKelas],
        kodeQR: data[i][kolomQR],
        message: 'QR cocok dengan data siswa.'
      };

      Logger.log(JSON.stringify(hasil, null, 2));
      return hasil;
    }
  }

  throw new Error('QR tidak cocok dengan data siswa.');
}

/**
 * TES BACKEND ABSENSI SISWA
 * Hanya membaca database.
 * Tidak menyimpan atau mengubah absensi.
 */
function tesBackendAbsensiSiswa() {
  const qrCode = 'NIS:121235180009240002';

  const nis = String(qrCode)
    .trim()
    .replace(/^NIS:/i, '')
    .trim();

  if (!nis) {
    throw new Error('NIS dari QR kosong.');
  }

  const ss = getDatabase_();
  const sheetSiswa = ss.getSheetByName(SHEET.SISWA);
  const sheetAbsensi = ss.getSheetByName(SHEET.ABSENSI_SISWA);

  if (!sheetSiswa || !sheetAbsensi) {
    throw new Error(
      'Sheet SISWA atau ABSENSI_SISWA tidak ditemukan.'
    );
  }

  const dataSiswa = sheetSiswa.getDataRange().getDisplayValues();

  if (dataSiswa.length < 2) {
    throw new Error('Belum ada data siswa.');
  }

  const headerSiswa = dataSiswa[0].map(function(v) {
    return String(v).trim().toUpperCase();
  });

  const idxNIS = headerSiswa.indexOf('NIS');
  const idxNama = headerSiswa.indexOf('NAMA');
  const idxKelas = headerSiswa.indexOf('KELAS');
  const idxJK = headerSiswa.indexOf('JK');
  const idxQR = headerSiswa.indexOf('KODE_QR');

  if ([idxNIS, idxNama, idxKelas, idxJK, idxQR].includes(-1)) {
    throw new Error('Header wajib pada sheet SISWA belum lengkap.');
  }

  let siswa = null;

  for (let i = 1; i < dataSiswa.length; i++) {
    if (
      String(dataSiswa[i][idxNIS]).trim() === nis &&
      String(dataSiswa[i][idxQR]).trim() === 'NIS:' + nis
    ) {
      siswa = {
        nis: nis,
        nama: dataSiswa[i][idxNama],
        kelas: dataSiswa[i][idxKelas],
        jk: dataSiswa[i][idxJK]
      };
      break;
    }
  }

  if (!siswa) {
    throw new Error('QR tidak cocok dengan data siswa.');
  }

  const dataAbsensi = sheetAbsensi.getDataRange().getValues();

  const headerAbsensi = dataAbsensi[0].map(function(v) {
    return String(v).trim().toUpperCase();
  });

  const idxTanggal = headerAbsensi.indexOf('TANGGAL');
  const idxNISAbsensi = headerAbsensi.indexOf('NIS');
  const idxKegiatan = headerAbsensi.indexOf('KEGIATAN');

  if ([idxTanggal, idxNISAbsensi, idxKegiatan].includes(-1)) {
    throw new Error(
      'Header TANGGAL, NIS, atau KEGIATAN tidak ditemukan pada ABSENSI_SISWA.'
    );
  }

  const zonaWaktu = Session.getScriptTimeZone();
  const sekarang = new Date();
  const tanggalHariIni = Utilities.formatDate(
    sekarang,
    zonaWaktu,
    'yyyy-MM-dd'
  );

  const kegiatanHariIni = [];

  for (let i = 1; i < dataAbsensi.length; i++) {
    const row = dataAbsensi[i];

    if (String(row[idxNISAbsensi]).trim() !== nis) {
      continue;
    }

    let tanggalBaris = '';

    if (row[idxTanggal] instanceof Date) {
      tanggalBaris = Utilities.formatDate(
        row[idxTanggal],
        zonaWaktu,
        'yyyy-MM-dd'
      );
    } else {
      const teksTanggal = String(row[idxTanggal] || '').trim();

      if (/^\d{4}-\d{2}-\d{2}/.test(teksTanggal)) {
        tanggalBaris = teksTanggal.substring(0, 10);
      }
    }

    if (tanggalBaris === tanggalHariIni) {
      kegiatanHariIni.push(
        String(row[idxKegiatan] || '').trim().toUpperCase()
      );
    }
  }

  let tindakanBerikutnya;

  if (
    kegiatanHariIni.includes('MASUK') &&
    kegiatanHariIni.includes('PULANG')
  ) {
    tindakanBerikutnya = 'SELESAI';
  } else if (kegiatanHariIni.includes('MASUK')) {
    tindakanBerikutnya = 'PULANG';
  } else if (kegiatanHariIni.includes('PULANG')) {
    tindakanBerikutnya = 'PERIKSA_DATA';
  } else {
    tindakanBerikutnya = 'MASUK';
  }

  const hasil = {
    success: true,
    mode: 'TES_SAJA_TIDAK_MENYIMPAN',
    nis: siswa.nis,
    nama: siswa.nama,
    kelas: siswa.kelas,
    qr: 'NIS:' + siswa.nis,
    tanggal: tanggalHariIni,
    kegiatanHariIni: kegiatanHariIni,
    tindakanBerikutnya: tindakanBerikutnya,
    message: 'Pemeriksaan selesai. Tidak ada absensi yang disimpan.'
  };

  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

/**
 * BACKEND ABSENSI BARU - MODE SIMULASI
 * Hanya menentukan tindakan berikutnya.
 * Tidak menyimpan absensi.
 */
function simulasiAbsensiBaru(qrCode) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    return {
      success: false,
      message: 'Sistem sedang memproses scan lain. Coba lagi.'
    };
  }

  try {
    const qr = String(qrCode || '').trim();
    const nis = qr.replace(/^NIS:/i, '').trim();

    if (!nis) {
      return {
        success: false,
        message: 'QR kosong atau tidak berisi NIS.'
      };
    }

    const ss = getDatabase_();
    const siswaSheet = ss.getSheetByName(SHEET.SISWA);
    const absensiSheet = ss.getSheetByName(SHEET.ABSENSI_SISWA);

    if (!siswaSheet || !absensiSheet) {
      throw new Error('Sheet database tidak ditemukan.');
    }

    const siswaData = siswaSheet.getDataRange().getDisplayValues();
    if (siswaData.length < 2) {
      throw new Error('Data siswa masih kosong.');
    }

    const siswaHeader = siswaData[0].map(v =>
      String(v).trim().toUpperCase()
    );

    const iNis = siswaHeader.indexOf('NIS');
    const iNama = siswaHeader.indexOf('NAMA');
    const iKelas = siswaHeader.indexOf('KELAS');
    const iJK = siswaHeader.indexOf('JK');
    const iQR = siswaHeader.indexOf('KODE_QR');

    if ([iNis, iNama, iKelas, iJK, iQR].includes(-1)) {
      throw new Error('Header SISWA tidak lengkap.');
    }

    let siswa = null;

    for (let i = 1; i < siswaData.length; i++) {
      const row = siswaData[i];

      if (
        String(row[iNis]).trim() === nis &&
        String(row[iQR]).trim().toUpperCase() === ('NIS:' + nis).toUpperCase()
      ) {
        siswa = {
          nis: nis,
          nama: row[iNama],
          kelas: row[iKelas],
          jk: row[iJK]
        };
        break;
      }
    }

    if (!siswa) {
      return {
        success: false,
        message: 'QR tidak cocok dengan data siswa.'
      };
    }

    const absensiData = absensiSheet.getDataRange().getValues();
    const absensiHeader = absensiData[0].map(v =>
      String(v).trim().toUpperCase()
    );

    const iTanggal = absensiHeader.indexOf('TANGGAL');
    const iNisAbsensi = absensiHeader.indexOf('NIS');
    const iKegiatan = absensiHeader.indexOf('KEGIATAN');

    if ([iTanggal, iNisAbsensi, iKegiatan].includes(-1)) {
      throw new Error('Header ABSENSI_SISWA tidak sesuai.');
    }

    const zona = Session.getScriptTimeZone();
    const hariIni = Utilities.formatDate(
      new Date(), zona, 'yyyy-MM-dd'
    );

    const kegiatanHariIni = [];

    for (let i = 1; i < absensiData.length; i++) {
      const row = absensiData[i];

      if (String(row[iNisAbsensi]).trim() !== nis) continue;

      let tanggal = '';
      if (row[iTanggal] instanceof Date) {
        tanggal = Utilities.formatDate(row[iTanggal], zona, 'yyyy-MM-dd');
      } else {
        tanggal = String(row[iTanggal] || '').trim().substring(0, 10);
      }

      if (tanggal === hariIni) {
        kegiatanHariIni.push(
          String(row[iKegiatan] || '').trim().toUpperCase()
        );
      }
    }

    let tindakan;

    if (
      kegiatanHariIni.includes('MASUK') &&
      kegiatanHariIni.includes('PULANG')
    ) {
      tindakan = 'SELESAI';
    } else if (kegiatanHariIni.includes('MASUK')) {
      tindakan = 'PULANG';
    } else if (kegiatanHariIni.includes('PULANG')) {
      tindakan = 'PERIKSA_DATA';
    } else {
      tindakan = 'MASUK';
    }

    return {
      success: true,
      mode: 'SIMULASI_TIDAK_MENYIMPAN',
      nis: siswa.nis,
      nama: siswa.nama,
      kelas: siswa.kelas,
      qr: 'NIS:' + siswa.nis,
      tanggal: hariIni,
      kegiatanHariIni: kegiatanHariIni,
      tindakanBerikutnya: tindakan,
      message: 'Simulasi selesai. Tidak ada data yang disimpan.'
    };

  } finally {
    lock.releaseLock();
  }
}

function tesSimulasiAbsensiBaru() {
  const hasil = simulasiAbsensiBaru('NIS:121235180009240002');
  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

/**
 * BACKEND ABSENSI SISWA BARU
 * Mendukung QR lama: NIS:123456
 * MASUK dan PULANG disimpan pada baris berbeda.
 */
function simpanAbsensiSiswaBaru_(qrCode, opsi) {
  opsi = opsi || {};
  const simulasi = opsi.simulasi !== false;

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    return {
      success: false,
      message: 'Sistem sedang memproses scan lain. Silakan coba lagi.'
    };
  }

  try {
    const qr = String(qrCode || '').trim();
    const nis = qr.replace(/^NIS:/i, '').trim();

    if (!nis) {
      return {
        success: false,
        message: 'QR kosong atau tidak berisi NIS.'
      };
    }

    const ss = getDatabase_();
    const sheetSiswa = ss.getSheetByName(SHEET.SISWA);
    const sheetAbsensi = ss.getSheetByName(SHEET.ABSENSI_SISWA);

    if (!sheetSiswa || !sheetAbsensi) {
      throw new Error('Sheet SISWA atau ABSENSI_SISWA tidak ditemukan.');
    }

    // Cari siswa berdasarkan NIS DAN KODE_QR.
    const dataSiswa = sheetSiswa.getDataRange().getDisplayValues();

    if (dataSiswa.length < 2) {
      throw new Error('Database siswa masih kosong.');
    }

    const hs = dataSiswa[0].map(v =>
      String(v).trim().toUpperCase()
    );

    const iNIS = hs.indexOf('NIS');
    const iNama = hs.indexOf('NAMA');
    const iKelas = hs.indexOf('KELAS');
    const iJK = hs.indexOf('JK');
    const iQR = hs.indexOf('KODE_QR');

    if ([iNIS, iNama, iKelas, iJK, iQR].includes(-1)) {
      throw new Error('Header wajib pada sheet SISWA tidak lengkap.');
    }

    let siswa = null;

    for (let i = 1; i < dataSiswa.length; i++) {
      const row = dataSiswa[i];

      if (
        String(row[iNIS]).trim() === nis &&
        String(row[iQR]).trim().toUpperCase() ===
          ('NIS:' + nis).toUpperCase()
      ) {
        siswa = {
          nis: nis,
          nama: row[iNama],
          kelas: row[iKelas],
          jk: row[iJK]
        };
        break;
      }
    }

    if (!siswa) {
      return {
        success: false,
        message: 'QR tidak cocok dengan data siswa.'
      };
    }

    // Periksa struktur sheet absensi.
    const lastColumn = sheetAbsensi.getLastColumn();
    const header = sheetAbsensi
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(v => String(v).trim().toUpperCase());

    const iTanggal = header.indexOf('TANGGAL');
    const iJam = header.indexOf('JAM');
    const iNisAbsensi = header.indexOf('NIS');
    const iKegiatan = header.indexOf('KEGIATAN');

    if ([iTanggal, iJam, iNisAbsensi, iKegiatan].includes(-1)) {
      throw new Error(
        'Header TANGGAL, JAM, NIS, atau KEGIATAN tidak ditemukan.'
      );
    }

    const sekarang = new Date();
    const zona = Session.getScriptTimeZone();
    const hariIni = Utilities.formatDate(
      sekarang, zona, 'yyyy-MM-dd'
    );

    const dataAbsensi = sheetAbsensi.getDataRange().getValues();
    const kegiatanHariIni = [];

    for (let i = 1; i < dataAbsensi.length; i++) {
      const row = dataAbsensi[i];

      if (String(row[iNisAbsensi]).trim() !== nis) continue;

      let tanggalBaris = '';

      if (row[iTanggal] instanceof Date) {
        tanggalBaris = Utilities.formatDate(
          row[iTanggal], zona, 'yyyy-MM-dd'
        );
      } else {
        const t = String(row[iTanggal] || '').trim();

        if (/^\d{4}-\d{2}-\d{2}/.test(t)) {
          tanggalBaris = t.substring(0, 10);
        }
      }

      if (tanggalBaris === hariIni) {
        kegiatanHariIni.push({
          kegiatan: String(row[iKegiatan] || '')
            .trim().toUpperCase(),
          waktu: row[iJam]
        });
      }
    }

    const sudahMasuk = kegiatanHariIni.some(
      x => x.kegiatan === 'MASUK'
    );
    const sudahPulang = kegiatanHariIni.some(
      x => x.kegiatan === 'PULANG'
    );

    let tindakan;

    if (sudahMasuk && sudahPulang) {
      return {
        success: false,
        tindakanBerikutnya: 'SELESAI',
        message: 'Absensi MASUK dan PULANG hari ini sudah tercatat.'
      };
    }

    if (sudahPulang && !sudahMasuk) {
      return {
        success: false,
        tindakanBerikutnya: 'PERIKSA_DATA',
        message: 'Ada catatan PULANG tanpa MASUK. Periksa data terlebih dahulu.'
      };
    }

    if (!sudahMasuk) {
      tindakan = 'MASUK';
    } else {
      // Cegah scan kedua yang terlalu cepat dianggap sebagai PULANG.
      const catatanMasuk = kegiatanHariIni.find(
        x => x.kegiatan === 'MASUK'
      );

      let menitSejakMasuk = null;

      if (catatanMasuk && catatanMasuk.waktu instanceof Date) {
        menitSejakMasuk =
          (sekarang.getTime() - catatanMasuk.waktu.getTime()) / 60000;
      }

      if (menitSejakMasuk !== null && menitSejakMasuk < 5) {
        return {
          success: false,
          tindakanBerikutnya: 'TUNGGU_PULANG',
          message: 'Absensi MASUK baru tercatat. Tunggu minimal 5 menit sebelum tes PULANG.'
        };
      }

      tindakan = 'PULANG';
    }

    const jam = Utilities.formatDate(
      sekarang, zona, 'HH:mm:ss'
    );

    // Mode simulasi tidak menulis apa pun.
    if (simulasi) {
      return {
        success: true,
        mode: 'SIMULASI_TIDAK_MENYIMPAN',
        nis: siswa.nis,
        nama: siswa.nama,
        kelas: siswa.kelas,
        kegiatan: tindakan,
        tanggal: hariIni,
        jam: jam,
        message: 'Simulasi berhasil. Tidak ada baris absensi yang ditambahkan.'
      };
    }

    // Bentuk baris sesuai nama header, bukan posisi kolom tetap.
    const barisBaru = new Array(lastColumn).fill('');

    const isi = {
      TIMESTAMP: sekarang,
      TANGGAL: new Date(
        Number(Utilities.formatDate(sekarang, zona, 'yyyy')),
        Number(Utilities.formatDate(sekarang, zona, 'MM')) - 1,
        Number(Utilities.formatDate(sekarang, zona, 'dd'))
      ),
      JAM: sekarang,
      NIS: siswa.nis,
      NAMA: siswa.nama,
      KELAS: siswa.kelas,
      JK: siswa.jk,
      KEGIATAN: tindakan,
      JAM_KE: '',
      STATUS: 'HADIR',
      KETERANGAN: 'QR'
    };

    Object.keys(isi).forEach(function(namaKolom) {
      const idx = header.indexOf(namaKolom);

      if (idx !== -1) {
        barisBaru[idx] = isi[namaKolom];
      }
    });

    sheetAbsensi.appendRow(barisBaru);
    SpreadsheetApp.flush();

    return {
      success: true,
      mode: 'TERSIMPAN',
      nis: siswa.nis,
      nama: siswa.nama,
      kelas: siswa.kelas,
      kegiatan: tindakan,
      jamMasuk: jam,
      status: 'HADIR',
      message: 'Absensi ' + tindakan + ' berhasil disimpan.'
    };

  } finally {
    lock.releaseLock();
  }
}


/**
 * Tes baca saja. Tidak menulis absensi.
 */
function tesSimpanAbsensiSiswaBaru() {
  const hasil = simpanAbsensiSiswaBaru_(
    'NIS:121235180009240002',
    { simulasi: true }
  );

  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

/**
 * TES ATURAN WAKTU ABSENSI
 * Hanya membaca data. Tidak menyimpan absensi.
 */
function tesAturanWaktuAbsensi() {
  const hasilTes = [];

  const skenario = [
    {
      nama: 'Tes sebelum jam pulang',
      jam: '10:00:00'
    },
    {
      nama: 'Tes tepat jam pulang',
      jam: '13:00:00'
    },
    {
      nama: 'Tes setelah jam pulang',
      jam: '13:30:00'
    }
  ];

  skenario.forEach(function(tes) {
    const hasil = simulasiAbsensiDenganWaktu_(
      'NIS:121235180009240002',
      tes.jam
    );

    hasilTes.push({
      skenario: tes.nama,
      jamTes: tes.jam,
      hasil: hasil
    });
  });

  Logger.log(JSON.stringify(hasilTes, null, 2));
  return hasilTes;
}


/**
 * Simulasi waktu untuk pengujian.
 * Tidak mengubah spreadsheet.
 */
function simulasiAbsensiDenganWaktu_(qrCode, jamTes) {
  const qr = String(qrCode || '').trim();
  const nis = qr.replace(/^NIS:/i, '').trim();

  if (!nis) {
    return {
      success: false,
      message: 'QR kosong.'
    };
  }

  const ss = getDatabase_();
  const siswaSheet = ss.getSheetByName(SHEET.SISWA);
  const absensiSheet = ss.getSheetByName(SHEET.ABSENSI_SISWA);

  if (!siswaSheet || !absensiSheet) {
    throw new Error('Sheet SISWA atau ABSENSI_SISWA tidak ditemukan.');
  }

  const siswaData = siswaSheet.getDataRange().getDisplayValues();
  const headerSiswa = siswaData[0].map(function(v) {
    return String(v).trim().toUpperCase();
  });

  const iNIS = headerSiswa.indexOf('NIS');
  const iNama = headerSiswa.indexOf('NAMA');
  const iKelas = headerSiswa.indexOf('KELAS');
  const iQR = headerSiswa.indexOf('KODE_QR');

  if ([iNIS, iNama, iKelas, iQR].includes(-1)) {
    throw new Error('Header SISWA tidak lengkap.');
  }

  let siswa = null;

  for (let i = 1; i < siswaData.length; i++) {
    if (
      String(siswaData[i][iNIS]).trim() === nis &&
      String(siswaData[i][iQR]).trim() === 'NIS:' + nis
    ) {
      siswa = {
        nis: nis,
        nama: siswaData[i][iNama],
        kelas: siswaData[i][iKelas]
      };
      break;
    }
  }

  if (!siswa) {
    return {
      success: false,
      message: 'QR tidak cocok dengan siswa.'
    };
  }

  const zona = Session.getScriptTimeZone();
  const sekarang = new Date();

  const tanggalHariIni = Utilities.formatDate(
    sekarang, zona, 'yyyy-MM-dd'
  );

  const absensiData = absensiSheet.getDataRange().getValues();
  const headerAbsensi = absensiData[0].map(function(v) {
    return String(v).trim().toUpperCase();
  });

  const iTanggal = headerAbsensi.indexOf('TANGGAL');
  const iNisAbsensi = headerAbsensi.indexOf('NIS');
  const iKegiatan = headerAbsensi.indexOf('KEGIATAN');

  if ([iTanggal, iNisAbsensi, iKegiatan].includes(-1)) {
    throw new Error('Header absensi tidak lengkap.');
  }

  const kegiatan = [];

  for (let i = 1; i < absensiData.length; i++) {
    const row = absensiData[i];

    if (String(row[iNisAbsensi]).trim() !== nis) continue;

    let tanggal = '';

    if (row[iTanggal] instanceof Date) {
      tanggal = Utilities.formatDate(
        row[iTanggal], zona, 'yyyy-MM-dd'
      );
    } else {
      tanggal = String(row[iTanggal] || '').trim().substring(0, 10);
    }

    if (tanggal === tanggalHariIni) {
      kegiatan.push(
        String(row[iKegiatan] || '').trim().toUpperCase()
      );
    }
  }

  const sudahMasuk = kegiatan.includes('MASUK');
  const sudahPulang = kegiatan.includes('PULANG');
  const menit = Number(String(jamTes).split(':')[0]) * 60 +
    Number(String(jamTes).split(':')[1]);

  let tindakan;
  let pesan;

  if (sudahMasuk && sudahPulang) {
    tindakan = 'DITOLAK';
    pesan = 'MASUK dan PULANG hari ini sudah tercatat.';
  } else if (!sudahMasuk && menit >= 13 * 60) {
    tindakan = 'DITOLAK';
    pesan = 'Tidak ada catatan MASUK. Periksa data atau kebijakan sekolah.';
  } else if (!sudahMasuk) {
    tindakan = 'MASUK';
    pesan = 'Siswa belum tercatat masuk hari ini.';
  } else if (menit < 13 * 60) {
    tindakan = 'DITOLAK';
    pesan = 'Belum jam pulang. Scan tidak disimpan.';
  } else {
    tindakan = 'PULANG';
    pesan = 'Jam pulang tercapai dan MASUK sudah tercatat.';
  }

  return {
    success: true,
    mode: 'SIMULASI_TIDAK_MENYIMPAN',
    nis: siswa.nis,
    nama: siswa.nama,
    kelas: siswa.kelas,
    jamTes: jamTes,
    absensiHariIni: kegiatan,
    tindakan: tindakan,
    message: pesan
  };
}

/**
 * TES LOGIKA MASUK-PULANG
 * Tidak membaca atau menulis spreadsheet.
 * Semua data tes hanya berada di memori.
 */
function tesUrutanMasukPulang() {
  const skenario = [
    {
      nama: '1. Belum ada absensi pukul 10.00',
      kegiatan: [],
      jam: '10:00:00',
      harapan: 'MASUK'
    },
    {
      nama: '2. Sudah MASUK, scan lagi pukul 10.00',
      kegiatan: ['MASUK'],
      jam: '10:00:00',
      harapan: 'DITOLAK'
    },
    {
      nama: '3. Sudah MASUK, scan pukul 13.00',
      kegiatan: ['MASUK'],
      jam: '13:00:00',
      harapan: 'PULANG'
    },
    {
      nama: '4. Sudah MASUK dan PULANG',
      kegiatan: ['MASUK', 'PULANG'],
      jam: '14:00:00',
      harapan: 'DITOLAK'
    }
  ];

  const hasilTes = skenario.map(function(tes) {
    const hasil = tentukanTindakanTes_(
      tes.kegiatan,
      tes.jam
    );

    return {
      skenario: tes.nama,
      jam: tes.jam,
      catatanSebelumnya: tes.kegiatan,
      tindakan: hasil.tindakan,
      harapan: tes.harapan,
      lulus: hasil.tindakan === tes.harapan,
      pesan: hasil.pesan
    };
  });

  Logger.log(JSON.stringify(hasilTes, null, 2));

  const semuaLulus = hasilTes.every(function(tes) {
    return tes.lulus;
  });

  Logger.log(
    semuaLulus
      ? 'SEMUA TES LULUS. Tidak ada data absensi yang disimpan.'
      : 'ADA TES YANG BELUM LULUS. Periksa hasil di atas.'
  );

  return hasilTes;
}


/**
 * Penentu tindakan khusus untuk tes.
 * Tidak menyimpan data dan tidak mengakses spreadsheet.
 */
function tentukanTindakanTes_(kegiatan, jam) {
  const daftar = kegiatan.map(function(item) {
    return String(item).trim().toUpperCase();
  });

  const bagianJam = String(jam).split(':');
  const menit = Number(bagianJam[0]) * 60 +
    Number(bagianJam[1]);

  const sudahMasuk = daftar.includes('MASUK');
  const sudahPulang = daftar.includes('PULANG');

  if (sudahMasuk && sudahPulang) {
    return {
      tindakan: 'DITOLAK',
      pesan: 'MASUK dan PULANG sudah tercatat.'
    };
  }

  if (sudahPulang && !sudahMasuk) {
    return {
      tindakan: 'DITOLAK',
      pesan: 'Catatan PULANG tanpa MASUK perlu diperiksa.'
    };
  }

  if (!sudahMasuk) {
    if (menit >= 13 * 60) {
      return {
        tindakan: 'DITOLAK',
        pesan: 'Tidak ada catatan MASUK.'
      };
    }

    return {
      tindakan: 'MASUK',
      pesan: 'Siswa belum tercatat masuk.'
    };
  }

  if (menit < 13 * 60) {
    return {
      tindakan: 'DITOLAK',
      pesan: 'Belum pukul 13.00. Scan tidak diterima.'
    };
  }

  return {
    tindakan: 'PULANG',
    pesan: 'Jam pulang tercapai dan MASUK sudah tercatat.'
  };
}

function tesKoneksiDatabase() {
  const hasil = ping();
  Logger.log(JSON.stringify(hasil, null, 2));
}

function tesCekDatabase() {
  const hasil = cekDatabase();
  Logger.log(JSON.stringify(hasil, null, 2));
}


/**
 * MODUL ABSENSI GURU - MTs PSM LOCERET
 * - MASUK dan PULANG dicatat pada baris berbeda.
 * - Simulasi aktif secara default.
 * - Tidak menerima absensi lebih dari satu MASUK dan satu PULANG per hari.
 * - Mencocokkan QR dengan KODE_QR terlebih dahulu.
 */
function simpanAbsensiGuruBaru_(kodeQR, opsi) {
  opsi = opsi || {};

  // Aman secara default: tidak menulis data kecuali diminta eksplisit.
  const simulasi = opsi.simulasi !== false;
  const qr = String(kodeQR || '').trim();

  if (!qr) {
    return { success: false, message: 'QR guru kosong.' };
  }

  const db = getDatabase_();
  const shGuru = db.getSheetByName('GURU');
  const shAbsensi = db.getSheetByName('ABSENSI_GURU');

  if (!shGuru || !shAbsensi) {
    return {
      success: false,
      message: 'Sheet GURU atau ABSENSI_GURU tidak ditemukan.'
    };
  }

  const dataGuru = shGuru.getDataRange().getDisplayValues();
  const dataAbsensi = shAbsensi.getDataRange().getDisplayValues();

  if (dataGuru.length < 2) {
    return { success: false, message: 'Data guru belum tersedia.' };
  }

  const normalisasiHeader = row =>
    row.map(v => String(v || '').trim().toUpperCase());

  const headerGuru = normalisasiHeader(dataGuru[0]);
  const headerAbsensi = normalisasiHeader(dataAbsensi[0]);

  const idxId = headerGuru.indexOf('ID_GURU');
  const idxNama = headerGuru.indexOf('NAMA');
  const idxQR = headerGuru.indexOf('KODE_QR');
  const idxStatusGuru = headerGuru.indexOf('STATUS');

  const idxTimestamp = headerAbsensi.indexOf('TIMESTAMP');
  const idxTanggal = headerAbsensi.indexOf('TANGGAL');
  const idxJam = headerAbsensi.indexOf('JAM');
  const idxIdAbsensi = headerAbsensi.indexOf('ID_GURU');
  const idxNamaAbsensi = headerAbsensi.indexOf('NAMA_GURU');
  const idxStatusAbsensi = headerAbsensi.indexOf('STATUS');
  const idxKeterangan = headerAbsensi.indexOf('KETERANGAN');

  if (
    [idxId, idxNama, idxQR].includes(-1) ||
    [
      idxTimestamp, idxTanggal, idxJam, idxIdAbsensi,
      idxNamaAbsensi, idxStatusAbsensi, idxKeterangan
    ].includes(-1)
  ) {
    return {
      success: false,
      message: 'Header wajib belum lengkap. Tidak ada data yang ditulis.'
    };
  }

  // Prioritas 1: cocokkan isi QR persis dengan KODE_QR.
  // Alternatif ID hanya digunakan jika KODE_QR pada baris guru kosong.
  const qrTanpaPrefix = qr
    .replace(/^(ID_GURU|GURU):/i, '')
    .trim();

  const guru = dataGuru.slice(1).find(row => {
    const id = String(row[idxId] || '').trim();
    const kodeTersimpan = String(row[idxQR] || '').trim();

    if (kodeTersimpan) {
      return kodeTersimpan === qr;
    }

    return id !== '' && (
      qr === id ||
      qrTanpaPrefix === id
    );
  });

  if (!guru) {
    return {
      success: false,
      message: 'QR belum cocok dengan data guru. Periksa KODE_QR dan QR lama.'
    };
  }

  const statusGuru = idxStatusGuru === -1
    ? ''
    : String(guru[idxStatusGuru] || '').trim().toUpperCase();

  if (statusGuru && !['AKTIF', 'ACTIVE'].includes(statusGuru)) {
    return { success: false, message: 'Status guru tidak aktif.' };
  }

  const idGuru = String(guru[idxId] || '').trim();
  const namaGuru = String(guru[idxNama] || '').trim();

  if (!idGuru || !namaGuru) {
    return {
      success: false,
      message: 'ID_GURU atau nama guru kosong.'
    };
  }

  const zona = Session.getScriptTimeZone();
  const sekarang = new Date();
  const tanggal = Utilities.formatDate(sekarang, zona, 'yyyy-MM-dd');
  const jam = Utilities.formatDate(sekarang, zona, 'HH:mm:ss');

  // Menyamakan format tanggal agar catatan lama tidak mudah terlewat.
  function tanggalNormal(nilai) {
    const teks = String(nilai || '').trim();
    if (!teks) return '';

    if (/^\d{4}-\d{2}-\d{2}$/.test(teks)) return teks;

    const cocok = teks.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if (cocok) {
      // Format tampilan dd/MM/yyyy atau dd-MM-yyyy.
      return cocok[3] + '-' +
        String(cocok[2]).padStart(2, '0') + '-' +
        String(cocok[1]).padStart(2, '0');
    }

    return teks;
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    // Baca ulang setelah mendapat lock untuk mencegah scan bersamaan
    // menghasilkan dua catatan MASUK atau dua catatan PULANG.
    const riwayat = shAbsensi.getDataRange().getDisplayValues();
    const barisHariIni = riwayat.slice(1).filter(row =>
      String(row[idxIdAbsensi] || '').trim() === idGuru &&
      tanggalNormal(row[idxTanggal]) === tanggal
    );

    const sudahMasuk = barisHariIni.some(row =>
      String(row[idxKeterangan] || '').trim().toUpperCase() === 'MASUK'
    );

    const sudahPulang = barisHariIni.some(row =>
      String(row[idxKeterangan] || '').trim().toUpperCase() === 'PULANG'
    );

    let kegiatan;

    if (sudahPulang && !sudahMasuk) {
      return {
        success: false,
        message: 'Ditemukan catatan PULANG tanpa MASUK. Periksa data absensi.'
      };
    }

    if (!sudahMasuk) {
      kegiatan = 'MASUK';
    } else if (!sudahPulang) {
      kegiatan = 'PULANG';
    } else {
      return {
        success: false,
        message: 'Absensi MASUK dan PULANG hari ini sudah lengkap.'
      };
    }

    const hasil = {
      success: true,
      simulasi: simulasi,
      idGuru: idGuru,
      nama: namaGuru,
      tanggal: tanggal,
      jam: jam,
      kegiatan: kegiatan,
      message: simulasi
        ? 'Simulasi berhasil. Spreadsheet belum diubah.'
        : 'Absensi guru berhasil disimpan.'
    };

    if (simulasi) return hasil;

    const barisBaru = Array(headerAbsensi.length).fill('');
    barisBaru[idxTimestamp] = sekarang;
    barisBaru[idxTanggal] = tanggal;
    barisBaru[idxJam] = jam;
    barisBaru[idxIdAbsensi] = idGuru;
    barisBaru[idxNamaAbsensi] = namaGuru;
    barisBaru[idxStatusAbsensi] = 'HADIR';
    barisBaru[idxKeterangan] = kegiatan;

    shAbsensi.appendRow(barisBaru);
    SpreadsheetApp.flush();

    return hasil;

  } catch (err) {
    return {
      success: false,
      message: 'Absensi guru gagal diproses: ' + err.message
    };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function tesSimulasiAbsensiGuru() {
const qrUji = 'SITI MAEMUNAH';

  const hasil = simpanAbsensiGuruBaru_(qrUji, {
    simulasi: true
  });

  Logger.log(JSON.stringify(hasil, null, 2));
}

/**
 * Pemeriksaan QR guru untuk integrasi Scanner.html.
 * Aman untuk pengujian: tidak menulis absensi.
 */
function periksaQRGuru(kodeQR) {
  const qr = String(kodeQR || '').trim();

  if (!qr) {
    return {
      success: false,
      message: 'QR kosong.'
    };
  }

  const db = getDatabase_();
  const sh = db.getSheetByName('GURU');

  if (!sh) {
    return {
      success: false,
      message: 'Sheet GURU tidak ditemukan.'
    };
  }

  const data = sh.getDataRange().getDisplayValues();

  if (data.length < 2) {
    return {
      success: false,
      message: 'Data guru belum tersedia.'
    };
  }

  const header = data[0].map(v =>
    String(v || '').trim().toUpperCase()
  );

  const iId = header.indexOf('ID_GURU');
  const iNama = header.indexOf('NAMA');
  const iQR = header.indexOf('KODE_QR');
  const iStatus = header.indexOf('STATUS');

  if ([iId, iNama, iQR].includes(-1)) {
    return {
      success: false,
      message: 'Header ID_GURU, NAMA, atau KODE_QR tidak ditemukan.'
    };
  }

  // Normalisasi hanya untuk membandingkan spasi dan huruf besar/kecil.
  const norm = v => String(v || '')
    .trim()
    .replace(/\s+/g, '')
    .toUpperCase();

  const ditemukan = data.slice(1).find(row => {
    const id = String(row[iId] || '').trim();
    const nama = String(row[iNama] || '').trim();
    const kode = String(row[iQR] || '').trim();

    if (!id || !nama) return false;

    // QR tersimpan menjadi acuan utama.
    if (kode) return norm(kode) === norm(qr);

    // Jika KODE_QR kosong, izinkan ID persis sebagai pemeriksaan.
    return norm(id) === norm(qr);
  });

  if (!ditemukan) {
    return {
      success: false,
      message: 'QR belum cocok dengan data guru.'
    };
  }

  if (iStatus !== -1) {
    const status = String(ditemukan[iStatus] || '')
      .trim()
      .toUpperCase();

    if (status && !['AKTIF', 'ACTIVE'].includes(status)) {
      return {
        success: false,
        message: 'Guru ditemukan, tetapi statusnya tidak aktif.'
      };
    }
  }

  return {
    success: true,
    idGuru: ditemukan[iId],
    nama: ditemukan[iNama],
    message: 'QR guru cocok. Belum ada absensi yang disimpan.'
  };
}

function tesPeriksaQRGuru() {
  const hasil = periksaQRGuru('SITIMAEMUNAH');
  Logger.log(JSON.stringify(hasil, null, 2));
}








/***************************************************************
 * MODUL PRODUKSI SCANNER TERPADU — VERSI KONSOLIDASI
 * Override fungsi lama dilakukan di akhir file agar modul lama
 * (dashboard, rekap, data master) tetap dipertahankan.
 ***************************************************************/

/** Ambil nilai pengaturan dari sheet baru PENGATURAN (KUNCI | NILAI). */
function bacaPengaturanTerpadu_(kunci, nilaiDefault) {
  const ss = getDatabase_();
  const sh = ss.getSheetByName('PENGATURAN');
  if (!sh || sh.getLastRow() < 2) return nilaiDefault;
  const rows = sh.getDataRange().getDisplayValues();
  const target = String(kunci || '').trim().toUpperCase();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim().toUpperCase() === target) {
      const v = String(rows[i][1] || '').trim();
      return v === '' ? nilaiDefault : v;
    }
  }
  return nilaiDefault;
}

/** Validasi GPS siswa terhadap PENGATURAN; default radius 150 meter. */
function validasiLokasiSekolah_(latitude, longitude) {
  const latSekolah = Number(bacaPengaturanTerpadu_('LATITUDE_SEKOLAH', -7.671100));
  const lonSekolah = Number(bacaPengaturanTerpadu_('LONGITUDE_SEKOLAH', 111.894600));
  const radius = Number(bacaPengaturanTerpadu_('RADIUS_SEKOLAH_METER', 150));
  const lat = Number(latitude), lon = Number(longitude);
  if (!Number.isFinite(lat) || !Number.isFinite(lon) || (lat === 0 && lon === 0)) {
    return {valid:false, message:'Lokasi GPS belum terbaca. Aktifkan lokasi dan izinkan akses browser.'};
  }
  if (!Number.isFinite(latSekolah) || !Number.isFinite(lonSekolah) || !Number.isFinite(radius) || radius <= 0) {
    return {valid:false, message:'Pengaturan lokasi sekolah belum valid di sheet PENGATURAN.'};
  }
  const jarak = hitungJarakMeter_(lat, lon, latSekolah, lonSekolah);
  return jarak <= radius
    ? {valid:true, jarak:Math.round(jarak), message:'Lokasi sesuai lingkungan sekolah.'}
    : {valid:false, jarak:Math.round(jarak), message:'Di luar area sekolah ('+Math.round(jarak)+' m). Radius yang diizinkan '+radius+' m.'};
}

/** Endpoint tunggal siswa: GPS -> validasi QR -> simpan satu kali dengan lock. */
function prosesScanSiswaTerpadu(qrCode, latitude, longitude) {
  try {
    const lokasi = validasiLokasiSekolah_(latitude, longitude);
    if (!lokasi.valid) return {success:false, message:lokasi.message, jarak:lokasi.jarak};
    const hasil = simpanAbsensiSiswaBaru_(qrCode, {simulasi:false});
    if (hasil && hasil.success) {
      hasil.jarak = lokasi.jarak;
      hasil.message = (hasil.message || 'Absensi tersimpan.') + ' Jarak dari titik sekolah: '+lokasi.jarak+' m.';
    }
    return hasil;
  } catch (e) {
    return {success:false, message:'Absensi siswa gagal diproses: '+(e && e.message ? e.message : e)};
  }
}

/** Simpan absensi siswa dengan QR lama NIS:<NIS>, tanpa mengubah QR yang telah dicetak. */
function simpanAbsensiSiswaBaru_(qrCode, opsi) {
  opsi = opsi || {};
  const simulasi = opsi.simulasi !== false;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return {success:false, message:'Scanner sedang memproses scan lain. Coba lagi sebentar.'};
  try {
    const qr = String(qrCode || '').trim();
    const nis = qr.replace(/^NIS:/i, '').trim();
    if (!nis || !/^[0-9A-Za-z._-]+$/.test(nis)) return {success:false, message:'Format QR siswa tidak valid. QR lama NIS:<NIS> tetap didukung.'};
    const ss = getDatabase_();
    const shSiswa = ss.getSheetByName('SISWA');
    const shAbs = ss.getSheetByName('ABSENSI_SISWA');
    if (!shSiswa || !shAbs) throw new Error('Sheet SISWA atau ABSENSI_SISWA tidak ditemukan.');
    const siswaData = shSiswa.getDataRange().getDisplayValues();
    if (siswaData.length < 2) return {success:false, message:'Data siswa belum tersedia.'};
    const hs = siswaData[0].map(v=>String(v||'').trim().toUpperCase());
    const ix = {nis:hs.indexOf('NIS'), nama:hs.indexOf('NAMA'), kelas:hs.indexOf('KELAS'), jk:hs.indexOf('JK'), qr:hs.indexOf('KODE_QR'), status:hs.indexOf('STATUS')};
    if ([ix.nis,ix.nama,ix.kelas,ix.qr].some(i=>i<0)) throw new Error('Header SISWA wajib: NIS, NAMA, KELAS, KODE_QR.');
    let siswa = null;
    for (let i=1;i<siswaData.length;i++) {
      const row=siswaData[i], rowNis=String(row[ix.nis]||'').trim(), rowQr=String(row[ix.qr]||'').trim();
      if (rowNis===nis && (!rowQr || rowQr.toUpperCase()===('NIS:'+nis).toUpperCase() || rowQr.toUpperCase()===qr.toUpperCase())) {
        if (ix.status>=0) { const st=String(row[ix.status]||'').trim().toUpperCase(); if (st && !['AKTIF','ACTIVE','YA'].includes(st)) return {success:false,message:'Status siswa tidak aktif.'}; }
        siswa={nis:rowNis,nama:row[ix.nama],kelas:row[ix.kelas],jk:ix.jk>=0?row[ix.jk]:''}; break;
      }
    }
    if (!siswa) return {success:false,message:'QR tidak cocok dengan NIS dan KODE_QR pada database siswa.'};
    const ah=shAbs.getRange(1,1,1,shAbs.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim().toUpperCase());
    const required=['TIMESTAMP','TANGGAL','JAM','NIS','NAMA','KELAS','KEGIATAN'];
    const ixA={}; required.forEach(k=>ixA[k]=ah.indexOf(k));
    if (required.some(k=>ixA[k]<0)) throw new Error('Header ABSENSI_SISWA belum lengkap: '+required.join(', '));
    const now=new Date(), tz='Asia/Jakarta';
    const today=Utilities.formatDate(now,tz,'yyyy-MM-dd');
    const hhmm=Utilities.formatDate(now,tz,'HH:mm');
    const rows=shAbs.getDataRange().getValues();
    const todayRows=[];
    for(let i=1;i<rows.length;i++){
      const r=rows[i]; if(String(r[ixA.NIS]||'').trim()!==nis) continue;
      let d=''; const raw=r[ixA.TANGGAL];
      if(raw instanceof Date && !isNaN(raw)) d=Utilities.formatDate(raw,tz,'yyyy-MM-dd');
      else { const s=String(raw||'').trim(); if(/^\d{4}-\d{2}-\d{2}/.test(s)) d=s.slice(0,10); else { const m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/); if(m)d=m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0'); } }
      if(d===today) todayRows.push({row:r, kegiatan:String(r[ixA.KEGIATAN]||'').trim().toUpperCase()});
    }
    const masuk=todayRows.find(x=>x.kegiatan==='MASUK');
    const pulang=todayRows.find(x=>x.kegiatan==='PULANG');
    if(pulang && !masuk) return {success:false,message:'Data hari ini tidak berurutan: PULANG tercatat tanpa MASUK. Hubungi admin.'};
    if(masuk && pulang) return {success:false,message:'Absensi MASUK dan PULANG hari ini sudah lengkap.'};
    let kegiatan;
    if(!masuk){
      if(hhmm<'07:00') return {success:false,message:'Absensi MASUK dibuka mulai pukul 07.00 WIB.'};
      if(hhmm>='13:00') return {success:false,message:'Belum ada catatan MASUK hari ini. Absensi tidak bisa langsung dicatat sebagai PULANG setelah pukul 13.00 WIB. Silakan hubungi petugas absensi.'};
      kegiatan='MASUK';
    } else {
      if(hhmm<'13:00') return {success:false,message:'Absensi PULANG baru dibuka mulai pukul 13.00 WIB. Absensi MASUK sudah tercatat.'};
      kegiatan='PULANG';
    }
    const jam=Utilities.formatDate(now,tz,'HH:mm:ss');
    if(simulasi) return {success:true,simulasi:true,nis:siswa.nis,nama:siswa.nama,kelas:siswa.kelas,kegiatan:kegiatan,tanggal:today,jam:jam,message:'Simulasi saja; spreadsheet tidak diubah.'};
    const newRow=Array(shAbs.getLastColumn()).fill('');
    const day=new Date(today+'T00:00:00+07:00');
    const values={TIMESTAMP:now,TANGGAL:day,JAM:now,NIS:siswa.nis,NAMA:siswa.nama,KELAS:siswa.kelas,JK:siswa.jk,KEGIATAN:kegiatan,JAM_KE:'',STATUS:'HADIR',KETERANGAN:'QR'};
    Object.keys(values).forEach(k=>{const i=ah.indexOf(k);if(i>=0)newRow[i]=values[k];});
    shAbs.appendRow(newRow); SpreadsheetApp.flush();
    return {success:true,simulasi:false,nis:siswa.nis,nama:siswa.nama,kelas:siswa.kelas,kegiatan:kegiatan,tanggal:today,jam:jam,jamMasuk:jam,status:'HADIR',message:'Absensi '+kegiatan+' berhasil disimpan.'};
  } catch(e) { return {success:false,message:'Gagal menyimpan absensi siswa: '+(e&&e.message?e.message:e)}; }
  finally { if(lock.hasLock()) lock.releaseLock(); }
}

/** Endpoint guru: simpan MASUK/PULANG langsung; tidak sekadar memeriksa QR. */
function prosesScanGuruTerpadu(kodeQR) {
  return simpanAbsensiGuruBaru_(kodeQR, {simulasi:false});
}

/** Override agar QR guru dibandingkan normalisasi spasi/huruf, dengan KODE_QR sebagai acuan utama. */
function simpanAbsensiGuruBaru_(kodeQR, opsi) {
  opsi=opsi||{}; const simulasi=opsi.simulasi!==false; const qr=String(kodeQR||'').trim();
  if(!qr)return {success:false,message:'QR guru kosong.'};
  const ss=getDatabase_(), shG=ss.getSheetByName('GURU'), shA=ss.getSheetByName('ABSENSI_GURU');
  if(!shG||!shA)return {success:false,message:'Sheet GURU atau ABSENSI_GURU tidak ditemukan.'};
  const gd=shG.getDataRange().getDisplayValues(), ad=shA.getDataRange().getDisplayValues();
  if(gd.length<2)return {success:false,message:'Data guru belum tersedia.'};
  const gh=gd[0].map(x=>String(x||'').trim().toUpperCase()), ah=ad[0].map(x=>String(x||'').trim().toUpperCase());
  const idI=gh.indexOf('ID_GURU'), namaI=gh.indexOf('NAMA'), qrI=gh.indexOf('KODE_QR'), stI=gh.indexOf('STATUS');
  const must=['TIMESTAMP','TANGGAL','JAM','ID_GURU','NAMA_GURU','STATUS','KETERANGAN'];
  if([idI,namaI,qrI].some(i=>i<0)||must.some(k=>ah.indexOf(k)<0))return {success:false,message:'Header GURU/ABSENSI_GURU belum lengkap.'};
  const norm=x=>String(x||'').trim().replace(/\s+/g,'').toUpperCase();
  const guru=gd.slice(1).find(r=>{const stored=String(r[qrI]||'').trim();return stored?norm(stored)===norm(qr):norm(r[idI])===norm(qr.replace(/^(ID_GURU|GURU):/i,''));});
  if(!guru)return {success:false,message:'QR guru tidak cocok dengan KODE_QR. QR lama tetap harus sama dengan kode pada kartu.'};
  if(stI>=0){const st=String(guru[stI]||'').trim().toUpperCase();if(st&&!['AKTIF','ACTIVE','YA'].includes(st))return {success:false,message:'Status guru tidak aktif.'};}
  const id=String(guru[idI]||'').trim(), nama=String(guru[namaI]||'').trim();
  const lock=LockService.getScriptLock(); if(!lock.tryLock(10000))return {success:false,message:'Scanner sedang memproses scan lain. Coba lagi.'};
  try{
    const now=new Date(), tz='Asia/Jakarta', today=Utilities.formatDate(now,tz,'yyyy-MM-dd'), jam=Utilities.formatDate(now,tz,'HH:mm:ss');
    const ix={};must.forEach(k=>ix[k]=ah.indexOf(k));
    const current=shA.getDataRange().getValues();
    const riwayat=current.slice(1).filter(r=>String(r[ix.ID_GURU]||'').trim()===id&&tanggalNormalTerpadu_(r[ix.TANGGAL],tz)===today);
    const masuk=riwayat.some(r=>String(r[ix.KETERANGAN]||'').trim().toUpperCase()==='MASUK');
    const pulang=riwayat.some(r=>String(r[ix.KETERANGAN]||'').trim().toUpperCase()==='PULANG');
    if(pulang&&!masuk)return {success:false,message:'Catatan PULANG ditemukan tanpa MASUK. Periksa data.'};
    let kegiatan=!masuk?'MASUK':(!pulang?'PULANG':'SELESAI');
    if(kegiatan==='SELESAI')return {success:false,message:'Absensi MASUK dan PULANG guru hari ini sudah lengkap.'};
    const result={success:true,simulasi:simulasi,idGuru:id,nama:nama,tanggal:today,jam:jam,kegiatan:kegiatan,message:simulasi?'Simulasi saja; tidak ada data yang ditulis.':'Absensi guru '+kegiatan+' berhasil disimpan.'};
    if(simulasi)return result;
    const newRow=Array(shA.getLastColumn()).fill('');
    const values={TIMESTAMP:now,TANGGAL:new Date(today+'T00:00:00+07:00'),JAM:now,ID_GURU:id,NAMA_GURU:nama,JAM_KE:'',KELAS:'',MAPEL:'',STATUS:'HADIR',KETERANGAN:kegiatan};
    Object.keys(values).forEach(k=>{const i=ah.indexOf(k);if(i>=0)newRow[i]=values[k];});
    shA.appendRow(newRow);SpreadsheetApp.flush();return result;
  }catch(e){return {success:false,message:'Absensi guru gagal: '+(e&&e.message?e.message:e)};}
  finally{if(lock.hasLock())lock.releaseLock();}
}
function tanggalNormalTerpadu_(v,tz){
  if(v instanceof Date&&!isNaN(v))return Utilities.formatDate(v,tz,'yyyy-MM-dd');
  const s=String(v||'').trim();if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);
  const m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})$/);return m?m[3]+'-'+m[2].padStart(2,'0')+'-'+m[1].padStart(2,'0'):s;
}
