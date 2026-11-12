/**
 * ============================================================================
 * SIKUTE (Sistem Kunjungan Terintegrasi Pelayanan Kesehatan)
 * Balai Kekarantinaan Kesehatan (BKK) Kelas I Mataram
 * Arsitektur Database Relasional Terpisah (Multi-Sheet Schema) & Zero-Trust RBAC
 * ============================================================================
 */

const APP_TITLE = 'SIKUTE - Sistem Kunjungan Terintegrasi Pelayanan Kesehatan (BKK Mataram)';
const SESSION_TTL_SECONDS = 21600; // 6 Jam (6 * 3600 detik)
const CACHE_REF_TTL_SECONDS = 600; // 10 Menit (600 detik)

const WILKER_LIST = [
  'Induk', 'Pemenang', 'Lembar', 'Bizam', 'Lab Lombok',
  'Pototano', 'Benete', 'Badas', 'Bima', 'Sape'
];

const WILKER_CODES = {
  'Induk': 'IND',
  'Pemenang': 'PMN',
  'Lembar': 'LBR',
  'Bizam': 'BZM',
  'Lab Lombok': 'LAB',
  'Pototano': 'PTT',
  'Benete': 'BNT',
  'Badas': 'BDS',
  'Bima': 'BMA',
  'Sape': 'SPE'
};

const SHEET_NAMES = {
  KUNJUNGAN: 'Data_Kunjungan',
  LAB: 'Data_Laboratorium',
  OBAT: 'Data_Terapi_Obat',
  REF_OBAT: 'Obat.ref',
  REF_KOTAKAB: 'DataKotakab',
  REF_ICD10: 'ICD-10',
  REF_DXKEP: 'DXKeperawatan',
  USERS: 'Users',
  LOG_ARSIP: 'Log_Arsip'
};

/**
 * Render halaman utama SPA Web Application
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * ============================================================================
 * 1. INISIALISASI & SETUP DATABASE RELASIONAL (9 SHEETS)
 * ============================================================================
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Data_Kunjungan (54 Kolom - termasuk Lokasi_Rujukan, Diagnosa_Kerja_E & Diagnosa_Kerja)
  const headersKunjungan = [
    'ID_Kunjungan', 'Timestamp', 'Tanggal_Kunjungan', 'Wilker', 'Petugas_Input',
    'Mode_RM', 'No_RM', 'No_Identitas', 'Jenis_Kunjungan', 'Sub_Kunjungan',
    'Fase_Haji', 'Kloter_Haji', 'Vaksin_Dipilih', 'Tarif_Pemeriksaan', 'Tarif_ICV',
    'Tarif_Kontraindikasi', 'Tarif_Vaksin', 'Total_PNBP', 'Nama_Pasien', 'JK',
    'Umur', 'Kategori_Umur', 'Alamat', 'Provinsi', 'KabKota',
    'Pekerjaan', 'Keluhan', 'Sistolik', 'Diastolik', 'Kategori_TD',
    'Respirasi', 'Nadi', 'SpO2', 'Suhu', 'TB',
    'BB', 'IMT', 'Kategori_IMT', 'Lingkar_Perut', 'Buta_Warna',
    'ICD10_Code', 'Diagnosa_ICD10', 'Group_Diag', 'Kategori_DXKep', 'DX_Keperawatan',
    'FR_Ditemukan', 'FR_Dikendalikan', 'Kategori_Penumpang', 'Dokter_Pemeriksa', 'Ada_Lab',
    'Ada_Obat', 'Lokasi_Rujukan', 'Diagnosa_Kerja_E', 'Diagnosa_Kerja'
  ];
  ensureSheetWithHeaders(ss, SHEET_NAMES.KUNJUNGAN, headersKunjungan, '#0d6efd');

  // 2. Sheet Data_Laboratorium (47 Kolom - termasuk Interp_Asam_Urat & Jenis_Dengue)
  const headersLab = [
    'ID_Lab', 'ID_Kunjungan', 'Tanggal_Kunjungan', 'Wilker', 'No_RM',
    'Nama_Pasien', 'Hemoglobin', 'Hematokrit', 'Leukosit', 'Platelet',
    'Eritrosit', 'GDP', 'Interp_GDP', 'GDS', 'Interp_GDS',
    'GD_2JPP', 'Interp_GD_2JPP', 'Kolesterol_Total', 'Interp_Kol_Total', 'LDL',
    'Interp_LDL', 'HDL', 'Interp_HDL', 'Trigliserida', 'Interp_Trigliserida',
    'Asam_Urat', 'Warna_Urine', 'Bau_Urine', 'Kejernihan_Urine', 'pH_Urine',
    'Protein_Urine', 'Glukosa_Urine', 'Keton_Urine', 'Bilirubin_Urine', 'Urobilinogen_Urine',
    'Darah_Samar_Urine', 'PPTest_Kehamilan', 'Malaria_Test', 'Dengue_Test', 'HIV_Test',
    'Widal_Test', 'RDT_Covid19', 'TB_Test', 'EKG', 'Status_Abnormal_Positif',
    'Interp_Asam_Urat', 'Jenis_Dengue'
  ];
  ensureSheetWithHeaders(ss, SHEET_NAMES.LAB, headersLab, '#198754');

  // 3. Sheet Data_Terapi_Obat (10 Kolom - Aturan_Pakai telah dihapus sesuai spesifikasi)
  const headersObat = [
    'ID_Terapi', 'ID_Kunjungan', 'Tanggal_Kunjungan', 'Wilker', 'No_RM',
    'Nama_Pasien', 'Nama_Obat', 'Sediaan', 'Jumlah', 'Dokter_Pemeriksa'
  ];
  ensureSheetWithHeaders(ss, SHEET_NAMES.OBAT, headersObat, '#0dcaf0');

  // 4. Sheet Obat.ref (Kolom A = no, Kolom B = Nama Obat, Kolom C = Sediaan)
  const shObatRef = ensureSheetWithHeaders(ss, SHEET_NAMES.REF_OBAT, ['no', 'Nama Obat', 'Sediaan'], '#6f42c1');
  if (shObatRef.getLastRow() === 1) {
    const sampleObat = [
      [1, 'Paracetamol 500 mg', 'Tablet'],
      [2, 'Amoxicillin 500 mg', 'Kaplet'],
      [3, 'Cetirizine 10 mg', 'Tablet'],
      [4, 'Omeprazole 20 mg', 'Kapsul'],
      [5, 'Antasida Doen', 'Tablet Kunyah'],
      [6, 'Ibuprofen 400 mg', 'Tablet'],
      [7, 'Amlodipine 5 mg', 'Tablet'],
      [8, 'Amlodipine 10 mg', 'Tablet'],
      [9, 'Metformin 500 mg', 'Tablet'],
      [10, 'Simvastatin 10 mg', 'Tablet'],
      [11, 'Allopurinol 100 mg', 'Tablet'],
      [12, 'Salbutamol 2 mg', 'Tablet'],
      [13, 'Ambroxol 30 mg', 'Tablet'],
      [14, 'Ambroxol Sirup 15mg/5ml', 'Sirup'],
      [15, 'Oralit Sachet', 'Serbuk'],
      [16, 'Domperidone 10 mg', 'Tablet'],
      [17, 'Loperamide 2 mg', 'Tablet'],
      [18, 'Hydrocortisone Cream 2.5%', 'Salep/Krim'],
      [19, 'Gentamicin Salep Kulit', 'Salep/Krim'],
      [20, 'Vitamin B Complex', 'Tablet'],
      [21, 'Vitamin C 500 mg', 'Tablet'],
      [22, 'Epinephrine 1 mg/ml Injeksi', 'Ampul'],
      [23, 'Dexamethasone 5 mg/ml Injeksi', 'Ampul'],
      [24, 'Vaksin Meningitis Menactra/Menivax', 'Vial'],
      [25, 'Vaksin Influenza Vaxigrip', 'Prefilled Syringe'],
      [26, 'Vaksin Yellow Fever Stamaril', 'Vial'],
      [27, 'Vaksin Typhoid Typhim Vi', 'Prefilled Syringe'],
      [28, 'Vaksin Polio IPV', 'Vial']
    ];
    shObatRef.getRange(2, 1, sampleObat.length, 3).setValues(sampleObat);
  }

  // 5. Sheet DataKotakab (Kolom A = No, Kolom B = Provinsi, Kolom C = Kabupaten/Kota)
  const shKota = ensureSheetWithHeaders(ss, SHEET_NAMES.REF_KOTAKAB, ['No', 'Provinsi', 'Kabupaten/Kota'], '#fd7e14');
  if (shKota.getLastRow() === 1) {
    const sampleKota = [
      [1, 'Nusa Tenggara Barat', 'Kota Mataram'],
      [2, 'Nusa Tenggara Barat', 'Kabupaten Lombok Barat'],
      [3, 'Nusa Tenggara Barat', 'Kabupaten Lombok Tengah'],
      [4, 'Nusa Tenggara Barat', 'Kabupaten Lombok Timur'],
      [5, 'Nusa Tenggara Barat', 'Kabupaten Lombok Utara'],
      [6, 'Nusa Tenggara Barat', 'Kabupaten Sumbawa Barat'],
      [7, 'Nusa Tenggara Barat', 'Kabupaten Sumbawa'],
      [8, 'Nusa Tenggara Barat', 'Kabupaten Dompu'],
      [9, 'Nusa Tenggara Barat', 'Kabupaten Bima'],
      [10, 'Nusa Tenggara Barat', 'Kota Bima'],
      [11, 'Bali', 'Kota Denpasar'],
      [12, 'Bali', 'Kabupaten Badung'],
      [13, 'Bali', 'Kabupaten Karangasem'],
      [14, 'Jawa Timur', 'Kota Surabaya'],
      [15, 'Jawa Timur', 'Kabupaten Sidoarjo'],
      [16, 'Jawa Timur', 'Kabupaten Banyuwangi'],
      [17, 'DKI Jakarta', 'Kota Jakarta Pusat'],
      [18, 'DKI Jakarta', 'Kota Jakarta Selatan'],
      [19, 'Jawa Barat', 'Kota Bandung'],
      [20, 'Sulawesi Selatan', 'Kota Makassar']
    ];
    shKota.getRange(2, 1, sampleKota.length, 3).setValues(sampleKota);
  }

  // 6. Sheet ICD-10 (Kolom A = No, Kolom B = ICD10_Code, Kolom C = Diagnosa ICD 10, Kolom D = Group_Diag, Kolom E = Kategori_Diagnosa_Kerja, Kolom F = Diagnosa_Kerja)
  const shICD = ensureSheetWithHeaders(ss, SHEET_NAMES.REF_ICD10, ['No', 'ICD10_Code', 'Diagnosa ICD 10', 'Group_Diag', 'Kategori_Diagnosa_Kerja', 'Diagnosa_Kerja'], '#20c997');
  if (shICD.getLastRow() === 1) {
    const sampleICD = [
      [1, 'Z00', 'General examination and investigation of persons without complaint and reported diagnosis', 'Factors influencing health status', 'Pemeriksaan Kesehatan Umum', 'Pemeriksaan Kesehatan / Fit to Travel'],
      [2, 'Z00.0', 'General medical examination (Kier Kesehatan)', 'Factors influencing health status', 'Pemeriksaan Kesehatan Umum', 'Kier Kesehatan Pelayaran / Penerbangan'],
      [3, 'Z23', 'Need for immunization against single bacterial diseases', 'Immunization and Vaccination', 'Vaksinasi & Imunisasi', 'Vaksinasi Meningitis / Perjalanan Internasional'],
      [4, 'Z24', 'Need for immunization against certain single viral diseases', 'Immunization and Vaccination', 'Vaksinasi & Imunisasi', 'Vaksinasi Yellow Fever / Polio / Influenza'],
      [5, 'Z26.9', 'Need for immunization against unspecified infectious disease', 'Immunization and Vaccination', 'Vaksinasi & Imunisasi', 'Imunisasi Penyakit Menular Lainnya'],
      [6, 'J00', 'Acute nasopharyngitis [common cold]', 'Diseases of the respiratory system', 'Penyakit Saluran Pernapasan', 'Common Cold / Nasofaringitis Akut'],
      [7, 'J06.9', 'Acute upper respiratory infection, unspecified (ISPA)', 'Diseases of the respiratory system', 'Penyakit Saluran Pernapasan', 'ISPA (Infeksi Saluran Pernapasan Akut)'],
      [8, 'I10', 'Essential (primary) hypertension', 'Diseases of the circulatory system', 'Penyakit Kardiovaskular', 'Hipertensi Esensial (Primer)'],
      [9, 'E11.9', 'Type 2 diabetes mellitus without complications', 'Endocrine, nutritional and metabolic diseases', 'Penyakit Metabolik & Endokrin', 'Diabetes Melitus Tipe 2'],
      [10, 'E78.0', 'Pure hypercholesterolaemia', 'Endocrine, nutritional and metabolic diseases', 'Penyakit Metabolik & Endokrin', 'Hiperkolesterolemia Murni'],
      [11, 'E79.0', 'Hyperuricaemia without signs of inflammatory arthritis and tophaceous disease', 'Endocrine, nutritional and metabolic diseases', 'Penyakit Metabolik & Endokrin', 'Hiperurisemia (Asam Urat Tinggi)'],
      [12, 'K30', 'Functional dyspepsia', 'Diseases of the digestive system', 'Penyakit Saluran Pencernaan', 'Dispepsia Fungsional / Gastritis'],
      [13, 'A09', 'Infectious gastroenteritis and colitis, unspecified (Diare Akut)', 'Certain infectious and parasitic diseases', 'Penyakit Infeksi & Parasit', 'Gastroenteritis / Diare Akut'],
      [14, 'A15.0', 'Tuberculosis of lung, confirmed by sputum microscopy with or without culture', 'Certain infectious and parasitic diseases', 'Penyakit Infeksi & Parasit', 'Tuberkulosis (TB) Paru'],
      [15, 'B20', 'Human immunodeficiency virus [HIV] disease', 'Certain infectious and parasitic diseases', 'Penyakit Infeksi & Parasit', 'Penyakit HIV (Human Immunodeficiency Virus)'],
      [16, 'B54', 'Unspecified malaria', 'Certain infectious and parasitic diseases', 'Penyakit Infeksi & Parasit', 'Malaria'],
      [17, 'A90', 'Dengue fever [classical dengue]', 'Certain infectious and parasitic diseases', 'Penyakit Infeksi & Parasit', 'Demam Berdarah Dengue (DBD / Dengue Fever)'],
      [18, 'A01.0', 'Typhoid fever', 'Certain infectious and parasitic diseases', 'Penyakit Infeksi & Parasit', 'Demam Tifoid (Typhoid Fever)'],
      [19, 'R50.9', 'Fever, unspecified', 'Symptoms, signs and abnormal clinical findings', 'Gejala & Tanda Klinis', 'Febris / Demam Tidak Spesifik'],
      [20, 'R51', 'Headache', 'Symptoms, signs and abnormal clinical findings', 'Gejala & Tanda Klinis', 'Cephalgia / Sakit Kepala'],
      [21, 'M79.1', 'Myalgia', 'Diseases of the musculoskeletal system', 'Penyakit Muskuloskeletal', 'Myalgia / Nyeri Otot']
    ];
    shICD.getRange(2, 1, sampleICD.length, 6).setValues(sampleICD);
  }

  // 7. Sheet DXKeperawatan (Kolom A = No, Kolom B = Kategori_DXKeperawatan, Kolom C = DX_Keperawatan)
  const shDX = ensureSheetWithHeaders(ss, SHEET_NAMES.REF_DXKEP, ['No', 'Kategori_DXKeperawatan', 'DX_Keperawatan'], '#d63384');
  if (shDX.getLastRow() === 1) {
    const sampleDX = [
      [1, 'Fisiologis - Respirasi', 'Bersihan Jalan Napas Tidak Efektif (D.0001)'],
      [2, 'Fisiologis - Respirasi', 'Pola Napas Tidak Efektif (D.0005)'],
      [3, 'Fisiologis - Sirkulasi', 'Perfusi Perifer Tidak Efektif (D.0009)'],
      [4, 'Fisiologis - Nutrisi & Cairan', 'Defisit Nutrisi (D.0019)'],
      [5, 'Fisiologis - Nutrisi & Cairan', 'Ketidakstabilan Kadar Glukosa Darah (D.0027)'],
      [6, 'Fisiologis - Nutrisi & Cairan', 'Hipovolemia / Kekurangan Volume Cairan (D.0023)'],
      [7, 'Psikologis - Nyeri & Kenyamanan', 'Nyeri Akut (D.0077)'],
      [8, 'Psikologis - Nyeri & Kenyamanan', 'Nausea / Mual (D.0076)'],
      [9, 'Psikologis - Integritas Ego', 'Ansietas / Kecemasan (D.0080)'],
      [10, 'Perilaku - Penyuluhan & Pembelajaran', 'Defisit Pengetahuan tentang Kesehatan/Vaksinasi (D.0111)'],
      [11, 'Perilaku - Penyuluhan & Pembelajaran', 'Kesiapan Peningkatan Manajemen Kesehatan (D.0112)'],
      [12, 'Lingkungan - Keamanan & Proteksi', 'Hipertermia (D.0130)'],
      [13, 'Lingkungan - Keamanan & Proteksi', 'Risiko Infeksi (D.0142)']
    ];
    shDX.getRange(2, 1, sampleDX.length, 3).setValues(sampleDX);
  }

  // 8. Sheet Users (6 Kolom)
  const shUsers = ensureSheetWithHeaders(ss, SHEET_NAMES.USERS, ['Username', 'Password_Hash', 'Nama_Lengkap', 'Role', 'Wilker', 'Status_Aktif'], '#212529');
  if (shUsers.getLastRow() === 1) {
    const defaultUsers = [
      ['masteradmin', hashPassword('Master@BKK2025!'), 'Administrator Utama BKK Mataram', 'Master Admin', 'Induk', 'Aktif'],
      ['admininduk', hashPassword('Induk@2025'), 'Koordinator Pelayanan Induk', 'Admin Induk', 'Induk', 'Aktif'],
      ['petugas_lembar', hashPassword('Lembar@2025'), 'Petugas Medis Wilker Lembar', 'Member', 'Lembar', 'Aktif']
    ];
    shUsers.getRange(2, 1, defaultUsers.length, 6).setValues(defaultUsers);
  }

  // 9. Sheet Log_Arsip
  ensureSheetWithHeaders(ss, SHEET_NAMES.LOG_ARSIP, [
    'ID_Log', 'Timestamp', 'Tahun_Arsip', 'Mode_Arsip', 'Dieksekusi_Oleh',
    'Jumlah_Kunjungan', 'Jumlah_Lab', 'Jumlah_Obat', 'File_ID_Arsip', 'URL_File_Arsip'
  ], '#6c757d');

  // Clear Reference Cache
  const cache = CacheService.getScriptCache();
  cache.remove('SIKUTE_REF_DATA');

  return {
    success: true,
    message: 'Database relasional 9 Sheet berhasil diinisialisasi beserta referensi awal dan akun Master Admin.'
  };
}

function ensureSheetWithHeaders(ss, sheetName, headers, headerBgColor) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sh.getRange(1, 1, 1, headers.length);
    headerRange.setBackground(headerBgColor || '#0d6efd')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    const headerRange = sh.getRange(1, 1, 1, headers.length);
    headerRange.setBackground(headerBgColor || '#0d6efd')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  return sh;
}

/**
 * ============================================================================
 * 2. KEAMANAN ZERO-TRUST, HASHING & SESSION TOKEN (CacheService)
 * ============================================================================
 */
function hashPassword(rawPassword) {
  const secretSalt = 'SIKUTE_BKK_MATARAM_SALT_2025';
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(rawPassword) + secretSalt,
    Utilities.Charset.UTF_8
  );
  return digest.map(function(byte) {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function authenticateUser(username, password) {
  if (!username || !password) {
    throw new Error('Username dan Password wajib diisi.');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let shUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!shUsers) {
    setupDatabase();
    shUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  }

  const cleanUser = String(username).trim().toLowerCase();
  const inputHash = hashPassword(String(password));
  const data = shUsers.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowUser = String(data[i][0] || '').trim().toLowerCase();
    const rowPassHash = String(data[i][1] || '').trim();
    const rowNama = String(data[i][2] || '').trim();
    const rowRole = String(data[i][3] || '').trim();
    const rowWilker = String(data[i][4] || '').trim();
    const rowStatus = String(data[i][5] || 'Aktif').trim();

    if (rowUser === cleanUser) {
      if (rowStatus.toLowerCase() !== 'aktif') {
        throw new Error('Akun Anda sedang dinonaktifkan. Hubungi Master Admin.');
      }
      // Support hashed or legacy plain match during migration
      if (rowPassHash === inputHash || rowPassHash === String(password)) {
        const sessionToken = Utilities.getUuid() + '-' + Utilities.getUuid();
        const sessionPayload = {
          username: String(data[i][0]).trim(),
          nama: rowNama,
          role: rowRole,
          wilker: rowWilker,
          loginTime: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
        };
        const rawSession = JSON.stringify(sessionPayload);
        const cache = CacheService.getScriptCache();
        cache.put('SESSION_' + sessionToken, rawSession, SESSION_TTL_SECONDS);
        try {
          PropertiesService.getScriptProperties().setProperty('SESSION_' + sessionToken, rawSession);
        } catch (e) {}

        return {
          success: true,
          sessionToken: sessionToken,
          user: sessionPayload
        };
      } else {
        throw new Error('Username atau Password tidak valid.');
      }
    }
  }
  throw new Error('Username atau Password tidak valid.');
}

function validateSession(sessionToken, allowedRoles) {
  if (!sessionToken || typeof sessionToken !== 'string') {
    throw new Error('Akses Ditolak: Sesi tidak ditemukan. Silakan login kembali.');
  }
  const cache = CacheService.getScriptCache();
  let raw = cache.get('SESSION_' + sessionToken);
  if (!raw) {
    try {
      raw = PropertiesService.getScriptProperties().getProperty('SESSION_' + sessionToken);
    } catch (e) {}
  }
  if (!raw) {
    throw new Error('Akses Ditolak: Sesi Anda telah berakhir (Expired). Silakan login kembali.');
  }
  const session = JSON.parse(raw);
  if (allowedRoles && Array.isArray(allowedRoles) && allowedRoles.length > 0) {
    if (allowedRoles.indexOf(session.role) === -1) {
      throw new Error('Akses Ditolak (403): Role "' + session.role + '" tidak memiliki izin untuk operasi ini.');
    }
  }
  // Perpanjang masa aktif sesi saat aktif bertransaksi
  cache.put('SESSION_' + sessionToken, raw, SESSION_TTL_SECONDS);
  return session;
}

function logoutUser(sessionToken) {
  if (sessionToken) {
    CacheService.getScriptCache().remove('SESSION_' + sessionToken);
    try {
      PropertiesService.getScriptProperties().deleteProperty('SESSION_' + sessionToken);
    } catch (e) {}
  }
  return { success: true };
}

function verifyActiveSession(sessionToken) {
  const session = validateSession(sessionToken);
  return { success: true, user: session };
}

/**
 * ============================================================================
 * 3. UTILITAS KONVERSI TANGGAL & FORMATTER
 * ============================================================================
 */
function formatDateYMD(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(val).trim();
  if (!str) return '';
  // Format YYYY-MM-DD atau YYYY/MM/DD
  const ymdMatch = str.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (ymdMatch) {
    return ymdMatch[1] + '-' + ('0' + ymdMatch[2]).slice(-2) + '-' + ('0' + ymdMatch[3]).slice(-2);
  }
  // Format DD/MM/YYYY atau DD-MM-YYYY (Format Indonesia)
  const dmyMatch = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (dmyMatch) {
    return dmyMatch[3] + '-' + ('0' + dmyMatch[2]).slice(-2) + '-' + ('0' + dmyMatch[1]).slice(-2);
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return str;
}

function formatDateTimeYMDHMS(val) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(val);
}

/**
 * Mengonversi nilai sel Google Sheets (termasuk sel angka yang tidak sengaja terformat sebagai Date oleh Sheets,
 * misalnya "9.6" yang diubah Sheets menjadi 6 September) menjadi string/angka aman agar google.script.run
 * TIDAK PERNAH mengembalikan null akibat objek Date mentah.
 */
function safeScalar(val, defaultVal) {
  const fallback = (defaultVal !== undefined) ? defaultVal : '';
  if (val === null || val === undefined || val === '') return fallback;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return fallback;
    // Jika sel bukan kolom tanggal tetapi terformat Date oleh Google Sheets (misal 9.6 -> 06 Sep),
    // kembalikan format ringkas tanggal/angka agar tidak merusak serialisasi JSON google.script.run
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof val === 'number') {
    if (isNaN(val) || !isFinite(val)) return fallback;
    return val;
  }
  if (typeof val === 'boolean') return val ? 'Ya' : 'Tidak';
  return String(val).trim();
}

function safeNumericText(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    // Jika pengguna mengetik "6.9" atau "9.6" pada kolom Asam Urat / Hb / IMT lalu Google Sheets mengubahnya menjadi Date:
    const d = val.getDate();
    const m = val.getMonth() + 1;
    return d + '.' + m;
  }
  if (typeof val === 'number') {
    if (isNaN(val) || !isFinite(val)) return '';
    return String(val);
  }
  const s = String(val).trim();
  // Jika sudah terlanjur tersimpan sebagai string "Sun Sep 06 2026 00:00:00 GMT+0800..."
  if (/GMT[+-]\d{4}/.test(s)) {
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      return parsed.getDate() + '.' + (parsed.getMonth() + 1);
    }
  }
  return s;
}

/**
 * ============================================================================
 * 4. CACHING REFERENSI (`Obat.ref`, `DataKotakab`, `DXKeperawatan`) & ICD-10
 * ============================================================================
 */
function getInitialReferences(sessionToken) {
  validateSession(sessionToken);
  const cache = CacheService.getScriptCache();
  const cached = cache.get('SIKUTE_REF_DATA_V3');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // Lanjut baca dari sheet jika korup
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obat.ref (Kolom A = no, Kolom B = Nama Obat, Kolom C = Sediaan)
  const obatList = [];
  const shObat = ss.getSheetByName(SHEET_NAMES.REF_OBAT);
  if (shObat && shObat.getLastRow() > 1) {
    const vals = shObat.getRange(2, 1, shObat.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < vals.length; i++) {
      const nama = String(vals[i][1] || '').trim();
      const sediaan = String(vals[i][2] || '').trim();
      if (nama) {
        obatList.push({ no: vals[i][0], nama: nama, sediaan: sediaan });
      }
    }
  }

  // 2. DataKotakab (Kolom B = Provinsi, Kolom C = Kabupaten/Kota)
  const kotaMap = {};
  const shKota = ss.getSheetByName(SHEET_NAMES.REF_KOTAKAB);
  if (shKota && shKota.getLastRow() > 1) {
    const vals = shKota.getRange(2, 1, shKota.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < vals.length; i++) {
      const prov = String(vals[i][1] || '').trim();
      const kab = String(vals[i][2] || '').trim();
      if (prov && kab) {
        if (!kotaMap[prov]) kotaMap[prov] = [];
        if (kotaMap[prov].indexOf(kab) === -1) {
          kotaMap[prov].push(kab);
        }
      }
    }
  }

  // 3. DXKeperawatan (Kolom B = Kategori_DXKeperawatan, Kolom C = DX_Keperawatan)
  const dxKepList = [];
  const shDX = ss.getSheetByName(SHEET_NAMES.REF_DXKEP);
  if (shDX && shDX.getLastRow() > 1) {
    const vals = shDX.getRange(2, 1, shDX.getLastRow() - 1, 3).getValues();
    for (let i = 0; i < vals.length; i++) {
      const kat = String(vals[i][1] || '').trim();
      const dx = String(vals[i][2] || '').trim();
      if (kat && dx) {
        dxKepList.push({ kategori: kat, diagnosa: dx });
      }
    }
  }

  // 4. Diagnosa Kerja dari Sheet ICD-10 Kolom E dan Kolom F (Tersinkronisasi Kolom E <-> Kolom F, tidak terikat dengan Diagnosa ICD-10 Kolom B-D)
  const diagnosaKerjaList = [];
  let icdColEHeader = 'Kategori / Kelompok Diagnosa Kerja';
  let icdColFHeader = 'Diagnosa Kerja';
  const shICD = ss.getSheetByName(SHEET_NAMES.REF_ICD10);
  if (shICD && shICD.getLastRow() > 1 && shICD.getLastColumn() >= 5) {
    const maxCol = Math.min(6, shICD.getLastColumn());
    const hdrs = shICD.getRange(1, 5, 1, maxCol - 4).getValues()[0];
    if (hdrs[0]) icdColEHeader = String(hdrs[0]).trim();
    if (hdrs.length > 1 && hdrs[1]) icdColFHeader = String(hdrs[1]).trim();

    const valsDK = shICD.getRange(2, 5, shICD.getLastRow() - 1, maxCol - 4).getValues(); // Hanya Kolom E & F
    const seenDK = {};
    let emptyStreak = 0;
    for (let i = 0; i < valsDK.length; i++) {
      const colE = String(valsDK[i][0] || '').trim();
      const colF = (valsDK[i].length > 1) ? String(valsDK[i][1] || '').trim() : '';
      if (!colE && !colF) {
        emptyStreak++;
        if (emptyStreak > 500 && diagnosaKerjaList.length > 0) break;
        continue;
      }
      emptyStreak = 0;
      const key = colE + '||' + colF;
      if (!seenDK[key]) {
        seenDK[key] = true;
        diagnosaKerjaList.push({
          colE: colE,
          colF: colF || colE
        });
        if (diagnosaKerjaList.length >= 2000) break;
      }
    }
  }

  // 5. Cache Ringkas ICD-10 (300 baris awal) untuk Pencarian Saat Mode Offline Lapangan
  const icd10OfflineList = [];
  if (shICD && shICD.getLastRow() > 1) {
    const maxIcdRows = Math.min(300, shICD.getLastRow() - 1);
    const valsIcdShort = shICD.getRange(2, 2, maxIcdRows, 3).getValues();
    for (let i = 0; i < valsIcdShort.length; i++) {
      const c = String(valsIcdShort[i][0] || '').trim();
      const d = String(valsIcdShort[i][1] || '').trim();
      const g = String(valsIcdShort[i][2] || '').trim();
      if (c || d) {
        icd10OfflineList.push({ code: c, diagnosa: d, group: g });
      }
    }
  }

  // 6. Cache Pasien Terakhir (150 Pasien Unik Terbaru) untuk Pencarian RM / NIK Saat Offline
  const recentPatientsCache = [];
  const shKunjCache = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  if (shKunjCache && shKunjCache.getLastRow() > 1) {
    const totalKRows = shKunjCache.getLastRow() - 1;
    const readCount = Math.min(250, totalKRows);
    const startR = Math.max(2, shKunjCache.getLastRow() - readCount + 1);
    const kVals = shKunjCache.getRange(startR, 1, readCount, 26).getValues();
    const seenRM = {};
    for (let i = kVals.length - 1; i >= 0; i--) {
      const r = kVals[i];
      const rm = String(r[6] || '').trim();
      const nik = String(r[7] || '').trim();
      const nama = String(r[18] || '').trim();
      if (!rm && !nik) continue;
      const keyP = rm + '|' + nik;
      if (!seenRM[keyP]) {
        seenRM[keyP] = true;
        recentPatientsCache.push({
          noRM: rm,
          noIdentitas: nik,
          namaPasien: nama,
          jk: String(r[19] || ''),
          umur: r[20] !== '' ? Number(r[20]) : '',
          kategoriUmur: String(r[21] || ''),
          alamat: String(r[22] || ''),
          provinsi: String(r[23] || ''),
          kabKota: String(r[24] || ''),
          pekerjaan: String(r[25] || 'Umum'),
          lastVisitDate: formatDateYMD(r[2]),
          lastWilker: String(r[3] || '')
        });
        if (recentPatientsCache.length >= 150) break;
      }
    }
  }

  const payload = {
    wilkerList: WILKER_LIST,
    obatRef: obatList,
    kotaKabMap: kotaMap,
    dxKepList: dxKepList,
    icdColEHeader: icdColEHeader,
    icdColFHeader: icdColFHeader,
    diagnosaKerjaList: diagnosaKerjaList,
    icd10OfflineList: icd10OfflineList,
    recentPatientsCache: recentPatientsCache,
    cachedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
  };

  try {
    const strPayload = JSON.stringify(payload);
    if (strPayload.length < 95000) {
      cache.put('SIKUTE_REF_DATA_V3', strPayload, CACHE_REF_TTL_SECONDS);
    }
  } catch (e) {}

  return payload;
}

/**
 * Autocomplete ICD-10 (>41.000 baris — Kolom B, C, D) dengan batas hasil maksimal 30 baris
 */
function searchICD10(sessionToken, query) {
  validateSession(sessionToken);
  if (!query || String(query).trim().length < 2) {
    return [];
  }
  const q = String(query).trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.REF_ICD10);
  if (!sh || sh.getLastRow() <= 1) return [];

  const lastRow = sh.getLastRow();
  const vals = sh.getRange(2, 2, lastRow - 1, 3).getValues(); // Hanya Kolom B, C, D
  const results = [];

  for (let i = 0; i < vals.length; i++) {
    const code = String(vals[i][0] || '').trim();
    const diag = String(vals[i][1] || '').trim();
    const grp = String(vals[i][2] || '').trim();
    if (!code && !diag) continue;
    if (
      code.toLowerCase().indexOf(q) !== -1 ||
      diag.toLowerCase().indexOf(q) !== -1
    ) {
      results.push({
        code: code,
        diagnosa: diag,
        group: grp
      });
      if (results.length >= 30) break;
    }
  }
  return results;
}

/**
 * ============================================================================
 * 5. PENCARIAN PASIEN & DETAIL KUNJUNGAN VIA `TextFinder` (O(1) Row Read)
 * ============================================================================
 */
function getPatientByRMOrIdentity(sessionToken, keyword) {
  validateSession(sessionToken);
  if (!keyword || !String(keyword).trim()) {
    return { found: false, message: 'Masukkan Nomor RM atau Nomor KTP/Paspor terlebih dahulu.' };
  }
  const cleanKey = String(keyword).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  if (!sh || sh.getLastRow() <= 1) {
    return { found: false, message: 'Belum ada data kunjungan di database.' };
  }

  const lastRow = sh.getLastRow();
  // Cari di Kolom G (No_RM - indeks kolom 7) dan Kolom H (No_Identitas - indeks kolom 8) dari bawah ke atas
  const searchRange = sh.getRange(2, 7, lastRow - 1, 2);
  const finder = searchRange.createTextFinder(cleanKey).matchEntireCell(true).matchCase(false);
  finder.findNext(); // Posisi awal sebelum mundur
  const matchedCell = finder.findPrevious();

  if (!matchedCell) {
    // Fallback pencarian parsial jika tidak persis
    const finderPartial = searchRange.createTextFinder(cleanKey).matchEntireCell(false).matchCase(false);
    const partialCell = finderPartial.findPrevious();
    if (!partialCell) {
      return { found: false, message: 'Data pasien dengan No. RM / KTP / Paspor "' + cleanKey + '" tidak ditemukan.' };
    }
    return extractPatientFromRow(sh, partialCell.getRow());
  }

  return extractPatientFromRow(sh, matchedCell.getRow());
}

function extractPatientFromRow(sh, rowIndex) {
  const rowVals = sh.getRange(rowIndex, 1, 1, 51).getValues()[0];
  return {
    found: true,
    patient: {
      noRM: String(rowVals[6] || ''),
      noIdentitas: String(rowVals[7] || ''),
      namaPasien: String(rowVals[18] || ''),
      jk: String(rowVals[19] || ''),
      umur: rowVals[20] !== '' ? Number(rowVals[20]) : '',
      kategoriUmur: String(rowVals[21] || ''),
      alamat: String(rowVals[22] || ''),
      provinsi: String(rowVals[23] || ''),
      kabKota: String(rowVals[24] || ''),
      pekerjaan: String(rowVals[25] || 'Umum'),
      lastVisitDate: formatDateYMD(rowVals[2]),
      lastWilker: String(rowVals[3] || '')
    }
  };
}

/**
 * Ambil detail lengkap kunjungan dari 3 Sheet Relasional via `TextFinder`
 */
function getVisitDetailById(sessionToken, idKunjungan) {
  const session = validateSession(sessionToken);
  if (!idKunjungan) throw new Error('ID Kunjungan wajib disertakan.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  if (!shKunj || shKunj.getLastRow() <= 1) {
    throw new Error('Data kunjungan tidak ditemukan.');
  }

  const cleanKey = String(idKunjungan).trim();
  let targetRow = -1;

  // Support fallback row reference "ROW-X" jika baris di-paste manual tanpa ID_Kunjungan
  if (/^ROW-\d+$/.test(cleanKey)) {
    targetRow = parseInt(cleanKey.replace('ROW-', ''), 10);
  } else {
    // 1. Cari di Data_Kunjungan (Kolom A = ID_Kunjungan, atau fallback Kolom G = No_RM)
    const finderKunj = shKunj.getRange(2, 1, shKunj.getLastRow() - 1, 1)
      .createTextFinder(cleanKey)
      .matchEntireCell(true);
    const cellKunj = finderKunj.findNext();
    if (cellKunj) {
      targetRow = cellKunj.getRow();
    } else {
      const finderRM = shKunj.getRange(2, 7, shKunj.getLastRow() - 1, 1)
        .createTextFinder(cleanKey)
        .matchEntireCell(true);
      const cellRM = finderRM.findPrevious();
      if (cellRM) targetRow = cellRM.getRow();
    }
  }

  if (targetRow < 2 || targetRow > shKunj.getLastRow()) {
    throw new Error('ID Kunjungan / No. RM "' + idKunjungan + '" tidak ditemukan.');
  }

  const rKunj = shKunj.getRange(targetRow, 1, 1, Math.max(54, shKunj.getLastColumn())).getValues()[0];
  if (!rKunj[0]) {
    rKunj[0] = 'KNJ-ROW-' + targetRow;
    shKunj.getRange(targetRow, 1).setValue(rKunj[0]);
  }
  const kunjunganObj = mapRowToKunjunganObject(rKunj);

  // Proteksi RBAC Member: Hanya boleh akses wilkernya sendiri
  if (session.role === 'Member' && kunjunganObj.Wilker !== session.wilker) {
    throw new Error('Akses Ditolak: Anda hanya diizinkan mengakses data Wilayah Kerja ' + session.wilker + '.');
  }

  // 2. Cari di Data_Laboratorium (Kolom B = ID_Kunjungan, atau fallback Kolom E = No_RM)
  let labObj = null;
  const shLab = ss.getSheetByName(SHEET_NAMES.LAB);
  if (shLab && shLab.getLastRow() > 1) {
    let cellLab = shLab.getRange(2, 2, shLab.getLastRow() - 1, 1)
      .createTextFinder(String(kunjunganObj.ID_Kunjungan || idKunjungan).trim())
      .matchEntireCell(true)
      .findNext();
    if (!cellLab && kunjunganObj.No_RM) {
      cellLab = shLab.getRange(2, 5, shLab.getLastRow() - 1, 1)
        .createTextFinder(String(kunjunganObj.No_RM).trim())
        .matchEntireCell(true)
        .findPrevious();
    }
    if (cellLab) {
      const rLab = shLab.getRange(cellLab.getRow(), 1, 1, Math.max(47, shLab.getLastColumn())).getValues()[0];
      labObj = mapRowToLabObject(rLab);
    }
  }

  // 3. Cari di Data_Terapi_Obat (Kolom B = ID_Kunjungan, Relasi 1-to-Many)
  const obatList = [];
  const shObat = ss.getSheetByName(SHEET_NAMES.OBAT);
  if (shObat && shObat.getLastRow() > 1) {
    const finderObat = shObat.getRange(2, 2, shObat.getLastRow() - 1, 1)
      .createTextFinder(String(idKunjungan).trim())
      .matchEntireCell(true);
    const matchedCells = finderObat.findAll();
    for (let i = 0; i < matchedCells.length; i++) {
      const rObat = shObat.getRange(matchedCells[i].getRow(), 1, 1, 10).getValues()[0];
      obatList.push(mapRowToObatObject(rObat));
    }
  }

  return {
    success: true,
    kunjungan: kunjunganObj,
    laboratorium: labObj,
    terapiObat: obatList
  };
}

function mapRowToKunjunganObject(r) {
  const tgl = formatDateYMD(r[2]) || formatDateYMD(r[1]);
  return {
    ID_Kunjungan: safeScalar(r[0], ''),
    Timestamp: formatDateTimeYMDHMS(r[1]),
    Tanggal_Kunjungan: tgl,
    Wilker: safeScalar(r[3], ''),
    Petugas_Input: safeScalar(r[4], ''),
    Mode_RM: safeScalar(r[5], 'Otomatis'),
    No_RM: safeScalar(r[6], ''),
    No_Identitas: safeScalar(r[7], ''),
    Jenis_Kunjungan: safeScalar(r[8], ''),
    Sub_Kunjungan: safeScalar(r[9], ''),
    Fase_Haji: safeScalar(r[10], ''),
    Kloter_Haji: safeScalar(r[11], ''),
    Vaksin_Dipilih: safeScalar(r[12], ''),
    Tarif_Pemeriksaan: Number(r[13] || 0) || 0,
    Tarif_ICV: Number(r[14] || 0) || 0,
    Tarif_Kontraindikasi: Number(r[15] || 0) || 0,
    Tarif_Vaksin: Number(r[16] || 0) || 0,
    Total_PNBP: Number(r[17] || 0) || 0,
    Nama_Pasien: safeScalar(r[18], ''),
    JK: safeScalar(r[19], ''),
    Umur: safeNumericText(r[20]),
    Kategori_Umur: safeScalar(r[21], ''),
    Alamat: safeScalar(r[22], ''),
    Provinsi: safeScalar(r[23], ''),
    KabKota: safeScalar(r[24], ''),
    Pekerjaan: safeScalar(r[25], 'Umum'),
    Keluhan: safeScalar(r[26], ''),
    Sistolik: safeNumericText(r[27]),
    Diastolik: safeNumericText(r[28]),
    Kategori_TD: safeScalar(r[29], ''),
    Respirasi: safeNumericText(r[30]),
    Nadi: safeNumericText(r[31]),
    SpO2: safeNumericText(r[32]),
    Suhu: safeNumericText(r[33]),
    TB: safeNumericText(r[34]),
    BB: safeNumericText(r[35]),
    IMT: safeNumericText(r[36]),
    Kategori_IMT: safeScalar(r[37], ''),
    Lingkar_Perut: safeNumericText(r[38]),
    Buta_Warna: safeScalar(r[39], ''),
    ICD10_Code: safeScalar(r[40], ''),
    Diagnosa_ICD10: safeScalar(r[41], ''),
    Group_Diag: safeScalar(r[42], ''),
    Kategori_DXKep: safeScalar(r[43], ''),
    DX_Keperawatan: safeScalar(r[44], ''),
    FR_Ditemukan: safeScalar(r[45], ''),
    FR_Dikendalikan: safeScalar(r[46], ''),
    Kategori_Penumpang: safeScalar(r[47], ''),
    Dokter_Pemeriksa: safeScalar(r[48], ''),
    Ada_Lab: safeScalar(r[49], 'Tidak'),
    Ada_Obat: safeScalar(r[50], 'Tidak'),
    Lokasi_Rujukan: safeScalar(r[51], ''),
    Diagnosa_Kerja_E: safeScalar(r[52], ''),
    Diagnosa_Kerja: safeScalar(r[53], '')
  };
}

function mapRowToLabObject(r) {
  return {
    ID_Lab: safeScalar(r[0], ''),
    ID_Kunjungan: safeScalar(r[1], ''),
    Tanggal_Kunjungan: formatDateYMD(r[2]),
    Wilker: safeScalar(r[3], ''),
    No_RM: safeScalar(r[4], ''),
    Nama_Pasien: safeScalar(r[5], ''),
    Hemoglobin: safeNumericText(r[6]),
    Hematokrit: safeNumericText(r[7]),
    Leukosit: safeNumericText(r[8]),
    Platelet: safeNumericText(r[9]),
    Eritrosit: safeNumericText(r[10]),
    GDP: safeNumericText(r[11]),
    Interp_GDP: safeScalar(r[12], ''),
    GDS: safeNumericText(r[13]),
    Interp_GDS: safeScalar(r[14], ''),
    GD_2JPP: safeNumericText(r[15]),
    Interp_GD_2JPP: safeScalar(r[16], ''),
    Kolesterol_Total: safeNumericText(r[17]),
    Interp_Kol_Total: safeScalar(r[18], ''),
    LDL: safeNumericText(r[19]),
    Interp_LDL: safeScalar(r[20], ''),
    HDL: safeNumericText(r[21]),
    Interp_HDL: safeScalar(r[22], ''),
    Trigliserida: safeNumericText(r[23]),
    Interp_Trigliserida: safeScalar(r[24], ''),
    Asam_Urat: safeNumericText(r[25]),
    Warna_Urine: safeScalar(r[26], ''),
    Bau_Urine: safeScalar(r[27], ''),
    Kejernihan_Urine: safeScalar(r[28], ''),
    pH_Urine: safeNumericText(r[29]),
    Protein_Urine: safeScalar(r[30], ''),
    Glukosa_Urine: safeScalar(r[31], ''),
    Keton_Urine: safeScalar(r[32], ''),
    Bilirubin_Urine: safeScalar(r[33], ''),
    Urobilinogen_Urine: safeScalar(r[34], ''),
    Darah_Samar_Urine: safeScalar(r[35], ''),
    PPTest_Kehamilan: safeScalar(r[36], '-'),
    Malaria_Test: safeScalar(r[37], '-'),
    Dengue_Test: safeScalar(r[38], '-'),
    HIV_Test: safeScalar(r[39], '-'),
    Widal_Test: safeScalar(r[40], '-'),
    RDT_Covid19: safeScalar(r[41], '-'),
    TB_Test: safeScalar(r[42], '-'),
    EKG: safeScalar(r[43], ''),
    Status_Abnormal_Positif: safeScalar(r[44], 'Normal'),
    Interp_Asam_Urat: safeScalar(r[45], ''),
    Jenis_Dengue: safeScalar(r[46], '-')
  };
}

function mapRowToObatObject(r) {
  return {
    ID_Terapi: safeScalar(r[0], ''),
    ID_Kunjungan: safeScalar(r[1], ''),
    Tanggal_Kunjungan: formatDateYMD(r[2]),
    Wilker: safeScalar(r[3], ''),
    No_RM: safeScalar(r[4], ''),
    Nama_Pasien: safeScalar(r[5], ''),
    Nama_Obat: safeScalar(r[6], ''),
    Sediaan: safeScalar(r[7], ''),
    Jumlah: Number(r[8] || 0) || 0,
    Dokter_Pemeriksa: safeScalar(r[9], '')
  };
}

/**
 * ============================================================================
 * 6. TRANSAKSI SIMPAN & EDIT RELASIONAL 3 SHEET DENGAN `LockService`
 * ============================================================================
 */
function generateNextNoRM(shKunjungan, wilker, tanggalStr) {
  const code = WILKER_CODES[wilker] || 'BKK';
  const ymd = String(tanggalStr || '').replace(/-/g, '') || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const prefix = 'RM-' + code + '-' + ymd + '-';

  let maxSeq = 0;
  const lastRow = shKunjungan.getLastRow();
  if (lastRow > 1) {
    const finder = shKunjungan.getRange(2, 7, lastRow - 1, 1).createTextFinder(prefix);
    const cells = finder.findAll();
    for (let i = 0; i < cells.length; i++) {
      const val = String(cells[i].getValue() || '');
      const parts = val.split('-');
      const seqNum = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(seqNum) && seqNum > maxSeq) {
        maxSeq = seqNum;
      }
    }
  }
  const nextSeq = ('0000' + (maxSeq + 1)).slice(-4);
  return prefix + nextSeq;
}

function evaluateLabAbnormalStatus(lab) {
  if (!lab) return 'Normal';
  const flags = [];
  if (lab.Interp_GDP && lab.Interp_GDP !== 'Normal') flags.push('GDP: ' + lab.Interp_GDP);
  if (lab.Interp_GDS && lab.Interp_GDS !== 'Normal') flags.push('GDS: ' + lab.Interp_GDS);
  if (lab.Interp_GD_2JPP && lab.Interp_GD_2JPP !== 'Normal') flags.push('GD2JPP: ' + lab.Interp_GD_2JPP);
  if (lab.Interp_Kol_Total && lab.Interp_Kol_Total !== 'Normal') flags.push('Kolesterol: ' + lab.Interp_Kol_Total);
  if (lab.Interp_LDL && lab.Interp_LDL !== 'Optimal (Normal)' && lab.Interp_LDL !== 'Mendekati Optimal') flags.push('LDL: ' + lab.Interp_LDL);
  if (lab.Interp_HDL && lab.Interp_HDL === 'Rendah (Berisiko)') flags.push('HDL Rendah');
  if (lab.Interp_Trigliserida && lab.Interp_Trigliserida !== 'Normal') flags.push('Trigliserida: ' + lab.Interp_Trigliserida);
  if (lab.Interp_Asam_Urat && lab.Interp_Asam_Urat !== 'Normal') flags.push('Asam Urat: ' + lab.Interp_Asam_Urat);

  if (lab.PPTest_Kehamilan === 'Positif') flags.push('PPTest Positif');
  if (lab.Malaria_Test === 'Positif') flags.push('Malaria Positif');
  if (lab.Dengue_Test && String(lab.Dengue_Test).indexOf('Positif') !== -1) {
    const jnsD = (lab.Jenis_Dengue && lab.Jenis_Dengue !== '-') ? (' (' + lab.Jenis_Dengue + ')') : '';
    flags.push('Dengue' + jnsD + ': ' + lab.Dengue_Test);
  }
  if (lab.HIV_Test === 'Positif') flags.push('HIV Positif');
  if (lab.Widal_Test === 'Sedang (1:160)' || lab.Widal_Test === 'Tinggi (> 1:320)') flags.push('Widal ' + lab.Widal_Test);
  if (lab.RDT_Covid19 === 'Positif') flags.push('Covid-19 Positif');
  if (lab.TB_Test === 'Positif') flags.push('TB Positif');

  return flags.length > 0 ? flags.join('; ') : 'Normal';
}

function hasAnyLabInput(lab) {
  if (!lab) return false;
  const checkKeys = [
    'Hemoglobin', 'Hematokrit', 'Leukosit', 'Platelet', 'Eritrosit',
    'GDP', 'GDS', 'GD_2JPP', 'Kolesterol_Total', 'LDL', 'HDL', 'Trigliserida', 'Asam_Urat',
    'Warna_Urine', 'Bau_Urine', 'Kejernihan_Urine', 'pH_Urine', 'Protein_Urine',
    'Glukosa_Urine', 'Keton_Urine', 'Bilirubin_Urine', 'Urobilinogen_Urine', 'Darah_Samar_Urine', 'EKG'
  ];
  for (let i = 0; i < checkKeys.length; i++) {
    if (lab[checkKeys[i]] !== undefined && String(lab[checkKeys[i]]).trim() !== '') return true;
  }
  const dropdownKeys = ['PPTest_Kehamilan', 'Malaria_Test', 'Jenis_Dengue', 'Dengue_Test', 'HIV_Test', 'Widal_Test', 'RDT_Covid19', 'TB_Test'];
  for (let j = 0; j < dropdownKeys.length; j++) {
    const val = String(lab[dropdownKeys[j]] || '-').trim();
    if (val !== '' && val !== '-') return true;
  }
  return false;
}

function validateMandatoryVisitServer(k, lab, validObatItems) {
  const missing = [];
  const jenis = String(k.Jenis_Kunjungan || '').trim();
  const sub = String(k.Sub_Kunjungan || '').trim();
  const jk = String(k.JK || '').trim();
  const umurNum = (k.Umur !== '' && k.Umur !== undefined && k.Umur !== null) ? Number(k.Umur) : NaN;

  // 1. Provinsi & Kabupaten/Kota
  if (!String(k.Provinsi || '').trim()) missing.push('Provinsi');
  if (!String(k.KabKota || '').trim()) missing.push('Kabupaten/Kota');

  // 2. Keluhan Pasien
  if (!String(k.Keluhan || '').trim()) missing.push('Keluhan Pasien');

  // 3. Pemeriksaan Fisik (Tekanan Darah, Respirasi, SpO2, Suhu Tubuh) Wajib untuk Pengobatan, Skrining CKG, Skrining Penjamah Makanan, Kier Kesehatan
  const wajibFisik = (
    jenis === 'Pengobatan' ||
    jenis === 'Kier Kesehatan' ||
    (jenis === 'Skrining' && (sub === 'Cek Kesehatan Gratis (CKG)' || sub === 'Skrining Penjamah Makanan'))
  );
  if (wajibFisik) {
    const fisikMissing = [];
    if (String(k.Sistolik || '').trim() === '' || String(k.Diastolik || '').trim() === '') {
      fisikMissing.push('Tekanan Darah (Sistolik & Diastolik)');
    }
    if (String(k.Respirasi || '').trim() === '') fisikMissing.push('Respirasi');
    if (String(k.SpO2 || '').trim() === '') fisikMissing.push('SpO2');
    if (String(k.Suhu || '').trim() === '') fisikMissing.push('Suhu Tubuh');
    if (fisikMissing.length > 0) {
      missing.push('Pemeriksaan Fisik Wajib (' + fisikMissing.join(', ') + ')');
    }
  }

  // 4. Pengisian ICD-10
  if (!String(k.ICD10_Code || '').trim() || !String(k.Diagnosa_ICD10 || '').trim()) {
    missing.push('Pengisian ICD-10 (Kode & Diagnosa ICD-10)');
  }

  // 5. Pengisian Diagnosa Kerja
  if (!String(k.Diagnosa_Kerja_E || '').trim() || !String(k.Diagnosa_Kerja || '').trim()) {
    missing.push('Pengisian Diagnosa Kerja (Kolom E & Kolom F)');
  }

  // 6. Pemeriksaan PP Test untuk Wanita Usia Subur (Perempuan 15-49 Tahun) dengan Kunjungan Vaksinasi
  const isWUSVaksinasi = (jenis === 'Vaksinasi' && jk === 'Perempuan' && !isNaN(umurNum) && umurNum >= 15 && umurNum <= 49);
  if (isWUSVaksinasi) {
    const ppVal = String((lab && lab.PPTest_Kehamilan) || '-').trim();
    if (!ppVal || ppVal === '-') {
      missing.push('Pemeriksaan PP Test Kehamilan (Wajib untuk Wanita Usia Subur 15-49 Tahun pada Kunjungan Vaksinasi)');
    }
  }

  // 7. Faktor Risiko Ditemukan dan Dikendalikan
  const frD = String(k.FR_Ditemukan || '').trim();
  const frK = String(k.FR_Dikendalikan || '').trim();
  if (!frD) missing.push('Faktor Risiko (FR) Ditemukan');
  if (!frK) missing.push('Pengendalian Faktor Risiko (FR Dikendalikan)');
  if (frK.toLowerCase().indexOf('rujukan') !== -1 && !String(k.Lokasi_Rujukan || '').trim()) {
    missing.push('Lokasi Rujukan (Puskesmas/Rumah Sakit)');
  }

  // 8. Nama Obat bila Pengendalian Faktor Risiko Diobati dan Vaksinasi
  if (frK === 'Diobati' || frK === 'Vaksinasi' || frK === 'Pengobatan') {
    if (!Array.isArray(validObatItems) || validObatItems.length === 0) {
      missing.push('Nama Obat / Terapi Obat (Wajib diisi karena Pengendalian FR = ' + frK + ')');
    }
  }

  if (missing.length > 0) {
    throw new Error('Data wajib belum lengkap: ' + missing.join('; ') + '.');
  }
}

function saveVisitTransaction(sessionToken, payload) {
  const session = validateSession(sessionToken, ['Master Admin', 'Admin Induk', 'Member']);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
    const shLab = ss.getSheetByName(SHEET_NAMES.LAB);
    const shObat = ss.getSheetByName(SHEET_NAMES.OBAT);

    const k = payload.kunjungan || {};
    const lab = payload.laboratorium || {};
    const obatItems = Array.isArray(payload.terapiObat) ? payload.terapiObat : [];

    // Enforce Wilker untuk Member
    const effectiveWilker = (session.role === 'Member') ? session.wilker : (k.Wilker || session.wilker || 'Induk');
    const tanggalKunjungan = formatDateYMD(k.Tanggal_Kunjungan) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    const isEdit = Boolean(k.ID_Kunjungan && String(k.ID_Kunjungan).trim() !== '');
    const idKunjungan = isEdit ? String(k.ID_Kunjungan).trim() : ('KNJ-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(1000 + Math.random() * 9000));

    let noRM = String(k.No_RM || '').trim();
    if (!noRM || (k.Mode_RM === 'Otomatis' && !isEdit)) {
      noRM = generateNextNoRM(shKunj, effectiveWilker, tanggalKunjungan);
    }

    const adaLabBool = hasAnyLabInput(lab);
    const validObatItems = obatItems.filter(function(item) {
      return item && String(item.Nama_Obat || '').trim() !== '' && Number(item.Jumlah || 0) > 0;
    });
    const adaObatBool = validObatItems.length > 0;

    validateMandatoryVisitServer(k, lab, validObatItems);

    const rowKunjungan = [
      idKunjungan,
      nowStr,
      tanggalKunjungan,
      effectiveWilker,
      session.nama + ' (' + session.username + ')',
      k.Mode_RM || 'Otomatis',
      noRM,
      String(k.No_Identitas || '').trim(),
      String(k.Jenis_Kunjungan || ''),
      String(k.Sub_Kunjungan || ''),
      String(k.Fase_Haji || ''),
      String(k.Kloter_Haji || ''),
      String(k.Vaksin_Dipilih || ''),
      Number(k.Tarif_Pemeriksaan || 0),
      Number(k.Tarif_ICV || 0),
      Number(k.Tarif_Kontraindikasi || 0),
      Number(k.Tarif_Vaksin || 0),
      Number(k.Total_PNBP || 0),
      String(k.Nama_Pasien || '').trim(),
      String(k.JK || ''),
      k.Umur !== '' && k.Umur !== undefined ? Number(k.Umur) : '',
      String(k.Kategori_Umur || ''),
      String(k.Alamat || ''),
      String(k.Provinsi || ''),
      String(k.KabKota || ''),
      String(k.Pekerjaan || 'Umum'),
      String(k.Keluhan || ''),
      k.Sistolik !== '' && k.Sistolik !== undefined ? k.Sistolik : '',
      k.Diastolik !== '' && k.Diastolik !== undefined ? k.Diastolik : '',
      String(k.Kategori_TD || ''),
      k.Respirasi !== '' && k.Respirasi !== undefined ? k.Respirasi : '',
      k.Nadi !== '' && k.Nadi !== undefined ? k.Nadi : '',
      k.SpO2 !== '' && k.SpO2 !== undefined ? k.SpO2 : '',
      k.Suhu !== '' && k.Suhu !== undefined ? k.Suhu : '',
      k.TB !== '' && k.TB !== undefined ? k.TB : '',
      k.BB !== '' && k.BB !== undefined ? k.BB : '',
      k.IMT !== '' && k.IMT !== undefined ? k.IMT : '',
      String(k.Kategori_IMT || ''),
      k.Lingkar_Perut !== '' && k.Lingkar_Perut !== undefined ? k.Lingkar_Perut : '',
      String(k.Buta_Warna || ''),
      String(k.ICD10_Code || ''),
      String(k.Diagnosa_ICD10 || ''),
      String(k.Group_Diag || ''),
      String(k.Kategori_DXKep || ''),
      String(k.DX_Keperawatan || ''),
      String(k.FR_Ditemukan || ''),
      String(k.FR_Dikendalikan || ''),
      String(k.Kategori_Penumpang || ''),
      String(k.Dokter_Pemeriksa || ''),
      adaLabBool ? 'Ya' : 'Tidak',
      adaObatBool ? 'Ya' : 'Tidak',
      (String(k.FR_Dikendalikan || '').toLowerCase().indexOf('rujukan') !== -1) ? String(k.Lokasi_Rujukan || '').trim() : '',
      String(k.Diagnosa_Kerja_E || '').trim(),
      String(k.Diagnosa_Kerja || '').trim()
    ];

    if (shKunj.getLastColumn() < 54) {
      shKunj.getRange(1, 52, 1, 3).setValues([['Lokasi_Rujukan', 'Diagnosa_Kerja_E', 'Diagnosa_Kerja']])
        .setBackground('#0d6efd').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    }

    if (isEdit) {
      const finderKunj = shKunj.getRange(2, 1, Math.max(1, shKunj.getLastRow() - 1), 1)
        .createTextFinder(idKunjungan)
        .matchEntireCell(true);
      const cellKunj = finderKunj.findNext();
      if (!cellKunj) {
        throw new Error('Data kunjungan yang akan diedit tidak ditemukan.');
      }
      const existingWilker = shKunj.getRange(cellKunj.getRow(), 4).getValue();
      if (session.role === 'Member' && existingWilker !== session.wilker) {
        throw new Error('Akses Ditolak: Anda tidak dapat mengubah data Wilker lain.');
      }
      shKunj.getRange(cellKunj.getRow(), 1, 1, 54).setValues([rowKunjungan]);
    } else {
      shKunj.appendRow(rowKunjungan);
    }

    // 2. Proses Sheet Data_Laboratorium (47 Kolom)
    if (shLab.getLastColumn() < 47) {
      shLab.getRange(1, 46, 1, 2).setValues([['Interp_Asam_Urat', 'Jenis_Dengue']])
        .setBackground('#198754').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    }
    let savedLabObj = null;
    if (shLab.getLastRow() > 1 && isEdit) {
      const finderLab = shLab.getRange(2, 2, shLab.getLastRow() - 1, 1)
        .createTextFinder(idKunjungan)
        .matchEntireCell(true);
      const cellLab = finderLab.findNext();
      if (cellLab) {
        if (adaLabBool) {
          const statusAbnormal = evaluateLabAbnormalStatus(lab);
          const idLab = shLab.getRange(cellLab.getRow(), 1).getValue() || ('LAB-' + idKunjungan);
          const rowLab = buildLabRowArray(idLab, idKunjungan, tanggalKunjungan, effectiveWilker, noRM, k.Nama_Pasien, lab, statusAbnormal);
          shLab.getRange(cellLab.getRow(), 1, 1, 47).setValues([rowLab]);
          savedLabObj = mapRowToLabObject(rowLab);
        } else {
          shLab.deleteRow(cellLab.getRow());
        }
      } else if (adaLabBool) {
        const statusAbnormal = evaluateLabAbnormalStatus(lab);
        const idLab = 'LAB-' + idKunjungan;
        const rowLab = buildLabRowArray(idLab, idKunjungan, tanggalKunjungan, effectiveWilker, noRM, k.Nama_Pasien, lab, statusAbnormal);
        shLab.appendRow(rowLab);
        savedLabObj = mapRowToLabObject(rowLab);
      }
    } else if (adaLabBool) {
      const statusAbnormal = evaluateLabAbnormalStatus(lab);
      const idLab = 'LAB-' + idKunjungan;
      const rowLab = buildLabRowArray(idLab, idKunjungan, tanggalKunjungan, effectiveWilker, noRM, k.Nama_Pasien, lab, statusAbnormal);
      shLab.appendRow(rowLab);
      savedLabObj = mapRowToLabObject(rowLab);
    }

    // 3. Proses Sheet Data_Terapi_Obat (10 Kolom tanpa Aturan_Pakai)
    if (isEdit && shObat.getLastRow() > 1) {
      const finderObat = shObat.getRange(2, 2, shObat.getLastRow() - 1, 1)
        .createTextFinder(idKunjungan)
        .matchEntireCell(true);
      const cellsObat = finderObat.findAll();
      // Hapus dari baris paling bawah ke atas agar indeks tidak bergeser
      for (let idx = cellsObat.length - 1; idx >= 0; idx--) {
        shObat.deleteRow(cellsObat[idx].getRow());
      }
    }

    const savedObatList = [];
    if (adaObatBool) {
      const obatRowsToInsert = [];
      for (let m = 0; m < validObatItems.length; m++) {
        const item = validObatItems[m];
        const idTerapi = 'OBT-' + idKunjungan + '-' + (m + 1);
        const rObat = [
          idTerapi,
          idKunjungan,
          tanggalKunjungan,
          effectiveWilker,
          noRM,
          String(k.Nama_Pasien || '').trim(),
          String(item.Nama_Obat || '').trim(),
          String(item.Sediaan || '').trim(),
          Number(item.Jumlah || 0),
          String(k.Dokter_Pemeriksa || '')
        ];
        obatRowsToInsert.push(rObat);
        savedObatList.push(mapRowToObatObject(rObat));
      }
      const startRowObat = shObat.getLastRow() + 1;
      shObat.getRange(startRowObat, 1, obatRowsToInsert.length, 10).setValues(obatRowsToInsert);
    }

    return {
      success: true,
      isEdit: isEdit,
      idKunjungan: idKunjungan,
      noRM: noRM,
      message: (isEdit ? 'Data kunjungan berhasil diperbarui' : 'Data kunjungan baru berhasil disimpan') + ' dengan No. RM: ' + noRM,
      savedRecord: {
        kunjungan: mapRowToKunjunganObject(rowKunjungan),
        laboratorium: savedLabObj,
        terapiObat: savedObatList
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function buildLabRowArray(idLab, idKunjungan, tanggalKunjungan, wilker, noRM, namaPasien, lab, statusAbnormal) {
  return [
    idLab,
    idKunjungan,
    tanggalKunjungan,
    wilker,
    noRM,
    String(namaPasien || '').trim(),
    String(lab.Hemoglobin || ''),
    String(lab.Hematokrit || ''),
    String(lab.Leukosit || ''),
    String(lab.Platelet || ''),
    String(lab.Eritrosit || ''),
    String(lab.GDP || ''),
    String(lab.Interp_GDP || ''),
    String(lab.GDS || ''),
    String(lab.Interp_GDS || ''),
    String(lab.GD_2JPP || ''),
    String(lab.Interp_GD_2JPP || ''),
    String(lab.Kolesterol_Total || ''),
    String(lab.Interp_Kol_Total || ''),
    String(lab.LDL || ''),
    String(lab.Interp_LDL || ''),
    String(lab.HDL || ''),
    String(lab.Interp_HDL || ''),
    String(lab.Trigliserida || ''),
    String(lab.Interp_Trigliserida || ''),
    String(lab.Asam_Urat || ''),
    String(lab.Warna_Urine || ''),
    String(lab.Bau_Urine || ''),
    String(lab.Kejernihan_Urine || ''),
    String(lab.pH_Urine || ''),
    String(lab.Protein_Urine || ''),
    String(lab.Glukosa_Urine || ''),
    String(lab.Keton_Urine || ''),
    String(lab.Bilirubin_Urine || ''),
    String(lab.Urobilinogen_Urine || ''),
    String(lab.Darah_Samar_Urine || ''),
    String(lab.PPTest_Kehamilan || '-'),
    String(lab.Malaria_Test || '-'),
    String(lab.Dengue_Test || '-'),
    String(lab.HIV_Test || '-'),
    String(lab.Widal_Test || '-'),
    String(lab.RDT_Covid19 || '-'),
    String(lab.TB_Test || '-'),
    String(lab.EKG || ''),
    statusAbnormal,
    String(lab.Interp_Asam_Urat || ''),
    String(lab.Jenis_Dengue || '-')
  ];
}

/**
 * Hapus Data Kunjungan Relasional 3 Sheet (KHUSUS MASTER ADMIN)
 */
function deleteVisitTransaction(sessionToken, idKunjungan) {
  validateSession(sessionToken, ['Master Admin']);
  if (!idKunjungan) throw new Error('ID Kunjungan tidak valid.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanId = String(idKunjungan).trim();

    // 1. Hapus di Data_Kunjungan
    const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
    if (shKunj && shKunj.getLastRow() > 1) {
      const cellsK = shKunj.getRange(2, 1, shKunj.getLastRow() - 1, 1).createTextFinder(cleanId).matchEntireCell(true).findAll();
      for (let i = cellsK.length - 1; i >= 0; i--) {
        shKunj.deleteRow(cellsK[i].getRow());
      }
    }

    // 2. Hapus di Data_Laboratorium
    const shLab = ss.getSheetByName(SHEET_NAMES.LAB);
    if (shLab && shLab.getLastRow() > 1) {
      const cellsL = shLab.getRange(2, 2, shLab.getLastRow() - 1, 1).createTextFinder(cleanId).matchEntireCell(true).findAll();
      for (let j = cellsL.length - 1; j >= 0; j--) {
        shLab.deleteRow(cellsL[j].getRow());
      }
    }

    // 3. Hapus di Data_Terapi_Obat
    const shObat = ss.getSheetByName(SHEET_NAMES.OBAT);
    if (shObat && shObat.getLastRow() > 1) {
      const cellsO = shObat.getRange(2, 2, shObat.getLastRow() - 1, 1).createTextFinder(cleanId).matchEntireCell(true).findAll();
      for (let k = cellsO.length - 1; k >= 0; k--) {
        shObat.deleteRow(cellsO[k].getRow());
      }
    }

    return {
      success: true,
      message: 'Data kunjungan "' + cleanId + '" beserta relasi Lab & Terapi Obat berhasil dihapus permanen.'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * 7. DASHBOARD PUBLIK (TANPA LOGIN - AGREGASI ANONIM TANPA PII PASIEN)
 * ============================================================================
 */
function getPublicDashboardData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  if (!shKunj) {
    setupDatabase();
    shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  }
  const shLab = ss.getSheetByName(SHEET_NAMES.LAB);

  const f = filters || {};
  const startDate = f.startDate ? String(f.startDate).trim() : '';
  const endDate = f.endDate ? String(f.endDate).trim() : '';
  const filterWilker = f.wilker && f.wilker !== 'SEMUA' ? String(f.wilker).trim() : 'SEMUA';

  let totalKunjungan = 0;
  let totalPNBP = 0;
  let totalFR = 0;
  let totalLabAbnormal = 0;

  const wilkerCounts = {};
  WILKER_LIST.forEach(function(w) { wilkerCounts[w] = 0; });

  const jenisCounts = {};
  const kategoriUmurCounts = {
    'Bayi': 0,
    'Balita': 0,
    'Anak Pra-Sekolah': 0,
    'Anak Usia Sekolah & Remaja': 0,
    'Dewasa / Usia Produktif': 0,
    'Lansia': 0
  };
  const jkCounts = { 'Laki-laki': 0, 'Perempuan': 0 };
  const dailyTrendMap = {};

  if (shKunj.getLastRow() > 1) {
    const vals = shKunj.getRange(2, 1, shKunj.getLastRow() - 1, 51).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const tgl = formatDateYMD(r[2]);
      const wlk = String(r[3] || '');
      if (startDate && tgl < startDate) continue;
      if (endDate && tgl > endDate) continue;
      if (filterWilker !== 'SEMUA' && wlk !== filterWilker) continue;

      totalKunjungan++;
      totalPNBP += Number(r[17] || 0);
      const fr = String(r[45] || '').trim();
      if (fr && fr !== '-') totalFR++;

      if (wilkerCounts[wlk] !== undefined) wilkerCounts[wlk]++;
      else if (wlk) wilkerCounts[wlk] = 1;

      const jns = String(r[8] || 'Lainnya');
      jenisCounts[jns] = (jenisCounts[jns] || 0) + 1;

      const katUmur = String(r[21] || '');
      if (kategoriUmurCounts[katUmur] !== undefined) kategoriUmurCounts[katUmur]++;

      const jk = String(r[19] || '');
      if (jkCounts[jk] !== undefined) jkCounts[jk]++;

      if (tgl) {
        dailyTrendMap[tgl] = (dailyTrendMap[tgl] || 0) + 1;
      }
    }
  }

  if (shLab && shLab.getLastRow() > 1) {
    const labVals = shLab.getRange(2, 1, shLab.getLastRow() - 1, 45).getValues();
    for (let j = 0; j < labVals.length; j++) {
      const rl = labVals[j];
      const tglL = formatDateYMD(rl[2]);
      const wlkL = String(rl[3] || '');
      if (startDate && tglL < startDate) continue;
      if (endDate && tglL > endDate) continue;
      if (filterWilker !== 'SEMUA' && wlkL !== filterWilker) continue;

      const status = String(rl[44] || 'Normal');
      if (status && status !== 'Normal') totalLabAbnormal++;
    }
  }

  const sortedDates = Object.keys(dailyTrendMap).sort();
  const trendData = sortedDates.map(function(d) {
    return { date: d, count: dailyTrendMap[d] };
  });

  return {
    success: true,
    lastUpdated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    summary: {
      totalKunjungan: totalKunjungan,
      totalPNBP: totalPNBP,
      totalFR: totalFR,
      totalLabAbnormal: totalLabAbnormal
    },
    wilkerCounts: wilkerCounts,
    jenisCounts: jenisCounts,
    kategoriUmurCounts: kategoriUmurCounts,
    jkCounts: jkCounts,
    trendData: trendData
  };
}

/**
 * ============================================================================
 * 8. DASHBOARD INTERNAL INTERAKTIF, PEMANTAUAN FR & PEMANTAUAN PEMAKAIAN OBAT
 * ============================================================================
 */
function getDashboardAnalytics(sessionToken, filters) {
  const session = validateSession(sessionToken);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  const shLab = ss.getSheetByName(SHEET_NAMES.LAB);
  const shObat = ss.getSheetByName(SHEET_NAMES.OBAT);

  const f = filters || {};
  const startDate = f.startDate ? String(f.startDate).trim() : '';
  const endDate = f.endDate ? String(f.endDate).trim() : '';
  let filterWilker = f.wilker ? String(f.wilker).trim() : 'SEMUA';
  if (session.role === 'Member') {
    filterWilker = session.wilker;
  }

  let totalKunjungan = 0;
  let totalPNBP = 0;
  let totalFRDitemukan = 0;
  let totalLabAbnormal = 0;

  const wilkerMap = {};
  WILKER_LIST.forEach(function(w) { wilkerMap[w] = 0; });

  const jenisMap = {};
  const icd10Map = {};
  const umurMap = {
    'Bayi': 0,
    'Balita': 0,
    'Anak Pra-Sekolah': 0,
    'Anak Usia Sekolah & Remaja': 0,
    'Dewasa / Usia Produktif': 0,
    'Lansia': 0
  };
  const jkMap = { 'Laki-laki': 0, 'Perempuan': 0 };
  const kabKotaMap = {};
  const frDitemukanMap = {};
  const frDikendalikanMap = {};
  const frMatrixByWilker = {};
  WILKER_LIST.forEach(function(w) {
    frMatrixByWilker[w] = { ditemukan: 0, dikendalikan: 0, rincianFR: {} };
  });
  const riskyPatients = [];

  if (shKunj && shKunj.getLastRow() > 1) {
    const vals = shKunj.getRange(2, 1, shKunj.getLastRow() - 1, Math.max(52, shKunj.getLastColumn())).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const tgl = formatDateYMD(r[2]);
      const wlk = String(r[3] || '');
      if (startDate && tgl < startDate) continue;
      if (endDate && tgl > endDate) continue;
      if (filterWilker !== 'SEMUA' && wlk !== filterWilker) continue;

      totalKunjungan++;
      totalPNBP += Number(r[17] || 0);

      if (wilkerMap[wlk] !== undefined) wilkerMap[wlk]++;
      else if (wlk) wilkerMap[wlk] = 1;

      const jns = String(r[8] || 'Lainnya');
      jenisMap[jns] = (jenisMap[jns] || 0) + 1;

      const icdCode = String(r[40] || '').trim();
      const icdName = String(r[41] || '').trim();
      if (icdCode) {
        const keyICD = icdCode + ' - ' + (icdName || 'Diagnosa');
        icd10Map[keyICD] = (icd10Map[keyICD] || 0) + 1;
      }

      const katUmur = String(r[21] || '');
      if (umurMap[katUmur] !== undefined) umurMap[katUmur]++;

      const jk = String(r[19] || '');
      if (jkMap[jk] !== undefined) jkMap[jk]++;

      const kab = String(r[24] || '').trim();
      if (kab) kabKotaMap[kab] = (kabKotaMap[kab] || 0) + 1;

      const frD = String(r[45] || '').trim();
      const frK = String(r[46] || '').trim();
      if (frD && frD !== '-') {
        totalFRDitemukan++;
        frDitemukanMap[frD] = (frDitemukanMap[frD] || 0) + 1;
        if (!frMatrixByWilker[wlk]) {
          frMatrixByWilker[wlk] = { ditemukan: 0, dikendalikan: 0, rincianFR: {} };
        }
        frMatrixByWilker[wlk].ditemukan++;
        frMatrixByWilker[wlk].rincianFR[frD] = (frMatrixByWilker[wlk].rincianFR[frD] || 0) + 1;

        if (riskyPatients.length < 150) {
          riskyPatients.push({
            ID_Kunjungan: String(r[0] || ''),
            Tanggal_Kunjungan: tgl,
            Wilker: wlk,
            No_RM: String(r[6] || ''),
            No_Identitas: String(r[7] || ''),
            Nama_Pasien: String(r[18] || ''),
            JK: jk,
            Umur: safeNumericText(r[20]),
            Jenis_Kunjungan: jns,
            Diagnosa_ICD10: icdCode ? (icdCode + ' - ' + icdName) : '-',
            FR_Ditemukan: frD,
            FR_Dikendalikan: frK || '-',
            Lokasi_Rujukan: String(r[51] || ''),
            Dokter_Pemeriksa: String(r[48] || '-')
          });
        }
      }
      if (frK && frK !== '-') {
        frDikendalikanMap[frK] = (frDikendalikanMap[frK] || 0) + 1;
        if (frMatrixByWilker[wlk]) {
          frMatrixByWilker[wlk].dikendalikan++;
        }
      }
    }
  }

  // Agregasi Lab Positif & Abnormal (Sheet Data_Laboratorium)
  const labSummaryCounts = {
    'Malaria Positif': 0,
    'Dengue Positif': 0,
    'HIV Positif': 0,
    'TB Positif': 0,
    'Covid-19 Positif': 0,
    'PPTest Positif': 0,
    'Widal Sedang/Tinggi': 0,
    'Gula Darah Tinggi/DM': 0,
    'Kolesterol Tinggi': 0,
    'Asam Urat Tinggi': 0
  };

  if (shLab && shLab.getLastRow() > 1) {
    const labVals = shLab.getRange(2, 1, shLab.getLastRow() - 1, Math.max(47, shLab.getLastColumn())).getValues();
    for (let j = 0; j < labVals.length; j++) {
      const rl = labVals[j];
      const tglL = formatDateYMD(rl[2]);
      const wlkL = String(rl[3] || '');
      if (startDate && tglL < startDate) continue;
      if (endDate && tglL > endDate) continue;
      if (filterWilker !== 'SEMUA' && wlkL !== filterWilker) continue;

      const st = String(rl[44] || 'Normal');
      if (st && st !== 'Normal') totalLabAbnormal++;

      if (String(rl[37]) === 'Positif') labSummaryCounts['Malaria Positif']++;
      if (String(rl[38]).indexOf('Positif') !== -1) labSummaryCounts['Dengue Positif']++;
      if (String(rl[45] || '').indexOf('Tinggi') !== -1) labSummaryCounts['Asam Urat Tinggi']++;
      if (String(rl[39]) === 'Positif') labSummaryCounts['HIV Positif']++;
      if (String(rl[42]) === 'Positif') labSummaryCounts['TB Positif']++;
      if (String(rl[41]) === 'Positif') labSummaryCounts['Covid-19 Positif']++;
      if (String(rl[36]) === 'Positif') labSummaryCounts['PPTest Positif']++;
      if (String(rl[40]) === 'Sedang (1:160)' || String(rl[40]) === 'Tinggi (> 1:320)') labSummaryCounts['Widal Sedang/Tinggi']++;

      if (
        String(rl[12]).indexOf('Tinggi') !== -1 ||
        String(rl[14]).indexOf('Tinggi') !== -1 ||
        String(rl[16]).indexOf('Tinggi') !== -1
      ) {
        labSummaryCounts['Gula Darah Tinggi/DM']++;
      }

      if (
        String(rl[18]).indexOf('Tinggi') !== -1 ||
        String(rl[20]).indexOf('Tinggi') !== -1 ||
        String(rl[24]).indexOf('Tinggi') !== -1
      ) {
        labSummaryCounts['Kolesterol Tinggi']++;
      }
    }
  }

  // Agregasi Pemakaian Terapi Obat (Sheet Data_Terapi_Obat)
  const obatUsageMap = {};
  const obatBySediaanMap = {};
  const obatByWilkerMap = {};
  WILKER_LIST.forEach(function(w) { obatByWilkerMap[w] = 0; });
  let totalItemObatKeluar = 0;
  let totalResepTransaksi = 0;

  if (shObat && shObat.getLastRow() > 1) {
    const obatVals = shObat.getRange(2, 1, shObat.getLastRow() - 1, 10).getValues();
    for (let k = 0; k < obatVals.length; k++) {
      const ro = obatVals[k];
      const tglO = formatDateYMD(ro[2]);
      const wlkO = String(ro[3] || '');
      if (startDate && tglO < startDate) continue;
      if (endDate && tglO > endDate) continue;
      if (filterWilker !== 'SEMUA' && wlkO !== filterWilker) continue;

      const namaObat = String(ro[6] || '').trim();
      const sediaan = String(ro[7] || 'Lainnya').trim();
      const jml = Number(ro[8] || 0);
      if (!namaObat) continue;

      totalResepTransaksi++;
      totalItemObatKeluar += jml;

      if (!obatUsageMap[namaObat]) {
        obatUsageMap[namaObat] = {
          namaObat: namaObat,
          sediaan: sediaan,
          totalJumlah: 0,
          frekuensiResep: 0,
          perWilker: {}
        };
      }
      obatUsageMap[namaObat].totalJumlah += jml;
      obatUsageMap[namaObat].frekuensiResep += 1;
      obatUsageMap[namaObat].perWilker[wlkO] = (obatUsageMap[namaObat].perWilker[wlkO] || 0) + jml;

      obatBySediaanMap[sediaan] = (obatBySediaanMap[sediaan] || 0) + jml;
      if (obatByWilkerMap[wlkO] !== undefined) obatByWilkerMap[wlkO] += jml;
      else if (wlkO) obatByWilkerMap[wlkO] = jml;
    }
  }

  const top10ICD10 = sortMapToTopArray(icd10Map, 10);
  const top10KabKota = sortMapToTopArray(kabKotaMap, 10);
  const obatMonitoringList = Object.keys(obatUsageMap).map(function(key) {
    return obatUsageMap[key];
  }).sort(function(a, b) {
    return b.totalJumlah - a.totalJumlah;
  });
  const top10Obat = obatMonitoringList.slice(0, 10).map(function(item) {
    return { label: item.namaObat + ' (' + item.sediaan + ')', value: item.totalJumlah };
  });

  return {
    success: true,
    summary: {
      totalKunjungan: totalKunjungan,
      totalPNBP: totalPNBP,
      totalFRDitemukan: totalFRDitemukan,
      totalLabAbnormal: totalLabAbnormal,
      totalItemObatKeluar: totalItemObatKeluar,
      totalResepTransaksi: totalResepTransaksi,
      jenisObatAktif: obatMonitoringList.length
    },
    wilkerCounts: wilkerMap,
    jenisCounts: jenisMap,
    top10ICD10: top10ICD10,
    umurCounts: umurMap,
    jkCounts: jkMap,
    top10KabKota: top10KabKota,
    labSummaryCounts: labSummaryCounts,
    frDitemukanMap: frDitemukanMap,
    frDikendalikanMap: frDikendalikanMap,
    frMatrixByWilker: frMatrixByWilker,
    riskyPatients: riskyPatients.reverse(),
    top10Obat: top10Obat,
    obatMonitoringList: obatMonitoringList,
    obatBySediaanMap: obatBySediaanMap,
    obatByWilkerMap: obatByWilkerMap
  };
}

function sortMapToTopArray(mapObj, limit) {
  return Object.keys(mapObj)
    .map(function(k) { return { label: k, value: mapObj[k] }; })
    .sort(function(a, b) { return b.value - a.value; })
    .slice(0, limit || 10);
}

/**
 * ============================================================================
 * 9. TABEL DATA INTERAKTIF (PAGINASI SERVER `pageSize = 50`) & REKAPITULASI
 * ============================================================================
 */
function getPaginatedTableData(sessionToken, params) {
  const session = validateSession(sessionToken);
  const p = params || {};
  const tabType = p.tabType || 'KUNJUNGAN'; // KUNJUNGAN | LAB | OBAT | REKAP
  const page = Math.max(1, parseInt(p.page || 1, 10));
  const pageSize = 50;
  const startDate = p.startDate ? String(p.startDate).trim() : '';
  const endDate = p.endDate ? String(p.endDate).trim() : '';
  const jenisFilter = p.jenisKunjungan && p.jenisKunjungan !== 'SEMUA' ? String(p.jenisKunjungan).trim() : 'SEMUA';
  const searchQuery = p.search ? String(p.search).trim().toLowerCase() : '';

  let wilkerFilter = p.wilker && p.wilker !== 'SEMUA' ? String(p.wilker).trim() : 'SEMUA';
  if (session.role === 'Member') {
    wilkerFilter = session.wilker;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (tabType === 'KUNJUNGAN') {
    const sh = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
    if (!sh || sh.getLastRow() <= 1) {
      return { success: true, tabType: 'KUNJUNGAN', rows: [], totalRows: 0, page: 1, totalPages: 1 };
    }
    const numCols = Math.max(54, sh.getLastColumn());
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, numCols).getValues();
    const filtered = [];
    for (let i = vals.length - 1; i >= 0; i--) {
      const r = vals[i];
      // Skip completely blank rows
      const hasAnyContent = r.some(function(cell) { return cell !== '' && cell !== null && cell !== undefined; });
      if (!hasAnyContent) continue;

      if (!r[0]) {
        r[0] = 'ROW-' + (i + 2);
      }
      const tgl = formatDateYMD(r[2]) || formatDateYMD(r[1]);
      const wlk = String(r[3] || '').trim();
      const jns = String(r[8] || '').trim();
      if (startDate && tgl && tgl < startDate) continue;
      if (endDate && tgl && tgl > endDate) continue;
      if (wilkerFilter !== 'SEMUA' && wlk.toLowerCase() !== wilkerFilter.toLowerCase()) continue;
      if (jenisFilter !== 'SEMUA' && jns.toLowerCase() !== jenisFilter.toLowerCase()) continue;

      if (searchQuery) {
        const noRM = String(r[6] || '').toLowerCase();
        const noId = String(r[7] || '').toLowerCase();
        const nama = String(r[18] || '').toLowerCase();
        const idK = String(r[0] || '').toLowerCase();
        if (
          noRM.indexOf(searchQuery) === -1 &&
          noId.indexOf(searchQuery) === -1 &&
          nama.indexOf(searchQuery) === -1 &&
          idK.indexOf(searchQuery) === -1
        ) {
          continue;
        }
      }
      const obj = mapRowToKunjunganObject(r);
      if (!obj.Tanggal_Kunjungan && tgl) obj.Tanggal_Kunjungan = tgl;
      filtered.push(obj);
    }

    const totalRows = filtered.length;
    const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
    const sliceStart = (page - 1) * pageSize;
    return {
      success: true,
      tabType: 'KUNJUNGAN',
      rows: filtered.slice(sliceStart, sliceStart + pageSize),
      totalRows: totalRows,
      page: page,
      totalPages: totalPages
    };
  }

  if (tabType === 'LAB') {
    const shL = ss.getSheetByName(SHEET_NAMES.LAB);
    if (!shL || shL.getLastRow() <= 1) {
      return { success: true, tabType: 'LAB', rows: [], totalRows: 0, page: 1, totalPages: 1 };
    }
    const valsL = shL.getRange(2, 1, shL.getLastRow() - 1, Math.max(47, shL.getLastColumn())).getValues();
    const filteredL = [];
    for (let j = valsL.length - 1; j >= 0; j--) {
      const rl = valsL[j];
      const tgl = formatDateYMD(rl[2]);
      const wlk = String(rl[3] || '');
      if (startDate && tgl < startDate) continue;
      if (endDate && tgl > endDate) continue;
      if (wilkerFilter !== 'SEMUA' && wlk !== wilkerFilter) continue;
      if (searchQuery) {
        const noRM = String(rl[4] || '').toLowerCase();
        const nama = String(rl[5] || '').toLowerCase();
        const idK = String(rl[1] || '').toLowerCase();
        if (noRM.indexOf(searchQuery) === -1 && nama.indexOf(searchQuery) === -1 && idK.indexOf(searchQuery) === -1) {
          continue;
        }
      }
      filteredL.push(mapRowToLabObject(rl));
    }
    const totalRowsL = filteredL.length;
    const totalPagesL = Math.max(1, Math.ceil(totalRowsL / pageSize));
    const sliceStartL = (page - 1) * pageSize;
    return {
      success: true,
      tabType: 'LAB',
      rows: filteredL.slice(sliceStartL, sliceStartL + pageSize),
      totalRows: totalRowsL,
      page: page,
      totalPages: totalPagesL
    };
  }

  if (tabType === 'OBAT') {
    const shO = ss.getSheetByName(SHEET_NAMES.OBAT);
    if (!shO || shO.getLastRow() <= 1) {
      return { success: true, tabType: 'OBAT', rows: [], totalRows: 0, page: 1, totalPages: 1 };
    }
    const valsO = shO.getRange(2, 1, shO.getLastRow() - 1, 10).getValues();
    const filteredO = [];
    for (let k = valsO.length - 1; k >= 0; k--) {
      const ro = valsO[k];
      const tgl = formatDateYMD(ro[2]);
      const wlk = String(ro[3] || '');
      if (startDate && tgl < startDate) continue;
      if (endDate && tgl > endDate) continue;
      if (wilkerFilter !== 'SEMUA' && wlk !== wilkerFilter) continue;
      if (searchQuery) {
        const noRM = String(ro[4] || '').toLowerCase();
        const nama = String(ro[5] || '').toLowerCase();
        const obat = String(ro[6] || '').toLowerCase();
        if (noRM.indexOf(searchQuery) === -1 && nama.indexOf(searchQuery) === -1 && obat.indexOf(searchQuery) === -1) {
          continue;
        }
      }
      filteredO.push(mapRowToObatObject(ro));
    }
    const totalRowsO = filteredO.length;
    const totalPagesO = Math.max(1, Math.ceil(totalRowsO / pageSize));
    const sliceStartO = (page - 1) * pageSize;
    return {
      success: true,
      tabType: 'OBAT',
      rows: filteredO.slice(sliceStartO, sliceStartO + pageSize),
      totalRows: totalRowsO,
      page: page,
      totalPages: totalPagesO
    };
  }

  if (tabType === 'REKAP') {
    return generateRecapSummary(ss, startDate, endDate, wilkerFilter, jenisFilter);
  }

  return { success: true, rows: [], totalRows: 0, page: 1, totalPages: 1 };
}

function generateRecapSummary(ss, startDate, endDate, wilkerFilter, jenisFilter) {
  const sh = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  const dailyMap = {};
  const weeklyMap = {};
  const monthlyMap = {};

  if (sh && sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 51).getValues();
    for (let i = 0; i < vals.length; i++) {
      const r = vals[i];
      const tgl = formatDateYMD(r[2]);
      const wlk = String(r[3] || '');
      const jns = String(r[8] || '');
      if (!tgl) continue;
      if (startDate && tgl < startDate) continue;
      if (endDate && tgl > endDate) continue;
      if (wilkerFilter !== 'SEMUA' && wlk !== wilkerFilter) continue;
      if (jenisFilter !== 'SEMUA' && jns !== jenisFilter) continue;

      const pnbp = Number(r[17] || 0);
      const hasFR = (String(r[45] || '').trim() !== '' && String(r[45] || '').trim() !== '-') ? 1 : 0;
      const hasLab = String(r[49] || '') === 'Ya' ? 1 : 0;
      const hasObat = String(r[50] || '') === 'Ya' ? 1 : 0;

      // Harian
      if (!dailyMap[tgl]) dailyMap[tgl] = { periode: tgl, kunjungan: 0, pnbp: 0, fr: 0, lab: 0, obat: 0 };
      dailyMap[tgl].kunjungan++;
      dailyMap[tgl].pnbp += pnbp;
      dailyMap[tgl].fr += hasFR;
      dailyMap[tgl].lab += hasLab;
      dailyMap[tgl].obat += hasObat;

      // Bulanan (YYYY-MM)
      const ym = tgl.substring(0, 7);
      if (!monthlyMap[ym]) monthlyMap[ym] = { periode: ym, kunjungan: 0, pnbp: 0, fr: 0, lab: 0, obat: 0 };
      monthlyMap[ym].kunjungan++;
      monthlyMap[ym].pnbp += pnbp;
      monthlyMap[ym].fr += hasFR;
      monthlyMap[ym].lab += hasLab;
      monthlyMap[ym].obat += hasObat;

      // Mingguan (Tahun-Minggu)
      const dt = new Date(tgl + 'T00:00:00');
      const weekStr = getISOWeekString(dt);
      if (!weeklyMap[weekStr]) weeklyMap[weekStr] = { periode: weekStr, kunjungan: 0, pnbp: 0, fr: 0, lab: 0, obat: 0 };
      weeklyMap[weekStr].kunjungan++;
      weeklyMap[weekStr].pnbp += pnbp;
      weeklyMap[weekStr].fr += hasFR;
      weeklyMap[weekStr].lab += hasLab;
      weeklyMap[weekStr].obat += hasObat;
    }
  }

  return {
    success: true,
    daily: Object.keys(dailyMap).sort().reverse().map(function(k) { return dailyMap[k]; }),
    weekly: Object.keys(weeklyMap).sort().reverse().map(function(k) { return weeklyMap[k]; }),
    monthly: Object.keys(monthlyMap).sort().reverse().map(function(k) { return monthlyMap[k]; })
  };
}

function getISOWeekString(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return date.getUTCFullYear() + '-W' + ('0' + weekNo).slice(-2);
}

/**
 * ============================================================================
 * 10. EXPORT & IMPORT MASSAL EXCEL (.XLSX) VIA SHEETJS
 * ============================================================================
 */
function getExportDataMultiSheet(sessionToken, filters) {
  const session = validateSession(sessionToken);
  const f = filters || {};
  const startDate = f.startDate ? String(f.startDate).trim() : '';
  const endDate = f.endDate ? String(f.endDate).trim() : '';
  const jenisFilter = f.jenisKunjungan && f.jenisKunjungan !== 'SEMUA' ? String(f.jenisKunjungan).trim() : 'SEMUA';
  let wilkerFilter = f.wilker && f.wilker !== 'SEMUA' ? String(f.wilker).trim() : 'SEMUA';
  if (session.role === 'Member') {
    wilkerFilter = session.wilker;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
  const shLab = ss.getSheetByName(SHEET_NAMES.LAB);
  const shObat = ss.getSheetByName(SHEET_NAMES.OBAT);

  const kunjunganRows = [];
  const allowedVisitIds = {};

  if (shKunj && shKunj.getLastRow() > 1) {
    const valsK = shKunj.getRange(2, 1, shKunj.getLastRow() - 1, Math.max(54, shKunj.getLastColumn())).getValues();
    for (let i = 0; i < valsK.length; i++) {
      const r = valsK[i];
      const tgl = formatDateYMD(r[2]);
      const wlk = String(r[3] || '');
      const jns = String(r[8] || '');
      if (startDate && tgl < startDate) continue;
      if (endDate && tgl > endDate) continue;
      if (wilkerFilter !== 'SEMUA' && wlk !== wilkerFilter) continue;
      if (jenisFilter !== 'SEMUA' && jns !== jenisFilter) continue;

      const objK = mapRowToKunjunganObject(r);
      kunjunganRows.push(objK);
      allowedVisitIds[objK.ID_Kunjungan] = true;
    }
  }

  const labRows = [];
  if (shLab && shLab.getLastRow() > 1) {
    const valsL = shLab.getRange(2, 1, shLab.getLastRow() - 1, Math.max(47, shLab.getLastColumn())).getValues();
    for (let j = 0; j < valsL.length; j++) {
      const rl = valsL[j];
      const idK = String(rl[1] || '');
      if (allowedVisitIds[idK]) {
        labRows.push(mapRowToLabObject(rl));
      }
    }
  }

  const obatRows = [];
  if (shObat && shObat.getLastRow() > 1) {
    const valsO = shObat.getRange(2, 1, shObat.getLastRow() - 1, 10).getValues();
    for (let k = 0; k < valsO.length; k++) {
      const ro = valsO[k];
      const idK = String(ro[1] || '');
      if (allowedVisitIds[idK]) {
        obatRows.push(mapRowToObatObject(ro));
      }
    }
  }

  return {
    success: true,
    kunjungan: kunjunganRows,
    laboratorium: labRows,
    terapiObat: obatRows
  };
}

/**
 * Upload Massal Data Kunjungan via SheetJS (.xlsx)
 */
function importBulkVisits(sessionToken, rowsArray) {
  const session = validateSession(sessionToken, ['Master Admin', 'Admin Induk', 'Member']);
  if (!Array.isArray(rowsArray) || rowsArray.length === 0) {
    throw new Error('Tidak ada baris data kunjungan untuk diimpor.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const batchKunjungan = [];

    for (let i = 0; i < rowsArray.length; i++) {
      const item = rowsArray[i];
      if (!item || !item.Nama_Pasien) continue;

      const tgl = formatDateYMD(item.Tanggal_Kunjungan) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const wlk = (session.role === 'Member') ? session.wilker : (item.Wilker || session.wilker || 'Induk');
      const idKunjungan = item.ID_Kunjungan || ('KNJ-IMP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + (i + 1));
      const noRM = item.No_RM || ('RM-' + (WILKER_CODES[wlk] || 'BKK') + '-' + tgl.replace(/-/g, '') + '-' + ('0000' + (i + 1)).slice(-4));

      batchKunjungan.push([
        idKunjungan,
        nowStr,
        tgl,
        wlk,
        session.nama + ' (Import)',
        item.Mode_RM || 'Manual',
        noRM,
        String(item.No_Identitas || ''),
        String(item.Jenis_Kunjungan || 'Pengobatan'),
        String(item.Sub_Kunjungan || 'Pengobatan Umum'),
        String(item.Fase_Haji || ''),
        String(item.Kloter_Haji || ''),
        String(item.Vaksin_Dipilih || ''),
        Number(item.Tarif_Pemeriksaan || 0),
        Number(item.Tarif_ICV || 0),
        Number(item.Tarif_Kontraindikasi || 0),
        Number(item.Tarif_Vaksin || 0),
        Number(item.Total_PNBP || 0),
        String(item.Nama_Pasien || ''),
        String(item.JK || 'Laki-laki'),
        item.Umur !== undefined ? Number(item.Umur) : 30,
        String(item.Kategori_Umur || 'Dewasa / Usia Produktif'),
        String(item.Alamat || ''),
        String(item.Provinsi || 'Nusa Tenggara Barat'),
        String(item.KabKota || 'Kota Mataram'),
        String(item.Pekerjaan || 'Umum'),
        String(item.Keluhan || ''),
        item.Sistolik || '',
        item.Diastolik || '',
        String(item.Kategori_TD || ''),
        item.Respirasi || '',
        item.Nadi || '',
        item.SpO2 || '',
        item.Suhu || '',
        item.TB || '',
        item.BB || '',
        item.IMT || '',
        String(item.Kategori_IMT || ''),
        item.Lingkar_Perut || '',
        String(item.Buta_Warna || ''),
        String(item.ICD10_Code || 'Z00'),
        String(item.Diagnosa_ICD10 || 'General medical examination'),
        String(item.Group_Diag || 'Factors influencing health status'),
        String(item.Kategori_DXKep || ''),
        String(item.DX_Keperawatan || ''),
        String(item.FR_Ditemukan || 'Pengujian/Pemeriksaan Kesehatan'),
        String(item.FR_Dikendalikan || 'Komunikasi, Informasi dan Edukasi (KIE)'),
        String(item.Kategori_Penumpang || ''),
        String(item.Dokter_Pemeriksa || 'dr. Ferry Wardhana Sp.KP'),
        'Tidak',
        'Tidak',
        String(item.Lokasi_Rujukan || ''),
        String(item.Diagnosa_Kerja_E || ''),
        String(item.Diagnosa_Kerja || '')
      ]);
    }

    if (batchKunjungan.length === 0) {
      throw new Error('Tidak ada baris valid yang memiliki Nama_Pasien.');
    }

    if (shKunj.getLastColumn() < 54) {
      shKunj.getRange(1, 52, 1, 3).setValues([['Lokasi_Rujukan', 'Diagnosa_Kerja_E', 'Diagnosa_Kerja']])
        .setBackground('#0d6efd').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    }
    const startRow = shKunj.getLastRow() + 1;
    shKunj.getRange(startRow, 1, batchKunjungan.length, 54).setValues(batchKunjungan);

    return {
      success: true,
      importedCount: batchKunjungan.length,
      message: 'Berhasil mengimpor ' + batchKunjungan.length + ' data kunjungan secara batch.'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * 11. MANAJEMEN AKUN USER & PROTEKSI MUTLAK MASTER ADMIN (ZERO-TRUST)
 * ============================================================================
 */
function getUsersList(sessionToken) {
  // HANYA Master Admin yang diizinkan memanggil endpoint ini
  validateSession(sessionToken, ['Master Admin']);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  if (!shUsers || shUsers.getLastRow() <= 1) return [];

  const data = shUsers.getRange(2, 1, shUsers.getLastRow() - 1, 6).getValues();
  const list = [];
  for (let i = 0; i < data.length; i++) {
    const role = String(data[i][3] || '');
    list.push({
      username: String(data[i][0] || ''),
      // Jangan pernah mengirimkan password / hash ke frontend
      namaLengkap: String(data[i][2] || ''),
      role: role,
      wilker: String(data[i][4] || ''),
      statusAktif: String(data[i][5] || 'Aktif'),
      isProtectedMaster: (role === 'Master Admin')
    });
  }
  return list;
}

function saveUserAccount(sessionToken, userObj) {
  validateSession(sessionToken, ['Master Admin']);
  if (!userObj || !userObj.username || !userObj.namaLengkap || !userObj.role || !userObj.wilker) {
    throw new Error('Data akun belum lengkap.');
  }

  // Proteksi Hardcoded: Tidak boleh membuat atau mengubah akun menjadi Master Admin
  if (userObj.role === 'Master Admin') {
    throw new Error('Proteksi Sistem Mutlak: Role "Master Admin" terkunci dan tidak dapat ditambahkan atau dimodifikasi.');
  }

  const cleanUsername = String(userObj.username).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = shUsers.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === cleanUsername.toLowerCase()) {
      const existingRole = String(data[i][3] || '').trim();
      if (existingRole === 'Master Admin') {
        throw new Error('Proteksi Sistem Mutlak: Akun Master Admin tidak dapat diedit, diganti role-nya, atau dinonaktifkan oleh siapa pun.');
      }
      if (!userObj.isEdit) {
        throw new Error('Username "' + cleanUsername + '" sudah terdaftar.');
      }
      const newPassHash = (userObj.password && String(userObj.password).trim() !== '')
        ? hashPassword(String(userObj.password).trim())
        : data[i][1];

      shUsers.getRange(i + 1, 1, 1, 6).setValues([[
        cleanUsername,
        newPassHash,
        String(userObj.namaLengkap).trim(),
        String(userObj.role).trim(),
        String(userObj.wilker).trim(),
        String(userObj.statusAktif || 'Aktif').trim()
      ]]);
      return { success: true, message: 'Akun "' + cleanUsername + '" berhasil diperbarui.' };
    }
  }

  if (!userObj.password || String(userObj.password).trim().length < 5) {
    throw new Error('Password akun baru wajib diisi minimal 5 karakter.');
  }

  shUsers.appendRow([
    cleanUsername,
    hashPassword(String(userObj.password).trim()),
    String(userObj.namaLengkap).trim(),
    String(userObj.role).trim(),
    String(userObj.wilker).trim(),
    String(userObj.statusAktif || 'Aktif').trim()
  ]);

  return { success: true, message: 'Akun baru "' + cleanUsername + '" berhasil ditambahkan.' };
}

function deleteUserAccount(sessionToken, targetUsername) {
  validateSession(sessionToken, ['Master Admin']);
  if (!targetUsername) throw new Error('Username wajib diisi.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  const data = shUsers.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(targetUsername).trim().toLowerCase()) {
      const role = String(data[i][3] || '').trim();
      if (role === 'Master Admin') {
        throw new Error('Proteksi Sistem Mutlak: Akun dengan role Master Admin TIDAK DAPAT DIHAPUS.');
      }
      shUsers.deleteRow(i + 1);
      return { success: true, message: 'Akun "' + targetUsername + '" berhasil dihapus.' };
    }
  }
  throw new Error('Akun tidak ditemukan.');
}

function importBulkMemberUsers(sessionToken, usersArray) {
  validateSession(sessionToken, ['Master Admin']);
  if (!Array.isArray(usersArray) || usersArray.length === 0) {
    throw new Error('File Excel akun kosong.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsers = ss.getSheetByName(SHEET_NAMES.USERS);
  const existingData = shUsers.getDataRange().getValues();
  const existingMap = {};
  for (let i = 1; i < existingData.length; i++) {
    existingMap[String(existingData[i][0]).trim().toLowerCase()] = true;
  }

  const newRows = [];
  for (let j = 0; j < usersArray.length; j++) {
    const u = usersArray[j];
    const uname = String(u.Username || '').trim();
    const rawPass = String(u.Password || 'BkkMataram2025').trim();
    const nama = String(u.Nama_Lengkap || '').trim();
    let role = String(u.Role || 'Member').trim();
    const wilker = String(u.Wilker || 'Induk').trim();
    const status = String(u.Status_Aktif || 'Aktif').trim();

    if (!uname || !nama) continue;
    if (existingMap[uname.toLowerCase()]) continue;
    // Jangan izinkan import role Master Admin
    if (role === 'Master Admin') role = 'Member';

    existingMap[uname.toLowerCase()] = true;
    newRows.push([uname, hashPassword(rawPass), nama, role, wilker, status]);
  }

  if (newRows.length === 0) {
    throw new Error('Tidak ada akun baru untuk diimpor (atau semua username sudah ada).');
  }

  shUsers.getRange(shUsers.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
  return {
    success: true,
    importedCount: newRows.length,
    message: 'Berhasil mengimpor ' + newRows.length + ' akun pengguna baru.'
  };
}

/**
 * ============================================================================
 * 12. FITUR ARSIP TAHUNAN OTOMATIS 1-KLIK MULTI-SHEET (KHUSUS MASTER ADMIN)
 * ============================================================================
 */
function getArchiveLogs(sessionToken) {
  validateSession(sessionToken, ['Master Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shLog = ss.getSheetByName(SHEET_NAMES.LOG_ARSIP);
  if (!shLog || shLog.getLastRow() <= 1) return [];

  const vals = shLog.getRange(2, 1, shLog.getLastRow() - 1, 10).getValues();
  const logs = [];
  for (let i = vals.length - 1; i >= 0; i--) {
    const r = vals[i];
    logs.push({
      ID_Log: String(r[0] || ''),
      Timestamp: formatDateTimeYMDHMS(r[1]),
      Tahun_Arsip: String(r[2] || ''),
      Mode_Arsip: String(r[3] || ''),
      Dieksekusi_Oleh: String(r[4] || ''),
      Jumlah_Kunjungan: Number(r[5] || 0),
      Jumlah_Lab: Number(r[6] || 0),
      Jumlah_Obat: Number(r[7] || 0),
      File_ID_Arsip: String(r[8] || ''),
      URL_File_Arsip: String(r[9] || '')
    });
  }
  return logs;
}

function archiveYearlyData(sessionToken, targetYear, mode) {
  const session = validateSession(sessionToken, ['Master Admin']);
  const yearStr = String(targetYear || '').trim();
  if (!/^\d{4}$/.test(yearStr)) {
    throw new Error('Tahun arsip harus berupa 4 digit angka (contoh: 2025).');
  }
  const archiveMode = mode || 'MOVE'; // MOVE (Pindahkan & Bersihkan) atau COPY (Salin Saja)

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
    const shLab = ss.getSheetByName(SHEET_NAMES.LAB);
    const shObat = ss.getSheetByName(SHEET_NAMES.OBAT);
    const shLog = ss.getSheetByName(SHEET_NAMES.LOG_ARSIP);

    // 1. Pisahkan baris Data_Kunjungan tahun target vs tahun lainnya (tanpa deleteRow di loop)
    const kunjColCount = Math.max(52, shKunj.getLastColumn());
    const kunjHeader = shKunj.getRange(1, 1, 1, kunjColCount).getValues()[0];
    if (!kunjHeader[51]) kunjHeader[51] = 'Lokasi_Rujukan';
    const kunjToArchive = [];
    const kunjToKeep = [];
    const archivedVisitIds = {};

    if (shKunj.getLastRow() > 1) {
      const allKunj = shKunj.getRange(2, 1, shKunj.getLastRow() - 1, kunjColCount).getValues();
      for (let i = 0; i < allKunj.length; i++) {
        const r = allKunj[i];
        const tgl = formatDateYMD(r[2]);
        if (tgl.substring(0, 4) === yearStr) {
          kunjToArchive.push(r);
          archivedVisitIds[String(r[0])] = true;
        } else {
          kunjToKeep.push(r);
        }
      }
    }

    if (kunjToArchive.length === 0) {
      throw new Error('Tidak ditemukan data kunjungan pada tahun ' + yearStr + ' untuk diarsipkan.');
    }

    // 2. Pisahkan baris Data_Laboratorium
    const labHeader = shLab.getRange(1, 1, 1, 45).getValues()[0];
    const labToArchive = [];
    const labToKeep = [];
    if (shLab.getLastRow() > 1) {
      const allLab = shLab.getRange(2, 1, shLab.getLastRow() - 1, 45).getValues();
      for (let j = 0; j < allLab.length; j++) {
        const rl = allLab[j];
        const idK = String(rl[1] || '');
        const tglL = formatDateYMD(rl[2]);
        if (archivedVisitIds[idK] || tglL.substring(0, 4) === yearStr) {
          labToArchive.push(rl);
        } else {
          labToKeep.push(rl);
        }
      }
    }

    // 3. Pisahkan baris Data_Terapi_Obat
    const obatHeader = shObat.getRange(1, 1, 1, 10).getValues()[0];
    const obatToArchive = [];
    const obatToKeep = [];
    if (shObat.getLastRow() > 1) {
      const allObat = shObat.getRange(2, 1, shObat.getLastRow() - 1, 10).getValues();
      for (let k = 0; k < allObat.length; k++) {
        const ro = allObat[k];
        const idK = String(ro[1] || '');
        const tglO = formatDateYMD(ro[2]);
        if (archivedVisitIds[idK] || tglO.substring(0, 4) === yearStr) {
          obatToArchive.push(ro);
        } else {
          obatToKeep.push(ro);
        }
      }
    }

    // 4. Buat / Buka Spreadsheet Arsip `ARSIP_SIKUTE_TAHUN_[targetYear]`
    const archiveFileName = 'ARSIP_SIKUTE_TAHUN_' + yearStr;
    let archiveSS = null;
    const files = DriveApp.getFilesByName(archiveFileName);
    if (files.hasNext()) {
      archiveSS = SpreadsheetApp.open(files.next());
    } else {
      archiveSS = SpreadsheetApp.create(archiveFileName);
    }

    const shArchKunj = ensureSheetWithHeaders(archiveSS, 'Data_Kunjungan_Arsip', kunjHeader, '#0d6efd');
    const shArchLab = ensureSheetWithHeaders(archiveSS, 'Data_Laboratorium_Arsip', labHeader, '#198754');
    const shArchObat = ensureSheetWithHeaders(archiveSS, 'Data_Terapi_Obat_Arsip', obatHeader, '#0dcaf0');

    // Hapus Sheet1 default jika baru dibuat
    const defaultSheet = archiveSS.getSheetByName('Sheet1');
    if (defaultSheet && archiveSS.getSheets().length > 1) {
      try { archiveSS.deleteSheet(defaultSheet); } catch (e) {}
    }

    // Tulis secara batch ke Spreadsheet Arsip
    if (kunjToArchive.length > 0) {
      shArchKunj.getRange(shArchKunj.getLastRow() + 1, 1, kunjToArchive.length, kunjColCount).setValues(kunjToArchive);
    }
    if (labToArchive.length > 0) {
      shArchLab.getRange(shArchLab.getLastRow() + 1, 1, labToArchive.length, 45).setValues(labToArchive);
    }
    if (obatToArchive.length > 0) {
      shArchObat.getRange(shArchObat.getLastRow() + 1, 1, obatToArchive.length, 10).setValues(obatToArchive);
    }

    // 5. Jika mode === 'MOVE', bersihkan data tahun tersebut dari ketiga sheet aktif secara batch
    if (archiveMode === 'MOVE') {
      // Rewrite Data_Kunjungan
      if (shKunj.getLastRow() > 1) {
        shKunj.getRange(2, 1, shKunj.getLastRow() - 1, kunjColCount).clearContent();
      }
      if (kunjToKeep.length > 0) {
        shKunj.getRange(2, 1, kunjToKeep.length, kunjColCount).setValues(kunjToKeep);
      }

      // Rewrite Data_Laboratorium
      if (shLab.getLastRow() > 1) {
        shLab.getRange(2, 1, shLab.getLastRow() - 1, 45).clearContent();
      }
      if (labToKeep.length > 0) {
        shLab.getRange(2, 1, labToKeep.length, 45).setValues(labToKeep);
      }

      // Rewrite Data_Terapi_Obat
      if (shObat.getLastRow() > 1) {
        shObat.getRange(2, 1, shObat.getLastRow() - 1, 10).clearContent();
      }
      if (obatToKeep.length > 0) {
        shObat.getRange(2, 1, obatToKeep.length, 10).setValues(obatToKeep);
      }
    }

    // 6. Catat riwayat ke Log_Arsip
    const logId = 'ARC-' + yearStr + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const nowLog = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    shLog.appendRow([
      logId,
      nowLog,
      yearStr,
      archiveMode,
      session.nama + ' (' + session.username + ')',
      kunjToArchive.length,
      labToArchive.length,
      obatToArchive.length,
      archiveSS.getId(),
      archiveSS.getUrl()
    ]);

    return {
      success: true,
      archiveUrl: archiveSS.getUrl(),
      archiveFileName: archiveFileName,
      counts: {
        kunjungan: kunjToArchive.length,
        lab: labToArchive.length,
        obat: obatToArchive.length
      },
      message: 'Arsip Tahun ' + yearStr + ' berhasil diproses (' + kunjToArchive.length + ' Kunjungan, ' + labToArchive.length + ' Lab, ' + obatToArchive.length + ' Terapi Obat).'
    };
  } finally {
    lock.releaseLock();
  }
}


/**
 * ============================================================================
 * 11. MODE OFFLINE & SINKRONISASI OTOMATIS BATCH (`syncOfflineVisitsBatch`)
 * ============================================================================
 */

/**
 * Mengambil paket lengkap referensi & riwayat pasien terbaru untuk disimpan
 * di penyimpanan lokal perangkat (LocalStorage) sebelum petugas bertugas di lapangan/pelabuhan.
 */
function getOfflineCacheBundle(sessionToken) {
  const session = validateSession(sessionToken, ['Master Admin', 'Admin Induk', 'Member']);
  const baseRefs = getInitialReferences(sessionToken);
  return {
    success: true,
    syncedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    wilker: session.wilker,
    references: baseRefs
  };
}

/**
 * Memproses antrean transaksi kunjungan yang diinput saat Mode Offline secara batch
 * ke 3 Sheet Relasional (`Data_Kunjungan`, `Data_Laboratorium`, `Data_Terapi_Obat`)
 * lengkap dengan proteksi `LockService` dan pencegahan duplikasi (Idempotency).
 */
function syncOfflineVisitsBatch(sessionToken, offlineQueueArray) {
  const session = validateSession(sessionToken, ['Master Admin', 'Admin Induk', 'Member']);
  if (!Array.isArray(offlineQueueArray) || offlineQueueArray.length === 0) {
    return {
      success: true,
      totalReceived: 0,
      syncedCount: 0,
      failedCount: 0,
      results: [],
      message: 'Tidak ada antrean offline untuk disinkronkan.'
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shKunj = ss.getSheetByName(SHEET_NAMES.KUNJUNGAN);
    const shLab = ss.getSheetByName(SHEET_NAMES.LAB);
    const shObat = ss.getSheetByName(SHEET_NAMES.OBAT);
    const scriptCache = CacheService.getScriptCache();

    if (shKunj.getLastColumn() < 54) {
      shKunj.getRange(1, 52, 1, 3).setValues([['Lokasi_Rujukan', 'Diagnosa_Kerja_E', 'Diagnosa_Kerja']])
        .setBackground('#0d6efd').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    }
    if (shLab.getLastColumn() < 47) {
      shLab.getRange(1, 46, 1, 2).setValues([['Interp_Asam_Urat', 'Jenis_Dengue']])
        .setBackground('#198754').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    }

    const results = [];
    let syncedCount = 0;
    let failedCount = 0;

    for (let idx = 0; idx < offlineQueueArray.length; idx++) {
      const item = offlineQueueArray[idx] || {};
      const offlineId = String(item.offlineId || ('OFF-ITEM-' + idx)).trim();
      const payload = item.payload || item;

      try {
        // Cek Idempotency agar item offline yang sama tidak ter-insert dua kali jika koneksi terputus saat respons
        const idempKey = 'SIKUTE_OFFSYNC_' + offlineId;
        const prevSynced = scriptCache.get(idempKey);
        if (prevSynced) {
          const parsedPrev = JSON.parse(prevSynced);
          results.push(parsedPrev);
          syncedCount++;
          continue;
        }

        const k = payload.kunjungan || {};
        const lab = payload.laboratorium || {};
        const obatItems = Array.isArray(payload.terapiObat) ? payload.terapiObat : [];

        const effectiveWilker = (session.role === 'Member') ? session.wilker : (k.Wilker || session.wilker || 'Induk');
        const tanggalKunjungan = formatDateYMD(k.Tanggal_Kunjungan) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

        const rawIdKunj = String(k.ID_Kunjungan || '').trim();
        const isTemporaryOfflineId = !rawIdKunj || rawIdKunj.indexOf('OFF-') === 0;
        const isEdit = !isTemporaryOfflineId;

        const idKunjungan = isEdit
          ? rawIdKunj
          : ('KNJ-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + '-' + Math.floor(1000 + Math.random() * 9000));

        const tempNoRM = String(k.No_RM || '').trim();
        let noRM = tempNoRM;
        if (!noRM || noRM.indexOf('OFF-RM-') === 0 || (k.Mode_RM === 'Otomatis' && !isEdit)) {
          noRM = generateNextNoRM(shKunj, effectiveWilker, tanggalKunjungan);
        }

        const adaLabBool = hasAnyLabInput(lab);
        const validObatItems = obatItems.filter(function(ob) {
          return ob && String(ob.Nama_Obat || '').trim() !== '' && Number(ob.Jumlah || 0) > 0;
        });
        const adaObatBool = validObatItems.length > 0;

        const rowKunjungan = [
          idKunjungan,
          nowStr,
          tanggalKunjungan,
          effectiveWilker,
          session.nama + ' (' + session.username + ') [Auto-Sync]',
          k.Mode_RM || 'Otomatis',
          noRM,
          String(k.No_Identitas || '').trim(),
          String(k.Jenis_Kunjungan || ''),
          String(k.Sub_Kunjungan || ''),
          String(k.Fase_Haji || ''),
          String(k.Kloter_Haji || ''),
          String(k.Vaksin_Dipilih || ''),
          Number(k.Tarif_Pemeriksaan || 0),
          Number(k.Tarif_ICV || 0),
          Number(k.Tarif_Kontraindikasi || 0),
          Number(k.Tarif_Vaksin || 0),
          Number(k.Total_PNBP || 0),
          String(k.Nama_Pasien || '').trim(),
          String(k.JK || ''),
          k.Umur !== '' && k.Umur !== undefined ? Number(k.Umur) : '',
          String(k.Kategori_Umur || ''),
          String(k.Alamat || ''),
          String(k.Provinsi || ''),
          String(k.KabKota || ''),
          String(k.Pekerjaan || 'Umum'),
          String(k.Keluhan || ''),
          k.Sistolik !== '' && k.Sistolik !== undefined ? k.Sistolik : '',
          k.Diastolik !== '' && k.Diastolik !== undefined ? k.Diastolik : '',
          String(k.Kategori_TD || ''),
          k.Respirasi !== '' && k.Respirasi !== undefined ? k.Respirasi : '',
          k.Nadi !== '' && k.Nadi !== undefined ? k.Nadi : '',
          k.SpO2 !== '' && k.SpO2 !== undefined ? k.SpO2 : '',
          k.Suhu !== '' && k.Suhu !== undefined ? k.Suhu : '',
          k.TB !== '' && k.TB !== undefined ? k.TB : '',
          k.BB !== '' && k.BB !== undefined ? k.BB : '',
          k.IMT !== '' && k.IMT !== undefined ? k.IMT : '',
          String(k.Kategori_IMT || ''),
          k.Lingkar_Perut !== '' && k.Lingkar_Perut !== undefined ? k.Lingkar_Perut : '',
          String(k.Buta_Warna || ''),
          String(k.ICD10_Code || ''),
          String(k.Diagnosa_ICD10 || ''),
          String(k.Group_Diag || ''),
          String(k.Kategori_DXKep || ''),
          String(k.DX_Keperawatan || ''),
          String(k.FR_Ditemukan || ''),
          String(k.FR_Dikendalikan || ''),
          String(k.Kategori_Penumpang || ''),
          String(k.Dokter_Pemeriksa || ''),
          adaLabBool ? 'Ya' : 'Tidak',
          adaObatBool ? 'Ya' : 'Tidak',
          (String(k.FR_Dikendalikan || '').toLowerCase().indexOf('rujukan') !== -1) ? String(k.Lokasi_Rujukan || '').trim() : '',
          String(k.Diagnosa_Kerja_E || '').trim(),
          String(k.Diagnosa_Kerja || '').trim()
        ];

        if (isEdit && shKunj.getLastRow() > 1) {
          const cellKunj = shKunj.getRange(2, 1, shKunj.getLastRow() - 1, 1)
            .createTextFinder(idKunjungan)
            .matchEntireCell(true)
            .findNext();
          if (cellKunj) {
            shKunj.getRange(cellKunj.getRow(), 1, 1, 54).setValues([rowKunjungan]);
          } else {
            shKunj.appendRow(rowKunjungan);
          }
        } else {
          shKunj.appendRow(rowKunjungan);
        }

        // 2. Simpan / Update Data_Laboratorium
        let savedLabObj = null;
        if (isEdit && shLab.getLastRow() > 1) {
          const cellLab = shLab.getRange(2, 2, shLab.getLastRow() - 1, 1)
            .createTextFinder(idKunjungan)
            .matchEntireCell(true)
            .findNext();
          if (cellLab) {
            if (adaLabBool) {
              const statusAbnormal = evaluateLabAbnormalStatus(lab);
              const idLab = shLab.getRange(cellLab.getRow(), 1).getValue() || ('LAB-' + idKunjungan);
              const rowLab = buildLabRowArray(idLab, idKunjungan, tanggalKunjungan, effectiveWilker, noRM, k.Nama_Pasien, lab, statusAbnormal);
              shLab.getRange(cellLab.getRow(), 1, 1, 47).setValues([rowLab]);
              savedLabObj = mapRowToLabObject(rowLab);
            } else {
              shLab.deleteRow(cellLab.getRow());
            }
          } else if (adaLabBool) {
            const statusAbnormal = evaluateLabAbnormalStatus(lab);
            const idLab = 'LAB-' + idKunjungan;
            const rowLab = buildLabRowArray(idLab, idKunjungan, tanggalKunjungan, effectiveWilker, noRM, k.Nama_Pasien, lab, statusAbnormal);
            shLab.appendRow(rowLab);
            savedLabObj = mapRowToLabObject(rowLab);
          }
        } else if (adaLabBool) {
          const statusAbnormal = evaluateLabAbnormalStatus(lab);
          const idLab = 'LAB-' + idKunjungan;
          const rowLab = buildLabRowArray(idLab, idKunjungan, tanggalKunjungan, effectiveWilker, noRM, k.Nama_Pasien, lab, statusAbnormal);
          shLab.appendRow(rowLab);
          savedLabObj = mapRowToLabObject(rowLab);
        }

        // 3. Simpan / Update Data_Terapi_Obat
        if (isEdit && shObat.getLastRow() > 1) {
          const cellsObat = shObat.getRange(2, 2, shObat.getLastRow() - 1, 1)
            .createTextFinder(idKunjungan)
            .matchEntireCell(true)
            .findAll();
          for (let cIdx = cellsObat.length - 1; cIdx >= 0; cIdx--) {
            shObat.deleteRow(cellsObat[cIdx].getRow());
          }
        }

        const savedObatList = [];
        if (adaObatBool) {
          const obatRowsToInsert = [];
          for (let m = 0; m < validObatItems.length; m++) {
            const obItem = validObatItems[m];
            const idTerapi = 'OBT-' + idKunjungan + '-' + (m + 1);
            const rObat = [
              idTerapi,
              idKunjungan,
              tanggalKunjungan,
              effectiveWilker,
              noRM,
              String(k.Nama_Pasien || '').trim(),
              String(obItem.Nama_Obat || '').trim(),
              String(obItem.Sediaan || '').trim(),
              Number(obItem.Jumlah || 0),
              String(k.Dokter_Pemeriksa || '')
            ];
            obatRowsToInsert.push(rObat);
            savedObatList.push(mapRowToObatObject(rObat));
          }
          const startRowObat = shObat.getLastRow() + 1;
          shObat.getRange(startRowObat, 1, obatRowsToInsert.length, 10).setValues(obatRowsToInsert);
        }

        const itemSuccessObj = {
          offlineId: offlineId,
          success: true,
          isEdit: isEdit,
          idKunjungan: idKunjungan,
          tempNoRM: tempNoRM,
          noRM: noRM,
          namaPasien: String(k.Nama_Pasien || '').trim(),
          wilker: effectiveWilker,
          tanggalKunjungan: tanggalKunjungan,
          totalPNBP: Number(k.Total_PNBP || 0),
          syncedAt: nowStr,
          savedRecord: {
            kunjungan: mapRowToKunjunganObject(rowKunjungan),
            laboratorium: savedLabObj,
            terapiObat: savedObatList
          },
          message: 'Tersinkronisasi dengan No. RM Resmi: ' + noRM
        };

        try {
          scriptCache.put(idempKey, JSON.stringify(itemSuccessObj), 21600);
        } catch (e) {}

        results.push(itemSuccessObj);
        syncedCount++;
      } catch (itemErr) {
        failedCount++;
        results.push({
          offlineId: offlineId,
          success: false,
          namaPasien: (payload.kunjungan && payload.kunjungan.Nama_Pasien) || '-',
          message: itemErr && itemErr.message ? itemErr.message : String(itemErr)
        });
      }
    }

    return {
      success: true,
      totalReceived: offlineQueueArray.length,
      syncedCount: syncedCount,
      failedCount: failedCount,
      serverTimestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      results: results,
      message: syncedCount + ' dari ' + offlineQueueArray.length + ' data kunjungan offline berhasil disinkronkan ke Google Sheets.'
    };
  } finally {
    lock.releaseLock();
  }
}
