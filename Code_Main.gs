/**
 * ==================================================================
 * FILE: Code_Main.gs
 * FUNGSI: Mesin Routing, Templating, dan Setup Database
 * PERBAIKAN:
 *   - Tambah setup "Master Pengeluaran Rutin"   [BUG #4 FIX]
 *   - Tambah setup "Trx Arus Kas"               [BUG #8 FIX]
 *   - Tambah setup "Trx Fee Klinik" yang benar  [BUG #5 FIX]
 * ==================================================================
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ERP Finansial Pribadi')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// FUNGSI MESIN TEMPLATING
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * FUNGSI PEMBUAT DATABASE OTOMATIS
 * Jalankan fungsi ini SEKALI dari Apps Script Editor (▶ Run)
 * untuk membuat semua sheet yang diperlukan.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── MASTER TABLES ──────────────────────────────────────────────

  // 1. Master Lokasi (Klinik)
  var sheetLokasi = ss.getSheetByName("Master Lokasi");
  if (!sheetLokasi) {
    sheetLokasi = ss.insertSheet("Master Lokasi");
    sheetLokasi.appendRow(["ID_Lokasi", "Nama_Lokasi", "Komisi_Persen", "Tanggal_Cutoff", "Tarif_BPJS"]);
    sheetLokasi.getRange("A1:E1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
    sheetLokasi.setColumnWidth(1, 250); // UUID column lebar
  }

  // 2. Master Kategori
  var sheetKategori = ss.getSheetByName("Master Kategori");
  if (!sheetKategori) {
    sheetKategori = ss.insertSheet("Master Kategori");
    sheetKategori.appendRow(["ID_Kategori", "Nama_Kategori"]);
    sheetKategori.getRange("A1:B1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
    sheetKategori.setColumnWidth(1, 250);
  }

  // 3. Master Platform
  var sheetPlatform = ss.getSheetByName("Master Platform");
  if (!sheetPlatform) {
    sheetPlatform = ss.insertSheet("Master Platform");
    sheetPlatform.appendRow(["ID_Platform", "Nama_Platform"]);
    sheetPlatform.getRange("A1:B1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
    sheetPlatform.setColumnWidth(1, 250);
  }

  // 4. Master Item
  var sheetItem = ss.getSheetByName("Master Item");
  if (!sheetItem) {
    sheetItem = ss.insertSheet("Master Item");
    sheetItem.appendRow(["ID_Item", "Nama_Kategori", "Nama_Item"]);
    sheetItem.getRange("A1:C1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
    sheetItem.setColumnWidth(1, 250);
  }

  // 5. [BUG #4 FIX] Master Pengeluaran Rutin — SEBELUMNYA TIDAK DI-SETUP!
  var sheetPengeluaran = ss.getSheetByName("Master Pengeluaran Rutin");
  if (!sheetPengeluaran) {
    sheetPengeluaran = ss.insertSheet("Master Pengeluaran Rutin");
    sheetPengeluaran.appendRow(["ID_Template", "Nama_Tagihan", "Estimasi_Nominal", "Catatan"]);
    sheetPengeluaran.getRange("A1:D1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
    sheetPengeluaran.setColumnWidth(1, 250);
  }

  // ── TRANSACTION TABLES ─────────────────────────────────────────

  // 6. Trx Fee Klinik (10 kolom)
  var sheetTrxKlinik = ss.getSheetByName("Trx Fee Klinik");
  if (!sheetTrxKlinik) {
    sheetTrxKlinik = ss.insertSheet("Trx Fee Klinik");
    sheetTrxKlinik.appendRow([
      "ID_Trx_Klinik",
      "Tanggal_Input",
      "Periode_Komisi",
      "Klinik",
      "Nama_Pasien",
      "Tindakan",
      "Omset",
      "Komisi_Persen",
      "Nominal_Komisi",
      "Laba_Bersih"
    ]);
    sheetTrxKlinik.getRange("A1:J1").setFontWeight("bold").setBackground("#881337").setFontColor("white");
    sheetTrxKlinik.setColumnWidth(1, 250);
  }

  // 7. Trx Tabungan Aset (10 kolom)
  var sheetTrxAset = ss.getSheetByName("Trx Tabungan Aset");
  if (!sheetTrxAset) {
    sheetTrxAset = ss.insertSheet("Trx Tabungan Aset");
    sheetTrxAset.appendRow([
      "ID_Trx_Aset",
      "Tanggal_Perolehan",
      "Platform",
      "Kategori",
      "Nama_Item",
      "Tipe_Transaksi",
      "Kuantitas",
      "Harga_Satuan",
      "Total_Nilai",
      "ID_Referensi"
    ]);
    sheetTrxAset.getRange("A1:J1").setFontWeight("bold").setBackground("#881337").setFontColor("white");
    sheetTrxAset.setColumnWidth(1, 250);
  }

  // 8. [BUG #8 FIX] Trx Arus Kas (General Ledger) — SEBELUMNYA TIDAK DI-SETUP!
  var sheetTrxKas = ss.getSheetByName("Trx Arus Kas");
  if (!sheetTrxKas) {
    sheetTrxKas = ss.insertSheet("Trx Arus Kas");
    sheetTrxKas.appendRow([
      "ID_Jurnal",
      "Tanggal_Transaksi",
      "Periode_Pembukuan",
      "Tipe",         // PEMASUKAN / PENGELUARAN
      "Akun_Item",    // COA / nama kategori
      "Catatan",
      "Debit",        // Uang Masuk
      "Kredit"        // Uang Keluar
    ]);
    sheetTrxKas.getRange("A1:H1").setFontWeight("bold").setBackground("#1e3a5f").setFontColor("white");
    sheetTrxKas.setColumnWidth(1, 250);
    sheetTrxKas.setColumnWidth(3, 120);
  }

  // Tampilkan konfirmasi
  SpreadsheetApp.getUi().alert(
    '✅ Setup Database Selesai!\n\n' +
    'Sheet yang dibuat:\n' +
    '• Master Lokasi\n' +
    '• Master Kategori\n' +
    '• Master Platform\n' +
    '• Master Item\n' +
    '• Master Pengeluaran Rutin  ← BARU\n' +
    '• Trx Fee Klinik\n' +
    '• Trx Tabungan Aset\n' +
    '• Trx Arus Kas  ← BARU (General Ledger)\n\n' +
    'Silakan isi data Master terlebih dahulu.'
  );
}