/**
 * ==================================================================
 * FILE: Code_Main.gs
 * FUNGSI: Mesin Routing, Templating, dan Setup Database
 * UPDATE: Tambah setup "Master COA Pemasukan" & Perbaikan UI Alert
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
    sheetLokasi.setColumnWidth(1, 250);
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

  // 5. Master Pengeluaran Rutin
  var sheetPengeluaran = ss.getSheetByName("Master Pengeluaran Rutin");
  if (!sheetPengeluaran) {
    sheetPengeluaran = ss.insertSheet("Master Pengeluaran Rutin");
    sheetPengeluaran.appendRow(["ID_Template", "Nama_Tagihan", "Estimasi_Nominal", "Catatan"]);
    sheetPengeluaran.getRange("A1:D1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
    sheetPengeluaran.setColumnWidth(1, 250);
  }

  // 6. Master COA Pemasukan — BARU
  var sheetCOAMasuk = ss.getSheetByName("Master COA Pemasukan");
  if (!sheetCOAMasuk) {
    sheetCOAMasuk = ss.insertSheet("Master COA Pemasukan");
    sheetCOAMasuk.appendRow(["ID_COA", "Nama_COA", "Kategori"]);
    sheetCOAMasuk.getRange("A1:C1").setFontWeight("bold").setBackground("#16a34a").setFontColor("white");
    sheetCOAMasuk.setColumnWidth(1, 250);
    sheetCOAMasuk.setColumnWidth(2, 260);
    sheetCOAMasuk.setColumnWidth(3, 180);

    // Seed data awal — bisa dihapus/tambah sesuai kebutuhan
    var seedData = [
      ["Pendapatan Aktif",  "Gaji Bulanan"],
      ["Pendapatan Aktif",  "Honor / Fee Profesional"],
      ["Pendapatan Aktif",  "Bonus & THR"],
      ["Pendapatan Aktif",  "Konsultasi / Praktek Mandiri"],
      ["Pendapatan Aktif",  "Jaga IGD / On-Call"],
      ["Pendapatan Pasif",  "Dividen Saham"],
      ["Pendapatan Pasif",  "Bunga Deposito / Tabungan"],
      ["Pendapatan Pasif",  "Hasil Sewa Properti"],
      ["Pendapatan Pasif",  "Bagi Hasil Reksadana"],
      ["Penerimaan Lain",   "Cashback / Reward / Poin"],
      ["Penerimaan Lain",   "Hibah / Hadiah / Arisan"],
      ["Penerimaan Lain",   "Pengembalian Dana (Refund)"],
      ["Transfer Masuk",    "Transfer Antar Rekening Sendiri"],
    ];
    seedData.forEach(function(row) {
      sheetCOAMasuk.appendRow([
        Utilities.getUuid(),
        row[0] + ": " + row[1],
        row[0]
      ]);
    });
  }

  // ── TRANSACTION TABLES ─────────────────────────────────────────

  // 7. Trx Fee Klinik (10 kolom)
  var sheetTrxKlinik = ss.getSheetByName("Trx Fee Klinik");
  if (!sheetTrxKlinik) {
    sheetTrxKlinik = ss.insertSheet("Trx Fee Klinik");
    sheetTrxKlinik.appendRow([
      "ID_Trx_Klinik", "Tanggal_Input", "Periode_Komisi",
      "Klinik", "Nama_Pasien", "Tindakan",
      "Omset", "Komisi_Persen", "Nominal_Komisi", "Laba_Bersih"
    ]);
    sheetTrxKlinik.getRange("A1:J1").setFontWeight("bold").setBackground("#881337").setFontColor("white");
    sheetTrxKlinik.setColumnWidth(1, 250);
  }

  // 8. Trx Tabungan Aset (10 kolom)
  var sheetTrxAset = ss.getSheetByName("Trx Tabungan Aset");
  if (!sheetTrxAset) {
    sheetTrxAset = ss.insertSheet("Trx Tabungan Aset");
    sheetTrxAset.appendRow([
      "ID_Trx_Aset", "Tanggal_Perolehan", "Platform",
      "Kategori", "Nama_Item", "Tipe_Transaksi",
      "Kuantitas", "Harga_Satuan", "Total_Nilai", "ID_Referensi"
    ]);
    sheetTrxAset.getRange("A1:J1").setFontWeight("bold").setBackground("#881337").setFontColor("white");
    sheetTrxAset.setColumnWidth(1, 250);
  }

  // 9. Trx Arus Kas (General Ledger) (8 kolom)
  var sheetTrxKas = ss.getSheetByName("Trx Arus Kas");
  if (!sheetTrxKas) {
    sheetTrxKas = ss.insertSheet("Trx Arus Kas");
    sheetTrxKas.appendRow([
      "ID_Jurnal", "Tanggal_Transaksi", "Periode_Pembukuan",
      "Tipe",       // PEMASUKAN / PENGELUARAN
      "Akun_Item",  // COA / nama kategori
      "Catatan",
      "Debit",      // Uang Masuk
      "Kredit"      // Uang Keluar
    ]);
    sheetTrxKas.getRange("A1:H1").setFontWeight("bold").setBackground("#1e3a5f").setFontColor("white");
    sheetTrxKas.setColumnWidth(1, 250);
    sheetTrxKas.setColumnWidth(3, 120);
  }

  // Tampilkan konfirmasi menggunakan Logger (Anti-Error)
  Logger.log(
    '✅ Setup Database Selesai!\n\n' +
    'Sheet yang dibuat/diverifikasi:\n' +
    '• Master Lokasi\n' +
    '• Master Kategori\n' +
    '• Master Platform\n' +
    '• Master Item\n' +
    '• Master Pengeluaran Rutin\n' +
    '• Master COA Pemasukan  ← BARU\n' +
    '• Trx Fee Klinik\n' +
    '• Trx Tabungan Aset\n' +
    '• Trx Arus Kas (General Ledger)\n\n' +
    'Master COA Pemasukan sudah diisi 13 akun awal.\n' +
    'Silakan isi data Master lainnya terlebih dahulu.'
  );
}
