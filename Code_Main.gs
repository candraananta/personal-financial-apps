/**
 * ==================================================================
 * FILE: Code_Main.gs
 * FUNGSI: Mesin Routing, Templating, dan Setup Database
 * ==================================================================
 */

function doGet() {
  // Menggunakan createTemplateFromFile agar fitur include() berfungsi
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ERP Finansial Pribadi')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// FUNGSI MESIN TEMPLATING (Untuk memanggil file HTML lain)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// FUNGSI PEMBUAT DATABASE OTOMATIS
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Master Lokasi (Klinik)
  var sheetLokasi = ss.getSheetByName("Master Lokasi");
  if (!sheetLokasi) {
    sheetLokasi = ss.insertSheet("Master Lokasi");
    sheetLokasi.appendRow(["ID_Lokasi", "Nama_Lokasi", "Komisi_Persen", "Tanggal_Cutoff", "Tarif_BPJS"]);
    sheetLokasi.getRange("A1:E1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
  }

  // 2. Setup Master Kategori
  var sheetKategori = ss.getSheetByName("Master Kategori");
  if (!sheetKategori) {
    sheetKategori = ss.insertSheet("Master Kategori");
    sheetKategori.appendRow(["ID_Kategori", "Nama_Kategori"]);
    sheetKategori.getRange("A1:B1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
  }

  // 3. Setup Master Platform
  var sheetPlatform = ss.getSheetByName("Master Platform");
  if (!sheetPlatform) {
    sheetPlatform = ss.insertSheet("Master Platform");
    sheetPlatform.appendRow(["ID_Platform", "Nama_Platform"]);
    sheetPlatform.getRange("A1:B1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
  }

  // 4. Setup Master Item
  var sheetItem = ss.getSheetByName("Master Item");
  if (!sheetItem) {
    sheetItem = ss.insertSheet("Master Item");
    sheetItem.appendRow(["ID_Item", "Nama_Kategori", "Nama_Item"]);
    sheetItem.getRange("A1:C1").setFontWeight("bold").setBackground("#be185d").setFontColor("white");
  }

  // 5. Setup Trx Fee Klinik
  var sheetTrxKlinik = ss.getSheetByName("Trx Fee Klinik");
  if (!sheetTrxKlinik) {
    sheetTrxKlinik = ss.insertSheet("Trx Fee Klinik");
    sheetTrxKlinik.appendRow(["ID_Trx_Klinik", "Tanggal_Input", "Periode_Komisi", "Klinik", "Nama_Pasien", "Tindakan", "Omset", "Komisi_Persen", "Nominal_Komisi", "Laba_Bersih"]);
    sheetTrxKlinik.getRange("A1:J1").setFontWeight("bold").setBackground("#881337").setFontColor("white");
  }

  // 6. Setup Trx Tabungan Aset
  var sheetTrxAset = ss.getSheetByName("Trx Tabungan Aset");
  if (!sheetTrxAset) {
    sheetTrxAset = ss.insertSheet("Trx Tabungan Aset");
    sheetTrxAset.appendRow(["ID_Trx_Aset", "Tanggal_Perolehan", "Platform", "Kategori", "Nama_Item", "Tipe_Transaksi", "Kuantitas", "Harga_Satuan", "Total_Nilai"]);
    sheetTrxAset.getRange("A1:I1").setFontWeight("bold").setBackground("#881337").setFontColor("white");
  }
}