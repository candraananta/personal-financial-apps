/**
 * ==================================================================
 * FILE: Code_Master.gs
 * FUNGSI: Logika CRUD untuk semua tabel Master
 * UPDATE: Tambah fungsi CRUD Master COA Pemasukan
 * ==================================================================
 */

// Fungsi dinamis untuk membaca data dari tabel Master apapun
function bacaDataMaster(namaSheet, jumlahKolom) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, jumlahKolom).getValues();
}

// Fungsi dinamis untuk menghapus baris berdasarkan ID
function hapusDataMaster(namaSheet, idTarget) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === idTarget) {
      sheet.deleteRow(i + 1);
      return "Data berhasil dihapus!";
    }
  }
  return "Gagal: ID tidak ditemukan.";
}

/**
 * getReferensiTransaksi()
 * Dipanggil oleh initModulTransaksi() di View_Transaksi.html
 */
function getReferensiTransaksi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = { lokasi: [], kategori: [], platform: [], item: [] };

  var sheetLokasi = ss.getSheetByName("Master Lokasi");
  if (sheetLokasi && sheetLokasi.getLastRow() > 1) {
    result.lokasi = sheetLokasi.getRange(2, 1, sheetLokasi.getLastRow() - 1, 5).getValues();
  }
  var sheetKategori = ss.getSheetByName("Master Kategori");
  if (sheetKategori && sheetKategori.getLastRow() > 1) {
    result.kategori = sheetKategori.getRange(2, 1, sheetKategori.getLastRow() - 1, 2).getValues();
  }
  var sheetPlatform = ss.getSheetByName("Master Platform");
  if (sheetPlatform && sheetPlatform.getLastRow() > 1) {
    result.platform = sheetPlatform.getRange(2, 1, sheetPlatform.getLastRow() - 1, 2).getValues();
  }
  var sheetItem = ss.getSheetByName("Master Item");
  if (sheetItem && sheetItem.getLastRow() > 1) {
    result.item = sheetItem.getRange(2, 1, sheetItem.getLastRow() - 1, 3).getValues();
  }
  return result;
}

// ============================================================
// FUNGSI COA PEMASUKAN — BARU
// ============================================================

/**
 * getReferensiCOAPemasukan()
 * Dipanggil oleh initModulArusKas() di View_ArusKas.html
 * Return: array of [ID_COA, Nama_COA, Kategori]
 */
function getReferensiCOAPemasukan() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master COA Pemasukan");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
}

/**
 * simpanCOAPemasukan()
 * Dipanggil dari form tambah di View_Master.html
 */
function simpanCOAPemasukan(kategori, nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master COA Pemasukan");
  if (!sheet) throw new Error("Sheet 'Master COA Pemasukan' tidak ditemukan! Jalankan setupDatabase() terlebih dahulu.");
  var namaLengkap = kategori + ": " + nama;
  sheet.appendRow([Utilities.getUuid(), namaLengkap, kategori]);
  return "COA Pemasukan '" + namaLengkap + "' berhasil ditambahkan!";
}

// ============================================================
// Fungsi Spesifik CRUD Master lainnya
// ============================================================

function simpanMasterLokasi(nama, komisi, cutoff, tarif) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Lokasi");
  sheet.appendRow([Utilities.getUuid(), nama, parseFloat(komisi), parseInt(cutoff), parseFloat(tarif) || 0]);
  return "Lokasi berhasil ditambahkan!";
}

function simpanMasterKategori(nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Kategori");
  sheet.appendRow([Utilities.getUuid(), nama]);
  return "Kategori berhasil ditambahkan!";
}

function simpanMasterPlatform(nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Platform");
  sheet.appendRow([Utilities.getUuid(), nama]);
  return "Platform berhasil ditambahkan!";
}

function simpanMasterItem(kategori, nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Item");
  sheet.appendRow([Utilities.getUuid(), kategori, nama]);
  return "Item berhasil ditambahkan!";
}

function simpanMasterPengeluaran(nama, nominal, catatan) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Pengeluaran Rutin");
  sheet.appendRow([Utilities.getUuid(), nama, parseFloat(nominal), catatan || '-']);
  return "Master pengeluaran berhasil ditambahkan!";
}