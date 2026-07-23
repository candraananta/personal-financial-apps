/**
 * ==================================================================
 * FILE: Code_Master.gs
 * FUNGSI: Logika CRUD untuk semua tabel Master
 * ==================================================================
 */

// Fungsi dinamis untuk membaca data dari tabel Master apapun
function bacaDataMaster(namaSheet, jumlahKolom) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, jumlahKolom).getValues();
}

// Fungsi dinamis untuk menghapus baris berdasarkan ID di tabel Master apapun
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

// Fungsi Spesifik: Simpan Master Lokasi
function simpanMasterLokasi(nama, komisi, cutoff, tarif) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Lokasi");
  sheet.appendRow([Utilities.getUuid(), nama, parseFloat(komisi), parseInt(cutoff), parseFloat(tarif)]);
  return "Lokasi berhasil ditambahkan!";
}

// Fungsi Spesifik: Simpan Master Kategori
function simpanMasterKategori(nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Kategori");
  sheet.appendRow([Utilities.getUuid(), nama]);
  return "Kategori berhasil ditambahkan!";
}

// Fungsi Spesifik: Simpan Master Platform
function simpanMasterPlatform(nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Platform");
  sheet.appendRow([Utilities.getUuid(), nama]);
  return "Platform berhasil ditambahkan!";
}

// Fungsi Spesifik: Simpan Master Item
function simpanMasterItem(kategori, nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Item");
  sheet.appendRow([Utilities.getUuid(), kategori, nama]);
  return "Item berhasil ditambahkan!";
}

// Fungsi Spesifik: Simpan Master Pengeluaran Rutin
function simpanMasterPengeluaran(nama, nominal, catatan) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Pengeluaran Rutin");
  sheet.appendRow([Utilities.getUuid(), nama, parseFloat(nominal), catatan]);
  return "Master pengeluaran berhasil ditambahkan!";
}