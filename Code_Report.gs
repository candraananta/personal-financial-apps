/**
 * ==================================================================
 * FILE: Code_Report.gs
 * FUNGSI: Menarik data transaksi untuk direkapitulasi oleh Frontend
 * ==================================================================
 */

function getDataLaporan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataReport = {
    klinik: [],
    aset: [],
    kas: [] // Tambahan keranjang penampung untuk data Arus Kas
  };
  
  // 1. Tarik Data Fee Klinik
  var sheetKlinik = ss.getSheetByName("Trx Fee Klinik");
  if (sheetKlinik && sheetKlinik.getLastRow() > 1) {
    dataReport.klinik = sheetKlinik.getRange(2, 1, sheetKlinik.getLastRow() - 1, 10).getDisplayValues();
  }
  
  // 2. Tarik Data Tabungan Aset
  var sheetAset = ss.getSheetByName("Trx Tabungan Aset");
  if (sheetAset && sheetAset.getLastRow() > 1) {
    dataReport.aset = sheetAset.getRange(2, 1, sheetAset.getLastRow() - 1, 10).getDisplayValues();
  }

  // 3. Tarik Data Arus Kas (BARU)
  var sheetKas = ss.getSheetByName("Trx Arus Kas");
  if (sheetKas && sheetKas.getLastRow() > 1) {
    // Mengambil 8 kolom sesuai dengan struktur tabel Arus Kas
    dataReport.kas = sheetKas.getRange(2, 1, sheetKas.getLastRow() - 1, 8).getDisplayValues();
  }
  
  return dataReport;
}