/**
 * ==================================================================
 * FILE: Code_Report.gs
 * FUNGSI: Menarik data transaksi untuk direkapitulasi oleh Frontend
 * PERBAIKAN:
 *   - Sesuaikan jumlah kolom Trx Tabungan Aset dari 9 menjadi 10
 *     (karena sekarang ada kolom ID_Referensi di kolom J)
 * ==================================================================
 */

function getDataLaporan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataReport = {
    klinik: [],
    aset: [],
    kas: []
  };

  // 1. Tarik Data Fee Klinik (10 kolom: A-J)
  var sheetKlinik = ss.getSheetByName("Trx Fee Klinik");
  if (sheetKlinik && sheetKlinik.getLastRow() > 1) {
    dataReport.klinik = sheetKlinik
      .getRange(2, 1, sheetKlinik.getLastRow() - 1, 10)
      .getDisplayValues();
  }

  // 2. Tarik Data Tabungan Aset (10 kolom: A-J, termasuk ID_Referensi)
  var sheetAset = ss.getSheetByName("Trx Tabungan Aset");
  if (sheetAset && sheetAset.getLastRow() > 1) {
    dataReport.aset = sheetAset
      .getRange(2, 1, sheetAset.getLastRow() - 1, 10)
      .getDisplayValues();
  }

  // 3. Tarik Data Arus Kas / General Ledger (8 kolom: A-H)
  var sheetKas = ss.getSheetByName("Trx Arus Kas");
  if (sheetKas && sheetKas.getLastRow() > 1) {
    dataReport.kas = sheetKas
      .getRange(2, 1, sheetKas.getLastRow() - 1, 8)
      .getDisplayValues();
  }

  return dataReport;
}