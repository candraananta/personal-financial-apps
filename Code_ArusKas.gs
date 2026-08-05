/**
 * ==================================================================
 * FILE: Code_ArusKas.gs
 * FUNGSI: Referensi data untuk modul Arus Kas
 * CATATAN: Tidak ada perubahan dari versi sebelumnya.
 *   - getReferensiArusKas() → template pengeluaran rutin
 *   - simpanTrxArusKas() ada di Code_Transaksi.gs (jangan duplikat)
 *   - getReferensiCOAPemasukan() ada di Code_Master.gs
 * ==================================================================
 */

/**
 * Mengambil daftar template pengeluaran rutin dari Master.
 * Dipanggil oleh:
 *   1. initModulArusKas()   → View_ArusKas.html (dropdown template pengeluaran)
 *   2. initModulTransaksi() → View_Transaksi.html (dropdown COA aset)
 *
 * Mengembalikan array of [ID, Nama, Nominal, Catatan]
 */
function getReferensiArusKas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Pengeluaran Rutin");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
}

/**
 * CATATAN ARSITEKTUR:
 * - simpanTrxArusKas(data) ada di Code_Transaksi.gs — JANGAN duplikat di sini.
 *   Apps Script V8 akan error "Already defined" jika ada 2 fungsi dengan nama sama.
 * - getReferensiCOAPemasukan() ada di Code_Master.gs
 */

/**
 * getRecentArusKas()
 * Mengambil 15 transaksi terbaru dari sheet Trx Arus Kas
 * Diurutkan berdasarkan baris terakhir masuk (newest first)
 */
function getRecentArusKas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trx Arus Kas");
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var lastRow = sheet.getLastRow();
  var startRow = Math.max(2, lastRow - 14); // ambil max 15 baris terakhir
  var numRows = lastRow - startRow + 1;

  var data = sheet.getRange(startRow, 1, numRows, 8).getValues();

  // Balik urutan: terbaru di atas
  data.reverse();

  return data.map(function(row) {
    var tgl = row[1];
    if (tgl instanceof Date) {
      var dd = String(tgl.getDate()).padStart(2, '0');
      var mm = String(tgl.getMonth() + 1).padStart(2, '0');
      var yy = tgl.getFullYear();
      tgl = dd + '/' + mm + '/' + yy;
    }
    return [
      String(row[0]),        // [0] ID
      String(tgl || '-'),    // [1] Tanggal (DD/MM/YYYY)
      String(row[2] || '-'), // [2] Periode
      String(row[3] || '-'), // [3] Tipe (PEMASUKAN/PENGELUARAN)
      String(row[4] || '-'), // [4] Item/Akun
      String(row[5] || '-'), // [5] Catatan
      parseFloat(row[6]) || 0, // [6] Debit (masuk)
      parseFloat(row[7]) || 0  // [7] Kredit (keluar)
    ];
  });
}