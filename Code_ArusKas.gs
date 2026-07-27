/**
 * ==================================================================
 * FILE: Code_ArusKas.gs
 * FUNGSI: Referensi data untuk modul Arus Kas
 * PERBAIKAN:
 *   - Hapus simpanTrxArusKas() yang duplikat [BUG #6 FIX]
 *     → Fungsi resmi ada di Code_Transaksi.gs
 *   - getReferensiArusKas() tetap di sini sebagai sumber dropdown
 *     template pengeluaran rutin
 * ==================================================================
 */

/**
 * Mengambil daftar template pengeluaran rutin dari Master.
 * Dipanggil oleh:
 *   1. initModulArusKas()    → View_ArusKas.html (dropdown template)
 *   2. initModulTransaksi()  → View_Transaksi.html (dropdown COA aset)
 *
 * Mengembalikan array of [ID, Nama, Nominal, Catatan]
 */
function getReferensiArusKas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Pengeluaran Rutin");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  // Kolom A=ID, B=Nama, C=Nominal, D=Catatan
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
}

/**
 * CATATAN ARSITEKTUR:
 * simpanTrxArusKas(data) ada di Code_Transaksi.gs — JANGAN duplikat di sini.
 * Apps Script V8 akan error "Already defined" jika ada 2 fungsi dengan nama sama.
 */