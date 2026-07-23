/**
 * ==================================================================
 * FILE: Code_ArusKas.gs
 * FUNGSI: Mengelola data Pemasukan dan Pengeluaran
 * ==================================================================
 */

function getReferensiArusKas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Pengeluaran Rutin");
  if (!sheet || sheet.getLastRow() <= 1) return [];
  // Ambil Kolom ID, Nama, Nominal, Catatan
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); 
}

function simpanTrxArusKas(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trx Arus Kas");
  
  // Logika Cut-off Dinamis
  var parts = data.tanggal.split('-');
  var tahun = parseInt(parts[0]);
  var bulan = parseInt(parts[1]) - 1; 
  var hari = parseInt(parts[2]);
  var cutoff = parseInt(data.cutoff);

  // Jika tanggal transaksi melewati cut-off, masuk pembukuan bulan depan
  if (hari > cutoff) {
    bulan = bulan + 1;
    if (bulan > 11) { bulan = 0; tahun = tahun + 1; }
  }

  var namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  var periodeFinal = namaBulan[bulan] + " " + tahun; 
  
  // Tentukan Debit (Masuk) atau Kredit (Keluar)
  var debit = data.tipe === "PEMASUKAN" ? parseFloat(data.nominal) : 0;
  var kredit = data.tipe === "PENGELUARAN" ? parseFloat(data.nominal) : 0;
  
  sheet.appendRow([
    Utilities.getUuid(), 
    data.tanggal, 
    periodeFinal, 
    data.tipe, 
    data.item, 
    data.catatan, 
    debit, 
    kredit
  ]);
  
  return "Tercatat di periode: " + periodeFinal;
}