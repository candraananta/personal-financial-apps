/**
 * ==================================================================
 * FILE: Code_Transaksi.gs
 * FUNGSI: Mesin Backend untuk Menyimpan Transaksi & Menghitung Periode
 * ==================================================================
 */

/**
 * FUNGSI PEMBANTU (SMART CUT-OFF ANTI-TIMEZONE BUG):
 * Memecah string tanggal secara langsung tanpa new Date() agar kebal error zona waktu.
 * Aturan: Jika tanggal >= 28, otomatis masuk ke bulan berikutnya.
 */
function hitungPeriodeSmart(tanggalInput) {
  if (!tanggalInput) return "-";
  
  // Memecah format YYYY-MM-DD langsung menjadi array angka [2026, 07, 28]
  var parts = String(tanggalInput).split('-');
  if (parts.length < 3) return tanggalInput;
  
  var tahun = parseInt(parts[0], 10);
  var bulan = parseInt(parts[1], 10) - 1; // 0 = Januari, 6 = Juli, dst.
  var hari = parseInt(parts[2], 10);
  
  // LOGIKA CUT-OFF: Jika tanggal 28 atau lebih, geser ke bulan depan
  if (hari >= 28) {
    bulan = bulan + 1;
    // Jika melewati Desember (bulan > 11), kembali ke Januari tahun berikutnya
    if (bulan > 11) {
      bulan = 0;
      tahun = tahun + 1;
    }
  }
  
  var namaBulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return namaBulan[bulan] + " " + tahun;
}

/**
 * FUNGSI PENYIMPAN ARUS KAS (BUKU BESAR)
 * Menggunakan hitungPeriodeSmart untuk mengunci periode pembukuan.
 */
function simpanTrxArusKas(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Trx Arus Kas");
    if (!sheet) throw new Error("Sheet 'Trx Arus Kas' tidak ditemukan!");
    
    // Panggil otak otomatis untuk menentukan bulan pembukuan
    var periodePembukuan = hitungPeriodeSmart(data.tanggal);
    
    var debit = 0;
    var kredit = 0;
    if (data.tipe === "PEMASUKAN") {
      debit = parseFloat(data.nominal) || 0;
    } else {
      kredit = parseFloat(data.nominal) || 0;
    }
    
    sheet.appendRow([
      Utilities.getUuid(),
      data.tanggal,
      periodePembukuan, // Pasti terisi bulan yang tepat (Misal: Agustus 2026)
      data.tipe,
      data.item,
      data.catatan || "-",
      debit,
      kredit
    ]);
    
    return "Transaksi berhasil disimpan ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan kas: " + error.message);
  }
}

/**
 * FUNGSI PENYIMPAN TABUNGAN ASET
 * Terintegrasi otomatis dengan Arus Kas dan Periode Pembukuan yang selaras.
 */
function simpanTrxAset(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAset = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheetAset) throw new Error("Sheet 'Trx Tabungan Aset' tidak ditemukan!");
    
    // Panggil otak otomatis yang sama agar sinkron dengan Arus Kas
    var periodePembukuan = hitungPeriodeSmart(data.tanggal);

    var idTransaksi = Utilities.getUuid();
    sheetAset.appendRow([
      idTransaksi, 
      data.tanggal, 
      periodePembukuan,
      data.kategori,
      data.item, 
      data.tipe, 
      parseFloat(data.kuantitas) || 0,
      parseFloat(data.harga) || 0, 
      parseFloat(data.total) || 0,
      data.idReferensi || "-"
    ]);
    
    // Jembatan otomatis ke Buku Besar Kas
    var sheetKas = ss.getSheetByName("Trx Arus Kas");
    if (data.tipe === "BELI" && data.coaKas) {
      sheetKas.appendRow([
        Utilities.getUuid(),    
        data.tanggal,           
        periodePembukuan,           
        "PENGELUARAN",          
        data.coaKas,            
        "Pembelian: " + data.item, 
        0,                      
        parseFloat(data.total) || 0  
      ]);
    } else if (data.tipe === "JUAL") {
      sheetKas.appendRow([
        Utilities.getUuid(),    
        data.tanggal,           
        periodePembukuan,           
        "PEMASUKAN",          
        "Pencairan Investasi",  
        "Penjualan: " + data.item, 
        parseFloat(data.total) || 0, 
        0                       
      ]);
    }

    return "Aset berhasil dicatat ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan aset: " + error.message);
  }
}