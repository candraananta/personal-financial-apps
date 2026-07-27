/**
 * ==================================================================
 * FILE: Code_Transaksi.gs
 * FUNGSI: Mesin Backend untuk Menyimpan Transaksi & Menghitung Periode
 * PERBAIKAN:
 *   - Tambah simpanTrxKlinik()       [BUG #3 FIX]
 *   - Tambah getRiwayatBeliAktif()   [BUG #2 FIX]
 *   - Perbaiki simpanTrxAset() agar konsisten dengan GL
 * ==================================================================
 */

/**
 * FUNGSI PEMBANTU (SMART CUT-OFF ANTI-TIMEZONE BUG):
 * Memecah string tanggal secara langsung tanpa new Date() agar kebal error zona waktu.
 * Aturan: Jika tanggal >= 28, otomatis masuk ke bulan berikutnya.
 */
function hitungPeriodeSmart(tanggalInput) {
  if (!tanggalInput) return "-";

  var parts = String(tanggalInput).split('-');
  if (parts.length < 3) return tanggalInput;

  var tahun = parseInt(parts[0], 10);
  var bulan = parseInt(parts[1], 10) - 1; // 0-indexed: 0=Jan, 11=Des
  var hari  = parseInt(parts[2], 10);

  // LOGIKA CUT-OFF: Tgl 28 ke atas masuk pembukuan bulan depan
  if (hari >= 28) {
    bulan = bulan + 1;
    if (bulan > 11) { bulan = 0; tahun = tahun + 1; }
  }

  var namaBulan = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];
  return namaBulan[bulan] + " " + tahun;
}

// ============================================================
// HELPER: Tulis ke Jurnal Arus Kas (General Ledger)
// ============================================================
function catatKeJurnalKas(ss, tanggal, periode, tipe, item, catatan, debit, kredit) {
  var sheet = ss.getSheetByName("Trx Arus Kas");
  if (!sheet) throw new Error("Sheet 'Trx Arus Kas' tidak ditemukan! Jalankan setupDatabase() terlebih dahulu.");
  sheet.appendRow([
    Utilities.getUuid(),
    tanggal,
    periode,
    tipe,       // "PEMASUKAN" atau "PENGELUARAN"
    item,       // Nama akun / COA
    catatan || "-",
    debit  || 0,
    kredit || 0
  ]);
}

// ============================================================
// [BUG #3 FIX] simpanTrxKlinik() — FUNGSI INI SEBELUMNYA TIDAK ADA!
// Dipanggil oleh submitTrxKlinik() di View_Transaksi.html
// ============================================================
function simpanTrxKlinik(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetKlinik = ss.getSheetByName("Trx Fee Klinik");
    if (!sheetKlinik) throw new Error("Sheet 'Trx Fee Klinik' tidak ditemukan!");

    // Hitung periode pembukuan otomatis
    var periodePembukuan = hitungPeriodeSmart(data.tanggal);

    var omset         = parseFloat(data.omset)        || 0;
    var persenKomisi  = parseFloat(data.persenKomisi) || 0;
    var nominalKomisi = parseFloat(data.nominalKomisi)|| 0;
    var labaBersih    = parseFloat(data.labaBersih)   || 0;

    // --- 1. Catat ke sheet Trx Fee Klinik ---
    sheetKlinik.appendRow([
      Utilities.getUuid(),      // A: ID
      data.tanggal,             // B: Tanggal Input
      periodePembukuan,         // C: Periode Komisi
      data.klinik,              // D: Nama Klinik
      data.pasien,              // E: Nama Pasien
      data.tindakan || "-",     // F: Tindakan
      omset,                    // G: Omset
      persenKomisi,             // H: Komisi %
      nominalKomisi,            // I: Nominal Komisi (PEMASUKAN dokter)
      labaBersih                // J: Laba Bersih Klinik
    ]);

    // --- 2. Catat ke Jurnal Arus Kas (GL) ---
    // Entry PEMASUKAN: Komisi dokter yang diterima
    catatKeJurnalKas(
      ss,
      data.tanggal,
      periodePembukuan,
      "PEMASUKAN",
      "Fee Klinik: " + data.klinik,
      "Pasien: " + data.pasien + " | " + (data.tindakan || "-"),
      nominalKomisi,  // debit = uang masuk
      0
    );

    return "Fee klinik berhasil disimpan ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan klinik: " + error.message);
  }
}

// ============================================================
// [BUG #2 FIX] getRiwayatBeliAktif() — FUNGSI INI SEBELUMNYA TIDAK ADA!
// Mengembalikan daftar aset BELI yang masih ada sisa kuantitas,
// untuk ditampilkan di modal "Pilih Aset Untuk Dijual".
// ============================================================
function getRiwayatBeliAktif() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheet || sheet.getLastRow() <= 1) return [];

    // Kolom: 0=ID, 1=Tanggal, 2=Periode, 3=Platform, 4=Kategori,
    //        5=NamaItem, 6=Tipe, 7=Qty, 8=Harga, 9=Total, 9=IdRef
    // Note: sheet punya 10 kolom (A-J), index 0-9
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    // Hitung saldo qty per ID transaksi BELI
    // - Setiap BELI punya ID sendiri
    // - Setiap JUAL punya kolom idReferensi yang menunjuk ke ID BELI asalnya
    var beliMap = {};

    data.forEach(function(row) {
      var id    = row[0];
      var tgl   = row[1];
      var platform = row[2]; // Kolom C = Platform (sesuai simpanTrxAset)
      var kategori = row[3]; // Kolom D = Kategori
      var item  = row[4];    // Kolom E = NamaItem
      var tipe  = row[5];    // Kolom F = Tipe (BELI/JUAL)
      var qty   = parseFloat(row[6]) || 0;
      var harga = parseFloat(row[7]) || 0;
      var idRef = row[9];    // Kolom J = ID Referensi (untuk transaksi JUAL)

      if (tipe === "BELI") {
        beliMap[id] = {
          id: id,
          tanggal: tgl,
          platform: platform,
          kategori: kategori,
          item: item,
          qtyBeli: qty,
          hargaBeli: harga,
          sisaQty: qty
        };
      } else if (tipe === "JUAL" && idRef && beliMap[idRef]) {
        // Kurangi sisa dari transaksi BELI yang dirujuk
        beliMap[idRef].sisaQty -= qty;
      }
    });

    // Filter hanya yang masih ada sisa, ubah ke array
    var hasil = [];
    Object.keys(beliMap).forEach(function(id) {
      var entry = beliMap[id];
      if (entry.sisaQty > 0) {
        hasil.push(entry);
      }
    });

    // Urutkan dari terbaru ke terlama
    hasil.sort(function(a, b) {
      return String(b.tanggal).localeCompare(String(a.tanggal));
    });

    return hasil;
  } catch (error) {
    throw new Error("Gagal memuat riwayat beli: " + error.message);
  }
}

// ============================================================
// simpanTrxArusKas() — Perbaikan dari duplikat di Code_ArusKas.gs
// ============================================================
function simpanTrxArusKas(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var periodePembukuan = hitungPeriodeSmart(data.tanggal);

    var debit  = data.tipe === "PEMASUKAN"   ? (parseFloat(data.nominal) || 0) : 0;
    var kredit = data.tipe === "PENGELUARAN" ? (parseFloat(data.nominal) || 0) : 0;

    catatKeJurnalKas(
      ss,
      data.tanggal,
      periodePembukuan,
      data.tipe,
      data.item,
      data.catatan || "-",
      debit,
      kredit
    );

    return "Transaksi berhasil disimpan ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan kas: " + error.message);
  }
}

// ============================================================
// simpanTrxAset() — Terintegrasi dengan GL (Jurnal Arus Kas)
// ============================================================
function simpanTrxAset(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAset = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheetAset) throw new Error("Sheet 'Trx Tabungan Aset' tidak ditemukan!");

    var periodePembukuan = hitungPeriodeSmart(data.tanggal);
    var idTransaksi = Utilities.getUuid();

    var qty   = parseFloat(data.kuantitas) || 0;
    var harga = parseFloat(data.harga)     || 0;
    var total = parseFloat(data.total)     || 0;

    // --- 1. Catat ke Trx Tabungan Aset ---
    // Kolom: ID, Tanggal, Platform, Kategori, NamaItem, Tipe, Qty, Harga, Total, IdReferensi
    sheetAset.appendRow([
      idTransaksi,
      data.tanggal,
      data.platform,
      data.kategori,
      data.item,
      data.tipe,
      qty,
      harga,
      total,
      data.idReferensi || "-"
    ]);

    // --- 2. Jembatan otomatis ke Jurnal Kas (GL) ---
    if (data.tipe === "BELI" && data.coaKas) {
      // Pengeluaran: uang keluar untuk beli aset
      catatKeJurnalKas(
        ss,
        data.tanggal,
        periodePembukuan,
        "PENGELUARAN",
        data.coaKas,
        "Pembelian Aset: " + data.item + " | Qty: " + qty,
        0,
        total
      );
    } else if (data.tipe === "JUAL") {
      // Pemasukan: uang masuk dari penjualan aset
      catatKeJurnalKas(
        ss,
        data.tanggal,
        periodePembukuan,
        "PEMASUKAN",
        "Pencairan Investasi",
        "Penjualan Aset: " + data.item + " | Qty: " + qty,
        total,
        0
      );
    }

    return "Aset berhasil dicatat ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan aset: " + error.message);
  }
}