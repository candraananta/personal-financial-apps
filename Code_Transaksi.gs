/**
 * ==================================================================
 * FILE: Code_Transaksi.gs
 * FUNGSI: Mesin Backend untuk Menyimpan Transaksi & Menghitung Periode
 * PERBAIKAN v2:
 *   - FIX KRITIS getRiwayatBeliAktif(): mapping kolom disesuaikan
 *     dengan urutan appendRow() di simpanTrxAset():
 *     Col A(0)=ID, B(1)=Tanggal, C(2)=Platform, D(3)=Kategori,
 *     E(4)=NamaItem, F(5)=Tipe, G(6)=Qty, H(7)=Harga, I(8)=Total,
 *     J(9)=IdReferensi
 *   - FIX: Sorting riwayat beli berdasarkan tanggal (kolom 1, bukan 0)
 *   - FIX: Harga beli per lot disimpan di beliMap untuk kalkulasi PNL
 * ==================================================================
 */

/**
 * FUNGSI PEMBANTU (SMART CUT-OFF ANTI-TIMEZONE BUG)
 */
function hitungPeriodeSmart(tanggalInput) {
  if (!tanggalInput) return "-";
  var parts = String(tanggalInput).split('-');
  if (parts.length < 3) return tanggalInput;

  var tahun = parseInt(parts[0], 10);
  var bulan = parseInt(parts[1], 10) - 1;
  var hari  = parseInt(parts[2], 10);

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
    tipe,
    item,
    catatan || "-",
    debit  || 0,
    kredit || 0
  ]);
}

// ============================================================
// simpanTrxKlinik()
// ============================================================
function simpanTrxKlinik(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetKlinik = ss.getSheetByName("Trx Fee Klinik");
    if (!sheetKlinik) throw new Error("Sheet 'Trx Fee Klinik' tidak ditemukan!");

    var periodePembukuan = hitungPeriodeSmart(data.tanggal);
    var omset         = parseFloat(data.omset)        || 0;
    var persenKomisi  = parseFloat(data.persenKomisi) || 0;
    var nominalKomisi = parseFloat(data.nominalKomisi)|| 0;
    var labaBersih    = parseFloat(data.labaBersih)   || 0;

    sheetKlinik.appendRow([
      Utilities.getUuid(),
      data.tanggal,
      periodePembukuan,
      data.klinik,
      data.pasien,
      data.tindakan || "-",
      omset,
      persenKomisi,
      nominalKomisi,
      labaBersih
    ]);

    catatKeJurnalKas(
      ss, data.tanggal, periodePembukuan, "PEMASUKAN",
      "Fee Klinik: " + data.klinik,
      "Pasien: " + data.pasien + " | " + (data.tindakan || "-"),
      nominalKomisi, 0
    );

    return "Fee klinik berhasil disimpan ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan klinik: " + error.message);
  }
}

// ============================================================
// [BUG FIX v2] getRiwayatBeliAktif()
//
// Urutan kolom di sheet "Trx Tabungan Aset" (sesuai simpanTrxAset):
//   A(idx 0) = ID_Trx_Aset
//   B(idx 1) = Tanggal_Perolehan
//   C(idx 2) = Platform
//   D(idx 3) = Kategori
//   E(idx 4) = Nama_Item
//   F(idx 5) = Tipe_Transaksi  (BELI / JUAL)
//   G(idx 6) = Kuantitas
//   H(idx 7) = Harga_Satuan
//   I(idx 8) = Total_Nilai
//   J(idx 9) = ID_Referensi
// ============================================================
function getRiwayatBeliAktif() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    var beliMap = {};

    data.forEach(function(row) {
      var id       = String(row[0]).trim();   // A: ID unik transaksi
      var tgl      = row[1];                  // B: Tanggal
      var platform = String(row[2]).trim();   // C: Platform
      var kategori = String(row[3]).trim();   // D: Kategori
      var item     = String(row[4]).trim();   // E: Nama Item
      var tipe     = String(row[5]).trim();   // F: Tipe (BELI/JUAL)
      var qty      = parseFloat(row[6]) || 0; // G: Kuantitas
      var harga    = parseFloat(row[7]) || 0; // H: Harga Satuan
      var total    = parseFloat(row[8]) || 0; // I: Total Nilai
      var idRef    = String(row[9]).trim();   // J: ID Referensi

      if (!id || id === "") return;

      if (tipe === "BELI") {
        beliMap[id] = {
          id:        id,
          tanggal:   tgl,
          platform:  platform,
          kategori:  kategori,
          item:      item,
          qtyBeli:   qty,
          hargaBeli: harga,
          totalBeli: total,
          sisaQty:   qty
        };
      } else if (tipe === "JUAL" && idRef && idRef !== "-" && beliMap[idRef]) {
        beliMap[idRef].sisaQty -= qty;
      }
    });

    // Filter yang masih punya sisa qty > 0
    var hasil = [];
    Object.keys(beliMap).forEach(function(id) {
      var entry = beliMap[id];
      if (entry.sisaQty > 0.000001) { // Pakai epsilon untuk float precision
        hasil.push(entry);
      }
    });

    // Urutkan: terbaru di atas
    hasil.sort(function(a, b) {
      var da = a.tanggal instanceof Date ? a.tanggal.getTime() : new Date(String(a.tanggal)).getTime();
      var db = b.tanggal instanceof Date ? b.tanggal.getTime() : new Date(String(b.tanggal)).getTime();
      return db - da;
    });

    // Format tanggal untuk tampilan
    hasil = hasil.map(function(e) {
      if (e.tanggal instanceof Date) {
        var d = e.tanggal;
        var dd = String(d.getDate()).padStart(2,'0');
        var mm = String(d.getMonth()+1).padStart(2,'0');
        var yy = d.getFullYear();
        e.tanggal = dd + '/' + mm + '/' + yy;
      }
      return e;
    });

    return hasil;
  } catch (error) {
    throw new Error("Gagal memuat riwayat beli: " + error.message);
  }
}

// ============================================================
// simpanTrxArusKas()
// ============================================================
function simpanTrxArusKas(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var periodePembukuan = hitungPeriodeSmart(data.tanggal);
    var debit  = data.tipe === "PEMASUKAN"   ? (parseFloat(data.nominal) || 0) : 0;
    var kredit = data.tipe === "PENGELUARAN" ? (parseFloat(data.nominal) || 0) : 0;

    catatKeJurnalKas(
      ss, data.tanggal, periodePembukuan, data.tipe,
      data.item, data.catatan || "-", debit, kredit
    );

    return "Transaksi berhasil disimpan ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan kas: " + error.message);
  }
}

// ============================================================
// simpanTrxAset()
// Kolom tersimpan: ID, Tanggal, Platform, Kategori, NamaItem,
//                 Tipe, Qty, Harga, Total, IdReferensi
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

    // Urutan kolom HARUS konsisten dengan getRiwayatBeliAktif()
    sheetAset.appendRow([
      idTransaksi,       // A(0): ID
      data.tanggal,      // B(1): Tanggal
      data.platform,     // C(2): Platform
      data.kategori,     // D(3): Kategori
      data.item,         // E(4): Nama Item
      data.tipe,         // F(5): Tipe
      qty,               // G(6): Qty
      harga,             // H(7): Harga Satuan
      total,             // I(8): Total
      data.idReferensi || "-"  // J(9): ID Referensi
    ]);

    // Catat ke GL / Jurnal Arus Kas
    if (data.tipe === "BELI" && data.coaKas) {
      catatKeJurnalKas(
        ss, data.tanggal, periodePembukuan, "PENGELUARAN",
        data.coaKas,
        "Pembelian Aset: " + data.item + " | " + data.platform + " | Qty: " + qty,
        0, total
      );
    } else if (data.tipe === "JUAL") {
      // Hitung harga pokok dari lot referensi untuk GL PNL
      var keteranganJual = "Penjualan Aset: " + data.item + " | " + data.platform + " | Qty: " + qty;
      if (data.idReferensi && data.idReferensi !== "-") {
        keteranganJual += " | Ref: " + data.idReferensi.substring(0,8) + "...";
      }
      catatKeJurnalKas(
        ss, data.tanggal, periodePembukuan, "PEMASUKAN",
        "Pencairan Investasi: " + data.item,
        keteranganJual,
        total, 0
      );
    }

    return "Aset berhasil dicatat ke periode: " + periodePembukuan;
  } catch (error) {
    throw new Error("Gagal menyimpan aset: " + error.message);
  }
}