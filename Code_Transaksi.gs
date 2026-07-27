/**
 * ==================================================================
 * FILE: Code_Transaksi.gs
 * FUNGSI: Mesin Backend untuk Menyimpan Transaksi & Menghitung Periode
 * PERBAIKAN v3:
 *   - getRiwayatBeliAktif():
 *       * Mode FISIK (Emas/Perak/dll): return per lot seperti sebelumnya
 *       * Mode NOMINAL (Reksadana/Saham): agregasi semua lot per item+platform
 *         sehingga frontend bisa jual sebagian nominal
 *   - simpanTrxAset(): handle jual nominal (qty=1, harga=nominal)
 *   - KATA_KUNCI_UNIT_FISIK sinkron dengan frontend
 * ==================================================================
 */

// === KONSTANTA — SINKRON DENGAN FRONTEND ===
var KATA_KUNCI_UNIT_FISIK = ['emas', 'perak', 'logam', 'gold', 'silver', 'antam', 'fisik', 'koin'];

function isKategoriFisik_(namaKategori) {
  if (!namaKategori) return false;
  var lower = String(namaKategori).toLowerCase();
  return KATA_KUNCI_UNIT_FISIK.some(function(k) { return lower.indexOf(k) !== -1; });
}

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
// getRiwayatBeliAktif()
//
// Mengembalikan campuran dua tipe entry:
//
// 1. MODE FISIK (Emas, Perak, dll):
//    Entry per lot dengan sisaQty dalam unit fisik (gram, keping, dll)
//    { id, tanggal, platform, kategori, item, qtyBeli, hargaBeli, totalBeli, sisaQty }
//
// 2. MODE NOMINAL (Reksadana, Saham, dll):
//    Entry diagregasi per kombinasi item+platform+kategori.
//    sisaQty = 1 (dummy), hargaBeli = total nominal sisa (bukan per unit).
//    Frontend akan menggunakan hargaBeli sebagai "sisa nominal".
//    { id (lot pertama sebagai referensi), tanggal, platform, kategori, item,
//      qtyBeli, hargaBeli (=totalNominalSisa), totalBeli, sisaQty=1 }
//
// KOLOM Trx Tabungan Aset:
//   A(0)=ID, B(1)=Tanggal, C(2)=Platform, D(3)=Kategori,
//   E(4)=NamaItem, F(5)=Tipe, G(6)=Qty, H(7)=Harga, I(8)=Total, J(9)=IdReferensi
// ============================================================
function getRiwayatBeliAktif() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    // ── PASS 1: Proses semua transaksi ke beliMap ──
    var beliMap    = {};  // id → entry fisik
    var nominalMap = {};  // key(item||platform||kategori) → entry teragregasi

    data.forEach(function(row) {
      var id       = String(row[0]).trim();
      var tgl      = row[1];
      var platform = String(row[2]).trim();
      var kategori = String(row[3]).trim();
      var item     = String(row[4]).trim();
      var tipe     = String(row[5]).trim();
      var qty      = parseFloat(row[6]) || 0;
      var harga    = parseFloat(row[7]) || 0;
      var total    = parseFloat(row[8]) || 0;
      var idRef    = String(row[9]).trim();

      if (!id || id === "") return;

      var fisik = isKategoriFisik_(kategori);

      if (tipe === "BELI") {
        if (fisik) {
          // Simpan per lot
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
        } else {
          // Agregasi nominal
          var key = item + "||" + platform + "||" + kategori;
          if (!nominalMap[key]) {
            nominalMap[key] = {
              id:            id,          // simpan ID lot pertama sebagai referensi
              tanggal:       tgl,
              platform:      platform,
              kategori:      kategori,
              item:          item,
              totalModalBeli: 0,
              totalTerjual:   0,
              lotIds:         []
            };
          }
          nominalMap[key].totalModalBeli += total;
          nominalMap[key].lotIds.push(id);
        }
      }

      else if (tipe === "JUAL") {
        if (fisik) {
          // Kurangi sisa qty lot referensi
          if (idRef && idRef !== "-" && beliMap[idRef]) {
            beliMap[idRef].sisaQty -= qty;
          }
        } else {
          // Kurangi sisa nominal
          var keyJ = item + "||" + platform + "||" + kategori;
          if (nominalMap[keyJ]) {
            nominalMap[keyJ].totalTerjual += total;
          }
        }
      }
    });

    var hasil = [];

    // ── PASS 2: Kumpulkan entry FISIK dengan sisaQty > 0 ──
    Object.keys(beliMap).forEach(function(id) {
      var entry = beliMap[id];
      if (entry.sisaQty > 0.000001) {
        hasil.push(entry);
      }
    });

    // ── PASS 3: Kumpulkan entry NOMINAL dengan sisa > 0 ──
    Object.keys(nominalMap).forEach(function(key) {
      var n   = nominalMap[key];
      var sisa = n.totalModalBeli - n.totalTerjual;
      if (sisa > 0.01) {
        hasil.push({
          id:        n.id,        // lot pertama sebagai referensi
          tanggal:   n.tanggal,
          platform:  n.platform,
          kategori:  n.kategori,
          item:      n.item,
          qtyBeli:   1,
          hargaBeli: sisa,        // frontend baca ini sebagai "sisa nominal"
          totalBeli: n.totalModalBeli,
          sisaQty:   1            // dummy — mode nominal tidak pakai qty
        });
      }
    });

    // ── Urutkan: terbaru di atas ──
    hasil.sort(function(a, b) {
      var da = a.tanggal instanceof Date ? a.tanggal.getTime() : new Date(String(a.tanggal)).getTime();
      var db = b.tanggal instanceof Date ? b.tanggal.getTime() : new Date(String(b.tanggal)).getTime();
      return db - da;
    });

    // ── Format tanggal untuk tampilan ──
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
//
// Kolom tersimpan: ID, Tanggal, Platform, Kategori, NamaItem,
//                 Tipe, Qty, Harga, Total, IdReferensi
//
// Mode FISIK: Qty = unit fisik sesungguhnya, Harga = per unit
// Mode NOMINAL: Qty = 1, Harga = total nominal, Total = total nominal
// ============================================================
function simpanTrxAset(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetAset = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheetAset) throw new Error("Sheet 'Trx Tabungan Aset' tidak ditemukan!");

    var periodePembukuan = hitungPeriodeSmart(data.tanggal);
    var idTransaksi = Utilities.getUuid();

    var qty   = parseFloat(data.kuantitas) || 1;
    var harga = parseFloat(data.harga)     || 0;
    var total = parseFloat(data.total)     || 0;

    // Untuk mode nominal: pastikan qty=1 dan harga=total
    var fisik = isKategoriFisik_(data.kategori);
    if (!fisik) {
      qty   = 1;
      harga = total;
    }

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

    // Catat ke GL
    if (data.tipe === "BELI" && data.coaKas) {
      catatKeJurnalKas(
        ss, data.tanggal, periodePembukuan, "PENGELUARAN",
        data.coaKas,
        "Pembelian Aset: " + data.item + " | " + data.platform +
          (fisik ? " | Qty: " + qty : " | Nominal: Rp" + total.toLocaleString()),
        0, total
      );
    } else if (data.tipe === "JUAL") {
      var keteranganJual = "Penjualan Aset: " + data.item + " | " + data.platform;
      if (fisik) {
        keteranganJual += " | Qty: " + qty;
      } else {
        keteranganJual += " | Nominal: Rp" + total.toLocaleString();
      }
      if (data.idReferensi && data.idReferensi !== "-") {
        keteranganJual += " | Ref: " + String(data.idReferensi).substring(0,8) + "...";
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