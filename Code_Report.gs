/**
 * ==================================================================
 * FILE: Code_Report.gs
 * FUNGSI: Menarik data transaksi untuk direkapitulasi oleh Frontend
 * PERBAIKAN v2:
 *   - getDataLaporan(): tambah field hargaSatuan & idReferensi
 *     agar frontend bisa hitung PNL per lot secara akurat
 *   - getRekapPNLAset(): fungsi baru — hitung realized PNL di backend
 *     berdasarkan harga beli lot yang dirujuk, bukan rata-rata
 * ==================================================================
 */

function getDataLaporan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataReport = { klinik: [], aset: [], kas: [] };

  // 1. Fee Klinik (10 kolom: A-J)
  var sheetKlinik = ss.getSheetByName("Trx Fee Klinik");
  if (sheetKlinik && sheetKlinik.getLastRow() > 1) {
    dataReport.klinik = sheetKlinik
      .getRange(2, 1, sheetKlinik.getLastRow() - 1, 10)
      .getDisplayValues();
  }

  // 2. Tabungan Aset (10 kolom: A-J)
  //    A=ID, B=Tanggal, C=Platform, D=Kategori, E=NamaItem,
  //    F=Tipe, G=Qty, H=Harga, I=Total, J=IdReferensi
  var sheetAset = ss.getSheetByName("Trx Tabungan Aset");
  if (sheetAset && sheetAset.getLastRow() > 1) {
    dataReport.aset = sheetAset
      .getRange(2, 1, sheetAset.getLastRow() - 1, 10)
      .getDisplayValues();
  }

  // 3. Arus Kas / GL (8 kolom: A-H)
  var sheetKas = ss.getSheetByName("Trx Arus Kas");
  if (sheetKas && sheetKas.getLastRow() > 1) {
    dataReport.kas = sheetKas
      .getRange(2, 1, sheetKas.getLastRow() - 1, 8)
      .getDisplayValues();
  }

  return dataReport;
}

/**
 * getRekapPNLAset()
 * Menghitung Realized PNL per transaksi JUAL secara akurat
 * menggunakan harga beli dari lot yang dirujuk (FIFO by reference).
 *
 * Return: array of { item, platform, kategori, tglJual, qtyJual,
 *                    hargaJual, hargaBeli, totalJual, hargaPokokJual,
 *                    pnl, periode }
 */
function getRekapPNLAset() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    // Buat map harga beli per ID lot
    var beliMap = {};
    data.forEach(function(row) {
      var id    = String(row[0]).trim();
      var tipe  = String(row[5]).trim();
      var harga = parseFloat(row[7]) || 0;
      var total = parseFloat(row[8]) || 0;
      var qty   = parseFloat(row[6]) || 0;
      if (tipe === "BELI" && id) {
        beliMap[id] = { harga: harga, total: total, qty: qty,
                        item: String(row[4]), platform: String(row[2]),
                        kategori: String(row[3]), tanggal: row[1] };
      }
    });

    var hasil = [];
    data.forEach(function(row) {
      var tipe  = String(row[5]).trim();
      if (tipe !== "JUAL") return;

      var idRef    = String(row[9]).trim();
      var qtyJual  = parseFloat(row[6]) || 0;
      var hargaJual= parseFloat(row[7]) || 0;
      var totalJual= parseFloat(row[8]) || 0;
      var item     = String(row[4]);
      var platform = String(row[2]);
      var kategori = String(row[3]);
      var tglJual  = row[1];

      var hargaBeli = 0, hargaPokokJual = 0;
      if (idRef && idRef !== "-" && beliMap[idRef]) {
        hargaBeli     = beliMap[idRef].harga;
        hargaPokokJual = hargaBeli * qtyJual;
      }

      var pnl = totalJual - hargaPokokJual;

      hasil.push({
        item:           item,
        platform:       platform,
        kategori:       kategori,
        tglJual:        tglJual,
        qtyJual:        qtyJual,
        hargaJual:      hargaJual,
        hargaBeli:      hargaBeli,
        totalJual:      totalJual,
        hargaPokokJual: hargaPokokJual,
        pnl:            pnl,
        idRef:          idRef
      });
    });

    return hasil;
  } catch(e) {
    throw new Error("Gagal hitung PNL: " + e.message);
  }
}