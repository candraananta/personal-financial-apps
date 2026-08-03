/**
 * ==================================================================
 * FILE: Code_Report.gs
 * FUNGSI: Menarik data transaksi untuk direkapitulasi oleh Frontend
 *
 * PERBAIKAN v3:
 *   - getDataLaporan():
 *       * Sheet Klinik  → getDisplayValues() TETAP (sudah string semua, OK)
 *       * Sheet Aset    → getValues() + normalisasi manual di sini
 *         (getDisplayValues() menghasilkan "1.234.567" yang merusak parseFloat di frontend)
 *       * Sheet Kas     → getValues() + normalisasi manual
 *       * Tanggal diformat ke "YYYY-MM-DD" agar frontend bisa parse konsisten
 *       * Angka numerik dikembalikan sebagai number, bukan string
 *   - renderCardAset() dipindah ke View_Report.html (lihat catatan di bawah)
 *   - getRekapPNLAset(): tidak berubah
 * ==================================================================
 */

/**
 * HELPER: Format tanggal Date → "YYYY-MM-DD"
 * Menggunakan kalkulasi lokal agar tidak kena bug timezone.
 */
function formatTanggalYMD_(tgl) {
  if (!tgl) return "";
  if (tgl instanceof Date) {
    var y  = tgl.getFullYear();
    var m  = String(tgl.getMonth() + 1).padStart(2, "0");
    var d  = String(tgl.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + d;
  }
  return String(tgl);
}

/**
 * HELPER: Normalisasi periode (Date atau string) → "NamaBulan TAHUN"
 */
function normPeriodeGS_(val) {
  if (!val || val === "") return "";
  var namaBulan = [
    "Januari","Februari","Maret","April","Mei","Juni",
    "Juli","Agustus","September","Oktober","November","Desember"
  ];
  if (val instanceof Date) {
    return namaBulan[val.getMonth()] + " " + val.getFullYear();
  }
  var str = String(val).trim();
  // Jika sudah "NamaBulan TAHUN" — kembalikan apa adanya
  if (/^[A-Za-z]+\s+\d{4}$/.test(str)) return str;
  // Coba parse string tanggal YYYY-MM-DD
  var parts = str.split("-");
  if (parts.length === 3) {
    var bln = parseInt(parts[1], 10) - 1;
    return namaBulan[bln] + " " + parts[0];
  }
  return str;
}

/**
 * getDataLaporan()
 * Dipanggil oleh initModulReport() di View_Report.html.
 *
 * Return: { klinik: [[...]], aset: [[...]], kas: [[...]] }
 *
 * Format per sheet:
 *   klinik : getDisplayValues() — sudah aman (semua data relevan = string/angka sederhana)
 *   aset   : getValues() → diproses → array dengan tipe data tepat
 *   kas    : getValues() → diproses → array dengan tipe data tepat
 *
 * Kolom Trx Tabungan Aset (10 kolom):
 *   [0]=ID  [1]=Tanggal  [2]=Platform  [3]=Kategori  [4]=NamaItem
 *   [5]=Tipe  [6]=Qty  [7]=Harga  [8]=Total  [9]=IdReferensi
 *
 * Kolom Trx Arus Kas (8 kolom):
 *   [0]=ID  [1]=Tanggal  [2]=Periode  [3]=Tipe  [4]=Akun
 *   [5]=Catatan  [6]=Debit  [7]=Kredit
 *
 * Kolom Trx Fee Klinik (10 kolom):
 *   [0]=ID  [1]=Tanggal  [2]=Periode  [3]=Klinik  [4]=Pasien
 *   [5]=Tindakan  [6]=Omset  [7]=KomisiPersen  [8]=NominalKomisi  [9]=LabaBersih
 */
function getDataLaporan() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var dataReport = { klinik: [], aset: [], kas: [] };

  /* ── 1. Fee Klinik ─────────────────────────────────────────── */
  var sheetKlinik = ss.getSheetByName("Trx Fee Klinik");
  if (sheetKlinik && sheetKlinik.getLastRow() > 1) {
    // getValues() agar tanggal & angka bertipe asli
    var rawKlinik = sheetKlinik
      .getRange(2, 1, sheetKlinik.getLastRow() - 1, 10)
      .getValues();

    dataReport.klinik = rawKlinik.map(function(row) {
      return [
        String(row[0]),                          // [0] ID
        formatTanggalYMD_(row[1]),               // [1] Tanggal → "YYYY-MM-DD"
        normPeriodeGS_(row[2]),                  // [2] Periode → "NamaBulan TAHUN"
        String(row[3] || ""),                    // [3] Klinik
        String(row[4] || ""),                    // [4] Pasien
        String(row[5] || "-"),                   // [5] Tindakan
        parseFloat(row[6])  || 0,                // [6] Omset
        parseFloat(row[7])  || 0,                // [7] KomisiPersen
        parseFloat(row[8])  || 0,                // [8] NominalKomisi
        parseFloat(row[9])  || 0                 // [9] LabaBersih
      ];
    });
  }

  /* ── 2. Tabungan Aset ──────────────────────────────────────── */
  var sheetAset = ss.getSheetByName("Trx Tabungan Aset");
  if (sheetAset && sheetAset.getLastRow() > 1) {
    var rawAset = sheetAset
      .getRange(2, 1, sheetAset.getLastRow() - 1, 10)
      .getValues();

    dataReport.aset = rawAset.map(function(row) {
      return [
        String(row[0]),                          // [0] ID
        formatTanggalYMD_(row[1]),               // [1] Tanggal → "YYYY-MM-DD"
        String(row[2] || ""),                    // [2] Platform
        String(row[3] || ""),                    // [3] Kategori
        String(row[4] || ""),                    // [4] NamaItem
        String(row[5] || ""),                    // [5] Tipe (BELI/JUAL)
        parseFloat(row[6])  || 0,                // [6] Qty
        parseFloat(row[7])  || 0,                // [7] Harga
        parseFloat(row[8])  || 0,                // [8] Total
        String(row[9] || "-")                    // [9] IdReferensi
      ];
    });
  }

  /* ── 3. Arus Kas / GL ──────────────────────────────────────── */
  var sheetKas = ss.getSheetByName("Trx Arus Kas");
  if (sheetKas && sheetKas.getLastRow() > 1) {
    var rawKas = sheetKas
      .getRange(2, 1, sheetKas.getLastRow() - 1, 8)
      .getValues();

    dataReport.kas = rawKas.map(function(row) {
      return [
        String(row[0]),                          // [0] ID Jurnal
        formatTanggalYMD_(row[1]),               // [1] Tanggal → "YYYY-MM-DD"
        normPeriodeGS_(row[2]),                  // [2] Periode → "NamaBulan TAHUN"
        String(row[3] || ""),                    // [3] Tipe (PEMASUKAN/PENGELUARAN)
        String(row[4] || ""),                    // [4] Akun/Item
        String(row[5] || "-"),                   // [5] Catatan
        parseFloat(row[6])  || 0,                // [6] Debit (Masuk)
        parseFloat(row[7])  || 0                 // [7] Kredit (Keluar)
      ];
    });
  }

  return dataReport;
}


/**
 * getRekapPNLAset()
 * Menghitung Realized PNL per transaksi JUAL secara akurat
 * menggunakan harga beli dari lot yang dirujuk (by reference).
 *
 * Return: array of {
 *   item, platform, kategori, tglJual, qtyJual,
 *   hargaJual, hargaBeli, totalJual, hargaPokokJual, pnl, idRef
 * }
 */
function getRekapPNLAset() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Trx Tabungan Aset");
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    // Buat map harga beli per ID lot
    var beliMap = {};
    data.forEach(function(row) {
      var id   = String(row[0]).trim();
      var tipe = String(row[5]).trim();
      if (tipe === "BELI" && id) {
        beliMap[id] = {
          harga:    parseFloat(row[7]) || 0,
          total:    parseFloat(row[8]) || 0,
          qty:      parseFloat(row[6]) || 0,
          item:     String(row[4]),
          platform: String(row[2]),
          kategori: String(row[3]),
          tanggal:  formatTanggalYMD_(row[1])
        };
      }
    });

    var hasil = [];
    data.forEach(function(row) {
      if (String(row[5]).trim() !== "JUAL") return;

      var idRef        = String(row[9]).trim();
      var qtyJual      = parseFloat(row[6]) || 0;
      var hargaJual    = parseFloat(row[7]) || 0;
      var totalJual    = parseFloat(row[8]) || 0;
      var hargaBeli    = 0;
      var hargaPokok   = 0;

      if (idRef && idRef !== "-" && beliMap[idRef]) {
        hargaBeli  = beliMap[idRef].harga;
        hargaPokok = hargaBeli * qtyJual;
      }

      hasil.push({
        item:           String(row[4]),
        platform:       String(row[2]),
        kategori:       String(row[3]),
        tglJual:        formatTanggalYMD_(row[1]),
        qtyJual:        qtyJual,
        hargaJual:      hargaJual,
        hargaBeli:      hargaBeli,
        totalJual:      totalJual,
        hargaPokokJual: hargaPokok,
        pnl:            totalJual - hargaPokok,
        idRef:          idRef
      });
    });

    return hasil;
  } catch(e) {
    throw new Error("Gagal hitung PNL: " + e.message);
  }
}