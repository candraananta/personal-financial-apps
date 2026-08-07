/**
 * ==================================================================
 * FILE: Code_Piutang.gs
 * FUNGSI: Backend modul Piutang / Accounts Receivable (AR)
 *
 * ARSITEKTUR (Opsi A — Kas Riil):
 *   TIMBUL   → kas keluar → PENGELUARAN di GL
 *   PELUNASAN → kas masuk → PEMASUKAN di GL
 *
 * SHEET:
 *   Master Debitur  : [ID, Nama, Kontak, Catatan]
 *   Trx Piutang     : [ID_Trx, Tanggal, ID_Induk, Nama_Debitur,
 *                       Tipe, Nominal, Sisa_Piutang,
 *                       Jatuh_Tempo, Keterangan, Status]
 *
 * KOLOM INDEX (0-based):
 *   0  ID_Trx
 *   1  Tanggal
 *   2  ID_Induk          ← ID transaksi TIMBUL (untuk grup cicilan)
 *   3  Nama_Debitur
 *   4  Tipe              ← "TIMBUL" | "PELUNASAN"
 *   5  Nominal           ← nominal transaksi ini
 *   6  Sisa_Piutang      ← sisa setelah transaksi ini
 *   7  Jatuh_Tempo       ← date | "" jika tidak diisi
 *   8  Keterangan
 *   9  Status            ← "AKTIF" | "LUNAS"
 * ==================================================================
 */

// ============================================================
// MASTER DEBITUR
// ============================================================

/**
 * Ambil semua debitur dari Master Debitur
 * Return: [[ID, Nama, Kontak, Catatan], ...]
 */
function getReferensiDebitur() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName('Master Debitur');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
}

/**
 * Simpan debitur baru ke Master Debitur
 */
function simpanMasterDebitur(nama, kontak, catatan) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName('Master Debitur');
  if (!sheet)
    throw new Error("Sheet 'Master Debitur' tidak ditemukan! Jalankan setupDatabase().");
  if (!nama || !nama.trim())
    throw new Error('Nama debitur wajib diisi.');

  // Cek duplikat nama (case-insensitive)
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    var lower    = nama.trim().toLowerCase();
    var isDup    = existing.some(function(r) {
      return String(r[0]).trim().toLowerCase() === lower;
    });
    if (isDup) throw new Error("Debitur '" + nama.trim() + "' sudah ada di Master.");
  }

  sheet.appendRow([
    Utilities.getUuid(),
    nama.trim(),
    (kontak || '').trim(),
    (catatan || '').trim()
  ]);
  return "Debitur '" + nama.trim() + "' berhasil ditambahkan!";
}

// ============================================================
// PIUTANG — BACA
// ============================================================

/**
 * Ambil semua piutang dengan status AKTIF (sisa > 0)
 * Dipakai untuk dropdown pelunasan di frontend.
 *
 * Return: array of {
 *   id, tanggal, namaDebitur, nominalAwal,
 *   sisaPiutang, jatuhTempo, keterangan
 * }
 */
function getPiutangAktif() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet()
                  .getSheetByName('Trx Piutang');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

    // Kumpulkan hanya baris TIMBUL yang masih AKTIF
    var hasil = [];
    data.forEach(function(row) {
      var tipe   = String(row[4]).trim();
      var status = String(row[9]).trim();
      if (tipe === 'TIMBUL' && status === 'AKTIF') {
        var tgl = row[1];
        if (tgl instanceof Date) {
          var dd = String(tgl.getDate()).padStart(2, '0');
          var mm = String(tgl.getMonth() + 1).padStart(2, '0');
          var yy = tgl.getFullYear();
          tgl = dd + '/' + mm + '/' + yy;
        } else {
          tgl = String(tgl || '-');
        }

        var jt = row[7];
        var jtStr = '';
        if (jt instanceof Date) {
          jtStr = String(jt.getDate()).padStart(2,'0') + '/' +
                  String(jt.getMonth()+1).padStart(2,'0') + '/' +
                  jt.getFullYear();
        } else if (jt && String(jt).trim() !== '') {
          // Coba format dari string YYYY-MM-DD
          var parts = String(jt).split('-');
          if (parts.length === 3) {
            jtStr = parts[2] + '/' + parts[1] + '/' + parts[0];
          } else {
            jtStr = String(jt);
          }
        }

        hasil.push({
          id:           String(row[0]),
          tanggal:      tgl,
          namaDebitur:  String(row[3]),
          nominalAwal:  parseFloat(row[5]) || 0,
          sisaPiutang:  parseFloat(row[6]) || 0,
          jatuhTempo:   jtStr,
          keterangan:   String(row[8] || '-')
        });
      }
    });

    // Urutkan: sisa terbesar dulu
    hasil.sort(function(a, b) { return b.sisaPiutang - a.sisaPiutang; });
    return hasil;

  } catch (e) {
    throw new Error('Gagal memuat piutang aktif: ' + e.message);
  }
}

// ============================================================
// PIUTANG — SIMPAN
// ============================================================

/**
 * simpanTrxPiutang(payload)
 *
 * payload untuk TIMBUL:
 * {
 *   tipe:        'TIMBUL',
 *   tanggal:     'YYYY-MM-DD',
 *   namaDebitur: string,
 *   nominal:     number,
 *   jatuhTempo:  'YYYY-MM-DD' | '',
 *   keterangan:  string
 * }
 *
 * payload untuk PELUNASAN:
 * {
 *   tipe:          'PELUNASAN',
 *   tanggal:       'YYYY-MM-DD',
 *   idInduk:       string (ID baris TIMBUL),
 *   namaDebitur:   string,
 *   nominalBayar:  number,
 *   sisaSekarang:  number (sisa SEBELUM pelunasan ini),
 *   keterangan:    string
 * }
 */
function simpanTrxPiutang(payload) {
  try {
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var sheetTrx   = ss.getSheetByName('Trx Piutang');
    if (!sheetTrx)
      throw new Error("Sheet 'Trx Piutang' tidak ditemukan! Jalankan setupDatabase().");

    var idTrx = Utilities.getUuid();
    var periode = hitungPeriodeSmart(payload.tanggal);

    // ── TIMBUL ──────────────────────────────────────────────
    if (payload.tipe === 'TIMBUL') {
      var nominal = parseFloat(payload.nominal) || 0;
      if (nominal <= 0) throw new Error('Nominal piutang harus lebih dari 0.');
      if (!payload.namaDebitur || !payload.namaDebitur.trim())
        throw new Error('Nama debitur wajib diisi.');

      sheetTrx.appendRow([
        idTrx,                           // [0] ID_Trx (sekaligus ID_Induk)
        payload.tanggal,                 // [1] Tanggal
        idTrx,                           // [2] ID_Induk = diri sendiri
        payload.namaDebitur.trim(),      // [3] Nama_Debitur
        'TIMBUL',                        // [4] Tipe
        nominal,                         // [5] Nominal
        nominal,                         // [6] Sisa_Piutang = nominal awal
        payload.jatuhTempo || '',        // [7] Jatuh_Tempo
        (payload.keterangan || '').trim() || '-', // [8] Keterangan
        'AKTIF'                          // [9] Status
      ]);

      // Jurnal GL → PENGELUARAN (kas keluar)
      catatKeJurnalKas(
        ss,
        payload.tanggal,
        periode,
        'PENGELUARAN',
        'Piutang: ' + payload.namaDebitur.trim(),
        (payload.keterangan || '').trim() ||
          'Piutang timbul · Jatuh tempo: ' + (payload.jatuhTempo || 'tidak ditentukan'),
        0,       // debit
        nominal  // kredit
      );

      return 'Piutang berhasil dicatat! Periode: ' + periode +
             ' · Sisa: Rp' + nominal.toLocaleString('id-ID');
    }

    // ── PELUNASAN ────────────────────────────────────────────
    else if (payload.tipe === 'PELUNASAN') {
      var nomBayar  = parseFloat(payload.nominalBayar) || 0;
      var sisaLama  = parseFloat(payload.sisaSekarang) || 0;

      if (nomBayar <= 0)
        throw new Error('Nominal pelunasan harus lebih dari 0.');
      if (nomBayar > sisaLama)
        throw new Error(
          'Nominal pelunasan (Rp' + nomBayar.toLocaleString('id-ID') +
          ') melebihi sisa piutang (Rp' + sisaLama.toLocaleString('id-ID') + ').'
        );

      var sisaBaru  = sisaLama - nomBayar;
      var isLunas   = sisaBaru < 0.01; // toleransi floating point
      if (isLunas) sisaBaru = 0;

      // Tulis baris PELUNASAN
      sheetTrx.appendRow([
        idTrx,                                     // [0] ID_Trx
        payload.tanggal,                           // [1] Tanggal
        payload.idInduk,                           // [2] ID_Induk
        payload.namaDebitur.trim(),                // [3] Nama_Debitur
        'PELUNASAN',                               // [4] Tipe
        nomBayar,                                  // [5] Nominal bayar
        sisaBaru,                                  // [6] Sisa setelah ini
        '',                                        // [7] Jatuh_Tempo (kosong)
        (payload.keterangan || '').trim() || '-',  // [8] Keterangan
        isLunas ? 'LUNAS' : 'AKTIF'               // [9] Status baris ini
      ]);

      // Update sisa_piutang & status di baris TIMBUL (induk)
      _updateSisaInduk(sheetTrx, payload.idInduk, sisaBaru, isLunas);

      // Jurnal GL → PEMASUKAN (kas masuk)
      catatKeJurnalKas(
        ss,
        payload.tanggal,
        periode,
        'PEMASUKAN',
        'Pelunasan Piutang: ' + payload.namaDebitur.trim(),
        (payload.keterangan || '').trim() ||
          'Pelunasan piutang · Sisa: Rp' + sisaBaru.toLocaleString('id-ID'),
        nomBayar,  // debit
        0          // kredit
      );

      var pesanStatus = isLunas
        ? '🎉 Piutang LUNAS! Terima kasih ' + payload.namaDebitur.trim()
        : 'Pelunasan berhasil · Sisa piutang: Rp' + sisaBaru.toLocaleString('id-ID');
      return pesanStatus + ' · Periode: ' + periode;
    }

    else {
      throw new Error('Tipe transaksi tidak valid: ' + payload.tipe);
    }

  } catch (e) {
    throw new Error('Gagal menyimpan piutang: ' + e.message);
  }
}

// ============================================================
// HELPER INTERNAL
// ============================================================

/**
 * Update kolom Sisa_Piutang (col 7) dan Status (col 10)
 * pada baris TIMBUL (induk) setelah ada pelunasan.
 */
function _updateSisaInduk(sheet, idInduk, sisaBaru, isLunas) {
  if (!idInduk) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === idInduk &&
        String(data[i][4]).trim() === 'TIMBUL') {
      // Col 7 (1-based) = Sisa_Piutang, Col 10 = Status
      sheet.getRange(i + 1, 7).setValue(sisaBaru);
      sheet.getRange(i + 1, 10).setValue(isLunas ? 'LUNAS' : 'AKTIF');
      return;
    }
  }
  // Tidak ditemukan — tidak throw, biarkan tetap jalan
  Logger.log('WARNING: ID Induk tidak ditemukan: ' + idInduk);
}