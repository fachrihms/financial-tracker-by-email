/** Titik masuk web app. */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dashboard Pengeluaran')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const KATEGORI_TETAP = ['Belanja', 'Belanja Makan', 'Belanja Online', 'Transport', 'Tagihan', 'Lainnya'];

/**
 * Seluruh data untuk dashboard.
 * Kategori dihitung di sini (koreksi > aturan config > Lainnya), tidak disimpan di Sheets.
 */
function ambilData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aturan = bacaAturan_(ss);
  const koreksi = bacaKoreksi_(ss);

  const transaksi = bacaTab_(ss, TAB_TRANSAKSI).map(r => {
    const merchant = String(r.merchant || '');
    return {
      id: String(r.id),
      tanggal: r.tanggal instanceof Date ? r.tanggal.toISOString() : null,
      jumlah: Number(r.jumlah) || 0,
      biaya: Number(r.biaya) || 0,
      akun: r.akun,
      merchant: merchant,
      tipe: r.tipe,
      kategori: tentukanKategori_(r.tipe, merchant, aturan, koreksi)
    };
  }).filter(t => t.tanggal);

  // Pemasukan disertai nomor baris di Sheets + sidik, supaya bisa dihapus dengan aman
  const pemasukan = bacaTab_(ss, TAB_PEMASUKAN).map((r, i) => ({
    baris: i + 2,                                   // baris di Sheets (header = baris 1)
    tanggal: r.tanggal instanceof Date ? r.tanggal.toISOString() : null,
    jumlah: Number(r.jumlah) || 0,
    sumber: String(r.sumber || ''),
    catatan: String(r.catatan || '')
  })).filter(p => p.tanggal);

  // Tabungan (Bibit dkk.): tidak masuk pengeluaran riil, ditampilkan terpisah
  const tabungan = bacaTab_(ss, TAB_TABUNGAN).map(r => ({
    tanggal: r.tanggal instanceof Date ? r.tanggal.toISOString() : null,
    jumlah: Number(r.jumlah) || 0,
    sumber: String(r.sumber || ''),
    produk: String(r.produk || '')
  })).filter(t => t.tanggal);

  return { transaksi: transaksi, pemasukan: pemasukan, tabungan: tabungan };
}

/* ---------------- Kategori ---------------- */

/** Baca tab config -> [{kategori, kata: [..]}], kata kunci huruf kecil. */
function bacaAturan_(ss) {
  return bacaTab_(ss, TAB_CONFIG).map(r => ({
    kategori: String(r.kategori || '').trim(),
    kata: String(r.kata_kunci || '').toLowerCase().split(',')
            .map(s => s.replace(/^\s+/, '')).filter(s => s.trim() !== '')
  })).filter(a => a.kategori && a.kata.length);
}

/** Baca tab koreksi -> {merchantHurufKecil: kategori}. */
function bacaKoreksi_(ss) {
  const peta = {};
  bacaTab_(ss, TAB_KOREKSI).forEach(r => {
    const m = String(r.merchant || '').trim().toLowerCase();
    const k = String(r.kategori || '').trim();
    if (m && k) peta[m] = k;
  });
  return peta;
}

/**
 * koreksi manual > aturan config > default tipe.
 * Tipe 'belanja' (QRIS, gesek kartu, dll.) yang tidak cocok keyword apa pun -> Belanja.
 * Top-up selalu 'Top-up'. Sisanya (transfer dll.) yang tidak cocok -> Lainnya.
 */
function tentukanKategori_(tipe, merchant, aturan, koreksi) {
  if (tipe === 'topup') return 'Top-up';
  const nama = String(merchant || '').toLowerCase();
  if (koreksi[nama.trim()]) return koreksi[nama.trim()];
  for (let i = 0; i < aturan.length; i++) {
    for (let j = 0; j < aturan[i].kata.length; j++) {
      if (nama.indexOf(aturan[i].kata[j]) !== -1) return aturan[i].kategori;
    }
  }
  return tipe === 'belanja' ? 'Belanja' : 'Lainnya';
}

/**
 * Simpan koreksi kategori untuk satu merchant (berlaku ke semua transaksi merchant itu).
 * Kategori harus salah satu dari daftar tetap.
 */
function simpanKoreksi(merchant, kategori) {
  const nama = String(merchant || '').trim();
  if (!nama) throw new Error('Merchant kosong.');
  if (KATEGORI_TETAP.indexOf(kategori) === -1) throw new Error('Kategori tidak valid.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_KOREKSI);
  if (!sheet) throw new Error('Tab koreksi belum ada. Jalankan setup() dulu.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const n = sheet.getLastRow() - 1;
    if (n > 0) {
      const data = sheet.getRange(2, 1, n, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === nama.toLowerCase()) {
          sheet.getRange(i + 2, 2).setValue(kategori);   // perbarui koreksi lama
          return { ok: true };
        }
      }
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 2).setValues([[nama, kategori]]);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function ambilDaftarKategori() { return KATEGORI_TETAP; }

/* ---------------- Tema (disimpan di server, per akun) ---------------- */

function ambilTema() {
  return PropertiesService.getUserProperties().getProperty('TEMA') || 'auto';
}
function simpanTema(tema) {
  if (['auto', 'light', 'dark'].indexOf(tema) === -1) throw new Error('Tema tidak valid.');
  PropertiesService.getUserProperties().setProperty('TEMA', tema);
  return tema;
}

/* ---------------- Input & hapus pemasukan ---------------- */

const SUMBER_PEMASUKAN = ['Gaji', 'Freelance', 'Lainnya'];
const MAKS_BARIS_PEMASUKAN = 20;

function ambilSumberPemasukan() { return SUMBER_PEMASUKAN; }

/**
 * form = { tanggal: 'yyyy-mm-dd', baris: [{ jumlah, sumber, catatan }, ...] }
 * Semua baris divalidasi dulu; satu saja salah -> tidak ada yang ditulis.
 */
function tambahPemasukan(form) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String((form && form.tanggal) || ''));
  if (!m) throw new Error('Tanggal tidak valid.');
  const tanggal = new Date(+m[1], +m[2] - 1, +m[3], 12, 0, 0);
  if (isNaN(tanggal.getTime())) throw new Error('Tanggal tidak valid.');

  const daftar = (form && form.baris) || [];
  if (!Array.isArray(daftar) || daftar.length === 0) throw new Error('Isi minimal satu pemasukan.');
  if (daftar.length > MAKS_BARIS_PEMASUKAN) {
    throw new Error('Maksimal ' + MAKS_BARIS_PEMASUKAN + ' baris per sekali simpan.');
  }

  const hasil = daftar.map((b, i) => {
    const no = i + 1;
    const jumlah = Number(b && b.jumlah);
    if (!isFinite(jumlah) || jumlah <= 0) throw new Error('Baris ' + no + ': jumlah harus angka lebih dari 0.');
    const sumber = String((b && b.sumber) || '').trim();
    if (SUMBER_PEMASUKAN.indexOf(sumber) === -1) throw new Error('Baris ' + no + ': sumber tidak valid.');
    const catatan = String((b && b.catatan) || '').trim().slice(0, 200);
    return [tanggal, jumlah, sumber, catatan];
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_PEMASUKAN);
  if (!sheet) throw new Error('Tab pemasukan belum ada. Jalankan setup() dulu.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.getRange(sheet.getLastRow() + 1, 1, hasil.length, 4).setValues(hasil);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, jumlahBaris: hasil.length };
}

/**
 * Hapus pemasukan tertentu. targets = [{ baris, tanggal (ISO), jumlah, sumber }, ...]
 * Tiap target diverifikasi: isi baris di Sheets harus COCOK dengan yang dikirim.
 * Kalau ada yang tidak cocok (Sheets berubah di sela-sela), SEMUA dibatalkan.
 */
function hapusPemasukan(targets) {
  if (!Array.isArray(targets) || targets.length === 0) throw new Error('Tidak ada yang dihapus.');
  if (targets.length > 100) throw new Error('Terlalu banyak baris sekaligus.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_PEMASUKAN);
  if (!sheet) throw new Error('Tab pemasukan belum ada.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const total = sheet.getLastRow();
    const nomor = [];

    targets.forEach(t => {
      const b = Number(t.baris);
      if (!isFinite(b) || b < 2 || b > total) {
        throw new Error('Data berubah di Sheets. Muat ulang dashboard lalu coba lagi.');
      }
      const cur = sheet.getRange(b, 1, 1, 3).getValues()[0];
      const cocok =
        cur[0] instanceof Date &&
        Math.abs(cur[0].getTime() - new Date(t.tanggal).getTime()) < 1000 &&
        Number(cur[1]) === Number(t.jumlah) &&
        String(cur[2]) === String(t.sumber);
      if (!cocok) throw new Error('Data berubah di Sheets. Muat ulang dashboard lalu coba lagi.');
      nomor.push(b);
    });

    // Hapus dari bawah ke atas supaya nomor baris di atasnya tidak bergeser
    Array.from(new Set(nomor)).sort((a, b) => b - a).forEach(b => sheet.deleteRow(b));
  } finally {
    lock.releaseLock();
  }
  return { ok: true, dihapus: targets.length };
}

/** Baca tab jadi array objek berdasarkan header baris 1. */
function bacaTab_(ss, nama) {
  const sheet = ss.getSheetByName(nama);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  return data.slice(1).map(baris => {
    const o = {};
    header.forEach((h, i) => { o[h] = baris[i]; });
    return o;
  });
}