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
function ambilData(token) {
  wajibLogin_(token);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aturan = bacaAturan_(ss);
  const koreksi = bacaKoreksi_(ss);

  // Keterangan bebas per transaksi email (mis. top-up buat "bayar kos"), by id
  const catatanMap = {};
  bacaTab_(ss, TAB_CATATAN).forEach(r => {
    const k = String(r.id == null ? '' : r.id);
    if (k) catatanMap[k] = String(r.catatan || '');
  });

  const transaksi = bacaTab_(ss, TAB_TRANSAKSI).map(r => {
    const merchant = String(r.merchant || '');
    return {
      id: String(r.id),
      tanggal: r.tanggal instanceof Date ? r.tanggal.toISOString() : null,
      jumlah: Number(r.jumlah) || 0,
      biaya: Number(r.biaya) || 0,
      akun: r.akun,
      merchant: merchant,
      detail: String(r.detail_penerima || ''),
      tipe: r.tipe,
      kategori: tentukanKategori_(r.tipe, merchant, aturan, koreksi),
      catatan: catatanMap[String(r.id)] || ''
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

  // Transaksi cash/manual (di luar QRIS/transfer): bisa tambah & hapus dari dashboard
  const manual = bacaTab_(ss, TAB_MANUAL).map((r, i) => ({
    baris: i + 2,                                   // baris di Sheets (header = baris 1)
    tanggal: r.tanggal instanceof Date ? r.tanggal.toISOString() : null,
    jumlah: Number(r.jumlah) || 0,
    keterangan: String(r.keterangan || ''),
    kategori: String(r.kategori || '').trim() || 'Lainnya'
  })).filter(t => t.tanggal);

  return { transaksi: transaksi, pemasukan: pemasukan, tabungan: tabungan, manual: manual };
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
function simpanKoreksi(token, merchant, kategori) {
  wajibLogin_(token);
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

function ambilDaftarKategori(token) { wajibLogin_(token); return KATEGORI_TETAP; }

const TIPE_TETAP = ['belanja', 'transfer', 'topup', 'transfer_internal'];

/**
 * Ubah tipe satu transaksi email (per transaksi, by id).
 * Dipakai untuk e-wallet: tandai sebagai pengeluaran (belanja/transfer)
 * atau simpan aja (topup), dan belanja atau bukan.
 * Langsung ubah kolom `tipe` di tab transaksi; sync tidak menimpa baris lama.
 */
function simpanTipe(token, id, tipe) {
  wajibLogin_(token);
  const nid = String(id == null ? '' : id).trim();
  if (!nid) throw new Error('ID transaksi kosong.');
  if (TIPE_TETAP.indexOf(tipe) === -1) throw new Error('Tipe tidak valid.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_TRANSAKSI);
  if (!sheet) throw new Error('Tab transaksi belum ada. Jalankan setup() dulu.');
  if (sheet.getLastRow() < 2) throw new Error('Tab transaksi kosong.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const n = sheet.getLastRow() - 1;
    const ids = sheet.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === nid) {
        sheet.getRange(i + 2, 10).setValue(tipe);   // kolom `tipe`
        return { ok: true };
      }
    }
  } finally {
    lock.releaseLock();
  }
  throw new Error('Transaksi tidak ditemukan. Muat ulang dashboard lalu coba lagi.');
}

/**
 * Simpan keterangan bebas untuk satu transaksi email (per transaksi, by id).
 * Mis. top-up e-wallet buat apa ("bayar kos", "dana jajan", ...).
 * Kosongkan untuk menghapus keterangan. Maks 200 karakter.
 */
function simpanCatatan(token, id, catatan) {
  wajibLogin_(token);
  const nid = String(id == null ? '' : id).trim();
  if (!nid) throw new Error('ID transaksi kosong.');
  const teks = String(catatan == null ? '' : catatan).trim().slice(0, 200);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TAB_CATATAN);
  if (!sheet) {
    sheet = ss.insertSheet(TAB_CATATAN);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['id', 'catatan', 'diubah_pada']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@');
    sheet.getRange('C:C').setNumberFormat('yyyy-mm-dd hh:mm');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const n = sheet.getLastRow() - 1;
    let ketemu = -1;
    if (n > 0) {
      const ids = sheet.getRange(2, 1, n, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === nid) { ketemu = i + 2; break; }
      }
    }
    if (teks) {
      if (ketemu > 0) {
        sheet.getRange(ketemu, 2, 1, 2).setValues([[teks, new Date()]]);
      } else {
        sheet.getRange(sheet.getLastRow() + 1, 1, 1, 3).setValues([[nid, teks, new Date()]]);
      }
    } else if (ketemu > 0) {
      sheet.deleteRow(ketemu);   // keterangan dikosongkan = hapus barisnya
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ---------------- Tema (disimpan di server, per akun) ---------------- */

function ambilTema() {
  try { return PropertiesService.getUserProperties().getProperty('TEMA') || 'auto'; }
  catch (e) { return 'auto'; }
}
function simpanTema(tema) {
  if (['auto', 'light', 'dark'].indexOf(tema) === -1) throw new Error('Tema tidak valid.');
  try { PropertiesService.getUserProperties().setProperty('TEMA', tema); } catch (e) {}
  return tema;
}

/* ---------------- Input & hapus pemasukan ---------------- */

const SUMBER_PEMASUKAN = ['Gaji', 'Freelance', 'Lainnya'];
const MAKS_BARIS_PEMASUKAN = 20;

function ambilSumberPemasukan(token) { wajibLogin_(token); return SUMBER_PEMASUKAN; }

/**
 * form = { tanggal: 'yyyy-mm-dd', baris: [{ jumlah, sumber, catatan }, ...] }
 * Semua baris divalidasi dulu; satu saja salah -> tidak ada yang ditulis.
 */
function tambahPemasukan(token, form) {
  wajibLogin_(token);
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
function hapusPemasukan(token, targets) {
  wajibLogin_(token);
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

/* ---------------- Input & hapus transaksi cash/manual ----------------
 * Tab `manual`: tanggal | jumlah | keterangan | kategori.
 * Ditampilkan menyatu dengan transaksi email sebagai tipe 'belanja'
 * dengan akun 'Cash', tapi bisa dihapus dari dashboard.
 */

const MAKS_BARIS_MANUAL = 20;

/**
 * form = { tanggal: 'yyyy-mm-dd', baris: [{ jumlah, keterangan, kategori }, ...] }
 * Semua baris divalidasi dulu; satu saja salah -> tidak ada yang ditulis.
 */
function tambahManual(token, form) {
  wajibLogin_(token);
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String((form && form.tanggal) || ''));
  if (!m) throw new Error('Tanggal tidak valid.');
  const tanggal = new Date(+m[1], +m[2] - 1, +m[3], 12, 0, 0);
  if (isNaN(tanggal.getTime())) throw new Error('Tanggal tidak valid.');

  const daftar = (form && form.baris) || [];
  if (!Array.isArray(daftar) || daftar.length === 0) throw new Error('Isi minimal satu transaksi.');
  if (daftar.length > MAKS_BARIS_MANUAL) {
    throw new Error('Maksimal ' + MAKS_BARIS_MANUAL + ' baris per sekali simpan.');
  }

  const hasil = daftar.map((b, i) => {
    const no = i + 1;
    const jumlah = Number(b && b.jumlah);
    if (!isFinite(jumlah) || jumlah <= 0) throw new Error('Baris ' + no + ': jumlah harus angka lebih dari 0.');
    const keterangan = String((b && b.keterangan) || '').trim().slice(0, 200);
    if (!keterangan) throw new Error('Baris ' + no + ': keterangan wajib diisi (mis. "Parkir cash").');
    let kategori = String((b && b.kategori) || '').trim();
    if (KATEGORI_TETAP.indexOf(kategori) === -1) kategori = 'Lainnya';
    return [tanggal, jumlah, keterangan, kategori];
  });

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_MANUAL);
  if (!sheet) throw new Error('Tab manual belum ada. Jalankan setup() dulu.');

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
 * Hapus transaksi cash tertentu. targets = [{ baris, tanggal (ISO), jumlah, keterangan }, ...]
 * Tiap target diverifikasi: isi baris di Sheets harus COCOK dengan yang dikirim.
 * Kalau ada yang tidak cocok (Sheets berubah di sela-sela), SEMUA dibatalkan.
 */
function hapusManual(token, targets) {
  wajibLogin_(token);
  if (!Array.isArray(targets) || targets.length === 0) throw new Error('Tidak ada yang dihapus.');
  if (targets.length > 100) throw new Error('Terlalu banyak baris sekaligus.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_MANUAL);
  if (!sheet) throw new Error('Tab manual belum ada.');

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
        String(cur[2]) === String(t.keterangan);
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

/* ---------------- Login, sesi & reset password ----------------
 * Akun disimpan sebagai hash SHA-256 + salt di ScriptProperties
 * (tidak di kode, tidak di Sheets). Sesi = token acak + kedaluwarsa.
 */

const PROP_AKUN = 'LOGIN_AKUN';   // JSON {user, garam, hash, email}
const PROP_SESI = 'LOGIN_SESI';   // JSON {token: expMillis}
const PROP_RESET = 'LOGIN_RESET'; // JSON {kode, expMillis}
const SESI_HARI = 7;              // sesi berlaku 7 hari

/** SHA-256 hex dari garam + password. */
function hashKata_(garam, kata) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, garam + '|' + kata, Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    let b = bytes[i];
    if (b < 0) b += 256;
    if (b < 16) hex += '0';
    hex += b.toString(16);
  }
  return hex;
}

function bacaAkun_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_AKUN) || 'null');
  } catch (e) { return null; }
}

function bacaSesi_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_SESI) || '{}') || {};
  } catch (e) { return {}; }
}

/**
 * Jalankan dari editor SEKALI untuk membuat akun login.
 * Username/password/email diminta lewat prompt, tidak ditulis di kode.
 */
function buatAkun() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_AKUN)) {
    throw new Error('Akun sudah ada. Minta reset password dari halaman login untuk menggantinya.');
  }
  const ui = (() => {
    try {
      return SpreadsheetApp.getUi();
    } catch (e) {
      throw new Error('buatAkun() harus di-run dari project Apps Script yang menempel ' +
        'ke spreadsheet (buka spreadsheet → Ekstensi → Apps Script), bukan project standalone.');
    }
  })();
  const user = ui.prompt('Buat akun dashboard', 'Username:', ui.ButtonSet.OK_CANCEL)
    .getResponseText().trim();
  if (!user) throw new Error('Username kosong.');
  const pass = ui.prompt('Buat akun dashboard', 'Password (min 6 karakter):', ui.ButtonSet.OK_CANCEL)
    .getResponseText();
  if (pass.length < 6) throw new Error('Password minimal 6 karakter.');
  const email = ui.prompt('Buat akun dashboard', 'Email pemulihan (untuk reset password):', ui.ButtonSet.OK_CANCEL)
    .getResponseText().trim();
  if (!email) throw new Error('Email kosong.');
  const garam = Utilities.getUuid();
  props.setProperty(PROP_AKUN, JSON.stringify({
    user: user, garam: garam, hash: hashKata_(garam, pass), email: email
  }));
  Logger.log('Akun dibuat untuk username: ' + user);
}

/** Login dari halaman web. Return {token} kalau cocok. */
function login(username, password) {
  const akun = bacaAkun_();
  if (!akun) throw new Error('Akun belum dibuat. Jalankan buatAkun() di Apps Script dulu.');
  const cocok = String(username).trim() === akun.user &&
                hashKata_(akun.garam, String(password)) === akun.hash;
  if (!cocok) throw new Error('Username atau password salah.');
  const token = Utilities.getUuid() + Utilities.getUuid();
  const sesi = bacaSesi_();
  sesi[token] = Date.now() + SESI_HARI * 24 * 3600 * 1000;
  Object.keys(sesi).forEach(function (k) { if (sesi[k] < Date.now()) delete sesi[k]; });
  PropertiesService.getScriptProperties().setProperty(PROP_SESI, JSON.stringify(sesi));
  return { token: token };
}

/** Keluar: hapus token sesi. */
function keluar(token) {
  const sesi = bacaSesi_();
  delete sesi[String(token) || ''];
  PropertiesService.getScriptProperties().setProperty(PROP_SESI, JSON.stringify(sesi));
  return { ok: true };
}

/** Guard di awal tiap fungsi data: token harus valid & belum kedaluwarsa. */
function wajibLogin_(token) {
  const exp = bacaSesi_()[String(token) || ''];
  if (!exp || exp < Date.now()) throw new Error('Sesi habis. Silakan login ulang.');
}

/** Kirim kode reset 6 digit (berlaku 15 menit) ke email pemulihan. */
function mintaReset() {
  const akun = bacaAkun_();
  if (!akun) throw new Error('Akun belum dibuat.');
  const kode = String(Math.floor(100000 + Math.random() * 900000));
  PropertiesService.getScriptProperties().setProperty(PROP_RESET, JSON.stringify({
    kode: kode, exp: Date.now() + 15 * 60 * 1000
  }));
  GmailApp.sendEmail(akun.email,
    'Kode reset password dashboard',
    'Kode reset password kamu: ' + kode + '\nBerlaku 15 menit. ' +
    'Abaikan email ini kalau kamu tidak memintanya.');
  return { ok: true };
}

/** Tukar kode reset dengan password baru. Semua sesi lama ikut dihapus. */
function resetPassword(kode, passwordBaru) {
  const akun = bacaAkun_();
  if (!akun) throw new Error('Akun belum dibuat.');
  let r = {};
  try {
    r = JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_RESET) || '{}') || {};
  } catch (e) {}
  if (!r.kode || r.kode !== String(kode).trim() || r.exp < Date.now()) {
    throw new Error('Kode salah atau kedaluwarsa.');
  }
  if (String(passwordBaru).length < 6) throw new Error('Password minimal 6 karakter.');
  akun.hash = hashKata_(akun.garam, String(passwordBaru));
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_AKUN, JSON.stringify(akun));
  props.deleteProperty(PROP_RESET);
  props.deleteProperty(PROP_SESI);
  return { ok: true };
}

/** Waktu sync terakhir per sumber, untuk label "terakhir diperbarui". */
function ambilStatusSync(token) {
  wajibLogin_(token);
  const props = PropertiesService.getScriptProperties();
  return {
    livin: props.getProperty('LIVIN_SYNC_TERAKHIR'),
    wondr: props.getProperty('WONDR_SYNC_TERAKHIR'),
    bibit: props.getProperty('BIBIT_SYNC_TERAKHIR'),
    gopay: props.getProperty('GOPAY_SYNC_TERAKHIR')
  };
}