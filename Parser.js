/**
 * PARSER EMAIL LIVIN' BY MANDIRI  (v2)
 * Perbaikan: dukung format "IDR", referensi toleran variasi label,
 * merchant fallback. Belum menulis ke Sheets.
 */

const LIVIN_QUERY = 'from:noreply.livin@bankmandiri.co.id newer_than:90d';

// 4 digit terakhir rekening sendiri, untuk deteksi transfer internal
const REKENING_SENDIRI = []; // contoh: ['4823']

const BULAN = {
  Jan: 0, Feb: 1, Mar: 2, Apr: 3, Mei: 4, Jun: 5,
  Jul: 6, Agu: 7, Sep: 8, Okt: 9, Nov: 10, Des: 11
};

const FALLBACK_MERCHANT = {
  belanja: 'Pembayaran (tanpa penerima)',
  topup: 'Top-up',
  transfer: 'Transfer'
};

/** Tes: log semua transaksi yang berhasil di-parse. */
function tesParserLivin() {
  const threads = GmailApp.search(LIVIN_QUERY, 0, 30);
  let ok = 0, gagal = 0;
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const hasil = parseLivin(msg);
      if (hasil) { ok++; Logger.log(JSON.stringify(hasil)); }
      else { gagal++; Logger.log('DITOLAK: ' + msg.getSubject()); }
    });
  });
  Logger.log('Selesai. Berhasil: ' + ok + ' | Ditolak: ' + gagal);
}

/** Statistik kualitas parsing atas 100 thread terakhir. */
function cekKualitasParser() {
  const threads = GmailApp.search(LIVIN_QUERY, 0, 100);
  const stat = { total: 0, tanpaTanggal: 0, tanpaMerchant: 0, tanpaSumber: 0,
                 perTipe: {}, ditolak: {} };
  const idSeen = {};
  let idDuplikat = 0;

  threads.forEach(t => t.getMessages().forEach(msg => {
    const h = parseLivin(msg);
    if (!h) {
      const s = msg.getSubject();
      stat.ditolak[s] = (stat.ditolak[s] || 0) + 1;
      return;
    }
    stat.total++;
    stat.perTipe[h.tipe] = (stat.perTipe[h.tipe] || 0) + 1;
    if (!h.tanggal) stat.tanpaTanggal++;
    if (!h.merchant) stat.tanpaMerchant++;
    if (h.akun.indexOf('????') !== -1) stat.tanpaSumber++;
    if (idSeen[h.id]) idDuplikat++;
    idSeen[h.id] = true;
  }));

  stat.idDuplikat = idDuplikat;
  Logger.log(JSON.stringify(stat, null, 2));
}

/** Parse satu pesan Gmail. Return objek transaksi, atau null kalau bukan transaksi. */
function parseLivin(msg) {
  const subject = msg.getSubject();
  const tipe = tentukanTipe(subject);
  if (!tipe) return null;

  const html = msg.getBody();
  const baris = ambilBaris(html);
  const penerima = ambilPenerima(html);

  const nominalUtama = cariNominal(baris, [
    'Nominal Transaksi', 'Nominal Top-up', 'Jumlah Transfer',
    'Nominal Transfer', 'Nominal'
  ]);
  const total = cariNominal(baris, ['Total Transaksi']);
  const biaya = cariNominal(baris, ['Biaya Transaksi', 'Biaya Transfer']) || 0;
  const jumlah = total != null ? total : nominalUtama;
  if (jumlah == null) return null;

  const tanggal = parseTanggal(baris['Tanggal'], baris['Jam']);
  const sumber = baris['__sumber'] || null;
  const merchant = penerima.nama || FALLBACK_MERCHANT[tipe] || 'Tidak diketahui';

  return {
    id: cariRef(baris) || msg.getId(),
    email_id: msg.getId(),
    tanggal: tanggal,
    jumlah: jumlah,
    biaya: biaya,
    arah: 'out',
    akun: 'Livin ****' + (sumber || '????'),
    merchant: merchant,
    detail_penerima: penerima.detail,
    kategori: null,
    tipe: koreksiTipe(tipe, penerima),
    sumber_subject: subject
  };
}

/** Tentukan jenis transaksi dari subject. */
function tentukanTipe(subject) {
  const s = subject.toLowerCase();
  if (s.indexOf('pembayaran berhasil') !== -1) return 'belanja';
  if (s.indexOf('top-up') !== -1) return 'topup';
  if (s.indexOf('transfer') !== -1 && s.indexOf('berhasil') !== -1) return 'transfer';
  return null;
}

/** Transfer ke rekening sendiri jadi transfer_internal. */
function koreksiTipe(tipe, penerima) {
  if (tipe !== 'transfer') return tipe;
  const empatAkhir = (penerima.detail || '').replace(/\D/g, '').slice(-4);
  return REKENING_SENDIRI.indexOf(empatAkhir) !== -1 ? 'transfer_internal' : 'transfer';
}

/** Nomor referensi transaksi, toleran terhadap variasi label antar-versi email. */
function cariRef(baris) {
  const kandidat = ['No. Referensi', 'Nomor Referensi', 'No. Referensi BI Fast'];
  for (let i = 0; i < kandidat.length; i++) {
    if (baris[kandidat[i]]) return baris[kandidat[i]];
  }
  return null;
}

/** Ambil semua pasangan label-nilai dari tabel email + 4 digit rekening sumber. */
function ambilBaris(html) {
  const hasil = {};
  const rx = /<td[^>]*>([^<]{1,60})<\/td>\s*<td[^>]*>([^<]{1,80})<\/td>/g;
  let m;
  while ((m = rx.exec(html)) !== null) {
    const label = bersihkan(m[1]);
    const nilai = bersihkan(m[2]);
    if (label && nilai && !(label in hasil)) hasil[label] = nilai;
  }
  const rs = /(?:Sumber Dana|Rekening Sumber)[\s\S]{0,300}?\*{4}(\d{4})/.exec(html);
  if (rs) hasil['__sumber'] = rs[1];
  return hasil;
}

/** Kotak "Penerima" atau "Penyedia Jasa" (nama + detail di bawahnya). */
function ambilPenerima(html) {
  const rx = /(?:Penerima|Penyedia Jasa)<\/p>\s*<h4[^>]*>([^<]*)<\/h4>\s*<p[^>]*>([^<]*)<\/p>/;
  const m = rx.exec(html);
  if (!m) return { nama: null, detail: null };
  return {
    nama: bersihkan(m[1]).replace(/\.{2,}$/, '').trim(),
    detail: bersihkan(m[2])
  };
}

/** Cari nominal dari daftar kemungkinan label; return angka atau null. */
function cariNominal(baris, labels) {
  for (let i = 0; i < labels.length; i++) {
    if (baris[labels[i]]) return parseRupiah(baris[labels[i]]);
  }
  return null;
}

/** "Rp 1.234.567,00" atau "IDR 1.234.567,00" -> 1234567 */
function parseRupiah(teks) {
  const m = /(?:Rp|IDR)\s*([\d.]+)(?:,(\d{1,2}))?/i.exec(teks);
  if (!m) return null;
  const bulat = parseInt(m[1].replace(/\./g, ''), 10);
  const desimal = m[2] ? parseInt(m[2], 10) / Math.pow(10, m[2].length) : 0;
  return bulat + desimal;
}

/** "20 Sep 2026" + "08:55:50 WIB" -> Date (WIB = UTC+7) */
function parseTanggal(tgl, jam) {
  if (!tgl) return null;
  const t = /(\d{1,2})\s+([A-Za-z]{3})\s+(\d{4})/.exec(tgl);
  if (!t || !(t[2] in BULAN)) return null;
  const j = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(jam || '') || [0, 0, 0, 0];
  return new Date(Date.UTC(+t[3], BULAN[t[2]], +t[1], +j[1] - 7, +j[2], +(j[3] || 0)));
}

/** Rapikan whitespace dan entitas HTML dasar. */
function bersihkan(s) {
  return String(s)
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
}
function diagnosisTransfer() {
  const threads = GmailApp.search(LIVIN_QUERY, 0, 100);
  threads.forEach(t => t.getMessages().forEach(msg => {
    const s = msg.getSubject();
    if (s.indexOf('Transfer') !== -1 && !parseLivin(msg)) {
      Logger.log('=== ' + s + ' | ' + msg.getDate());
      Logger.log('LABEL: ' + JSON.stringify(ambilBaris(msg.getBody())));
      Logger.log('PENERIMA: ' + JSON.stringify(ambilPenerima(msg.getBody())));
    }
  }));
}

/* ============================================================
 *  PENYIMPANAN KE GOOGLE SHEETS
 * ============================================================ */

const TAB_TRANSAKSI = 'transaksi';
const TAB_PEMASUKAN = 'pemasukan';
const TAB_DITOLAK   = 'log_ditolak';
const TAB_CONFIG  = 'config';
const TAB_KOREKSI = 'koreksi';
const TAB_TABUNGAN  = 'tabungan';

const KOLOM_TRANSAKSI = [
  'id', 'tanggal', 'jumlah', 'biaya', 'arah', 'akun', 'merchant',
  'detail_penerima', 'kategori', 'tipe', 'sumber_subject', 'email_id', 'diproses_pada'
];
const KOLOM_PEMASUKAN = ['tanggal', 'jumlah', 'sumber', 'catatan'];
const KOLOM_DITOLAK   = ['email_id', 'sender', 'subject', 'alasan', 'waktu'];
const KOLOM_TABUNGAN  = ['id', 'tanggal', 'jumlah', 'sumber', 'produk', 'unit', 'email_id', 'diproses_pada'];

/** Jalankan SEKALI: membuat tab + header + format. Aman diulang. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buatTab_(ss, TAB_TRANSAKSI, KOLOM_TRANSAKSI);
  buatTab_(ss, TAB_PEMASUKAN, KOLOM_PEMASUKAN);
  buatTab_(ss, TAB_DITOLAK, KOLOM_DITOLAK);
  buatTab_(ss, TAB_KOREKSI, ['merchant', 'kategori']);
  buatTab_(ss, TAB_TABUNGAN, KOLOM_TABUNGAN);

  // Tab config: hanya diisi contoh kalau masih kosong (tidak menimpa aturan lu)
  const cfg = buatTab_(ss, TAB_CONFIG, ['kategori', 'kata_kunci']);
  if (cfg.getLastRow() < 2) {
    cfg.getRange(2, 1, 5, 2).setValues([
      ['Belanja Makan',  'warung, resto, restoran, rumah makan, kopi, coffee, cafe, kafe, bakso, mie, nasi, ayam, sate, lontong, soto, roti, bakery, geprek, martabak, es, kedai, depot, warteg, padang, boba, tea, cream, aroma'],
      ['Transport',      'grab, gojek, gocar, goride, bluebird, taksi, ojek, parkir, tol, krl, mrt, lrt, transjakarta, kai, bensin, pertamina, spbu, shell, e-money'],
      ['Tagihan',        'pln, listrik, pdam, air, indihome, telkom, telkomsel, xl, axis, indosat, tri, pulsa, paket data, bpjs, asuransi, kos, kost, sewa, cicilan, pajak'],
      ['Belanja',        'indomaret, alfamart, alfamidi, superindo, hypermart, transmart, lotte, giant, ranch, mall, toko, store, mart, minimarket'],
      ['Belanja Online', 'shopee, tokopedia, lazada, blibli, tiktok']
    ]);
  }
  cfg.setColumnWidth(2, 600);
  migrasiKategori_(ss);

  // Kolom id & email_id WAJIB teks: nomor referensi 19 digit rusak kalau jadi angka
  const t = ss.getSheetByName(TAB_TRANSAKSI);
  t.getRange('A:A').setNumberFormat('@');
  t.getRange('L:L').setNumberFormat('@');
  t.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
  t.getRange('M:M').setNumberFormat('yyyy-mm-dd hh:mm');
  t.getRange('C:D').setNumberFormat('#,##0');

  const p = ss.getSheetByName(TAB_PEMASUKAN);
  p.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  p.getRange('B:B').setNumberFormat('#,##0');

  const tb = ss.getSheetByName(TAB_TABUNGAN);
  tb.getRange('A:A').setNumberFormat('@');          // id = nomor order, simpan sebagai teks
  tb.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm');
  tb.getRange('C:C').setNumberFormat('#,##0');
  tb.getRange('G:G').setNumberFormat('@');
  tb.getRange('H:H').setNumberFormat('yyyy-mm-dd hh:mm');

  Logger.log('Setup selesai.');
}

function buatTab_(ss, nama, kolom) {
  let sheet = ss.getSheetByName(nama);
  if (!sheet) sheet = ss.insertSheet(nama);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, kolom.length).setValues([kolom]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Migrasi sekali jalan: Makan -> Belanja Makan, dan keyword belanja online dipisah
 * dari Belanja ke kategori Belanja Online (jika belum ada). Tidak menimpa kata
 * kunci lain yang sudah dikustomisasi.
 */
function migrasiKategori_(ss) {
  const cfg = ss.getSheetByName(TAB_CONFIG);
  if (cfg && cfg.getLastRow() >= 2) {
    const n = cfg.getLastRow() - 1;
    const data = cfg.getRange(2, 1, n, 2).getValues();
    const online = ['shopee', 'tokopedia', 'lazada', 'blibli', 'tiktok'];
    let belanjaOnline = '';
    let berubah = false;

    data.forEach(r => {
      const kat = String(r[0]).trim();
      const kata = String(r[1] || '');
      if (kat === 'Makan') {
        r[0] = 'Belanja Makan';
        berubah = true;
      } else if (kat === 'Belanja') {
        const kataArr = kata.split(',').map(s => s.trim()).filter(Boolean);
        const onlineArr = kataArr.filter(s => online.indexOf(s.toLowerCase()) !== -1);
        const sisa = kataArr.filter(s => online.indexOf(s.toLowerCase()) === -1).join(', ');
        if (sisa !== kata) { r[1] = sisa; berubah = true; }
        if (onlineArr.length) belanjaOnline = onlineArr.join(', ');
      }
    });

    const adaOnline = data.some(r => String(r[0]).trim() === 'Belanja Online');
    if (belanjaOnline && !adaOnline) {
      data.push(['Belanja Online', belanjaOnline]);
      berubah = true;
    }
    if (berubah) cfg.getRange(2, 1, data.length, 2).setValues(data);
  }

  // Koreksi manual lama bertuliskan 'Makan' disesuaikan namanya
  const k = ss.getSheetByName(TAB_KOREKSI);
  if (k && k.getLastRow() >= 2) {
    const m = k.getLastRow() - 1;
    const kd = k.getRange(2, 1, m, 2).getValues();
    let kubah = false;
    kd.forEach(r => {
      if (String(r[1]).trim() === 'Makan') { r[1] = 'Belanja Makan'; kubah = true; }
    });
    if (kubah) k.getRange(2, 1, kd.length, 2).setValues(kd);
  }
}

/**
 * Ambil SEMUA thread yang cocok (GmailApp.search mentok 500/call).
 * Dipakai backfill histori panjang supaya tidak ada yang kelewat.
 */
function cariSemuaThread_(query) {
  const semua = [];
  let start = 0;
  while (true) {
    const batch = GmailApp.search(query, start, 500);
    if (!batch.length) break;
    semua.push.apply(semua, batch);
    if (batch.length < 500) break;
    start += 500;
  }
  return semua;
}

/**
 * Batas waktu untuk query Gmail.
 * - hariKebelakang <= 365 tanpa offset -> newer_than (perilaku lama).
 * - Lebih panjang / pakai offset -> after: (+ before:) yyyy/MM/dd.
 * Contoh chunk per tahun: syncLivin(365, 365) = rentang 1-2 tahun lalu.
 */
function tambahBatasWaktu_(query, hariKebelakang, hariOffset) {
  if (!hariKebelakang) return query;
  const offset = hariOffset || 0;
  if (!offset && hariKebelakang <= 365) return query + ' newer_than:' + hariKebelakang + 'd';
  const sekarang = new Date();
  const mulai = new Date(sekarang.getTime() - (hariKebelakang + offset) * 86400000);
  query += ' after:' + Utilities.formatDate(mulai, 'Asia/Jakarta', 'yyyy/MM/dd');
  if (offset) {
    const selesai = new Date(sekarang.getTime() - offset * 86400000);
    query += ' before:' + Utilities.formatDate(selesai, 'Asia/Jakarta', 'yyyy/MM/dd');
  }
  return query;
}

/**
 * Sync utama: tarik email Livin' baru, parse, dedup, tulis ke Sheets.
 * hariKebelakang: jendela pencarian. Kosong = ikuti sync terakhir (inkremental).
 * hariOffset (opsional): geser jendela ke masa lalu, mis. syncLivin(365, 365).
 */
function syncLivin(hariKebelakang, hariOffset) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetT = ss.getSheetByName(TAB_TRANSAKSI);
  const sheetD = ss.getSheetByName(TAB_DITOLAK);
  if (!sheetT || !sheetD) throw new Error('Jalankan setup() dulu.');

  const props = PropertiesService.getScriptProperties();
  const terakhir = props.getProperty('LIVIN_SYNC_TERAKHIR');

  // Tentukan jendela pencarian
  let query = 'from:noreply.livin@bankmandiri.co.id';
  if (hariKebelakang) {
    query = tambahBatasWaktu_(query, hariKebelakang, hariOffset);
  } else if (terakhir) {
    // mundur 1 hari sebagai margin; dedup menjamin tidak ada duplikat
    const tgl = new Date(new Date(terakhir).getTime() - 24 * 3600 * 1000);
    query += ' after:' + Utilities.formatDate(tgl, 'Asia/Jakarta', 'yyyy/MM/dd');
  } else {
    query += ' newer_than:90d'; // sync pertama
  }

  // Dedup: baca semua id yang sudah ada (sekali)
  const idAda = ambilIdAda_(sheetT);
  const emailDitolakAda = ambilEmailIdDitolak_(sheetD);

  const barisBaru = [];
  const barisDitolak = [];
  const sekarang = new Date();
  let dilewati = 0;

  const threads = cariSemuaThread_(query);
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const h = parseLivin(msg);

      if (!h) {
        const eid = msg.getId();
        if (!emailDitolakAda[eid]) {
          emailDitolakAda[eid] = true;
          barisDitolak.push([
            eid, 'livin', msg.getSubject(),
            alasanTolak_(msg.getSubject()), sekarang
          ]);
        }
        return;
      }

      const id = String(h.id);
      if (idAda[id]) { dilewati++; return; }
      idAda[id] = true; // cegah duplikat di dalam batch yang sama

      barisBaru.push([
        id, h.tanggal, h.jumlah, h.biaya, h.arah, h.akun, h.merchant,
        h.detail_penerima || '', h.kategori || '', h.tipe,
        h.sumber_subject, h.email_id, sekarang
      ]);
    });
  });

  // Tulis massal
  if (barisBaru.length) {
    sheetT.getRange(sheetT.getLastRow() + 1, 1, barisBaru.length, KOLOM_TRANSAKSI.length)
          .setValues(barisBaru);
  }
  if (barisDitolak.length) {
    sheetD.getRange(sheetD.getLastRow() + 1, 1, barisDitolak.length, KOLOM_DITOLAK.length)
          .setValues(barisDitolak);
  }

  props.setProperty('LIVIN_SYNC_TERAKHIR', sekarang.toISOString());
  Logger.log('Sync selesai. Baru: ' + barisBaru.length +
             ' | Sudah ada: ' + dilewati +
             ' | Ditolak baru: ' + barisDitolak.length);
}

/** Backfill panjang: panggil manual, mis. syncLivin(365). */
function backfill90() { syncLivin(90); }

/** Kumpulan id yang sudah tersimpan -> objek {id: true}. */
function ambilIdAda_(sheet) {
  const ada = {};
  const n = sheet.getLastRow() - 1;
  if (n < 1) return ada;
  sheet.getRange(2, 1, n, 1).getValues().forEach(r => {
    if (r[0] !== '') ada[String(r[0])] = true;
  });
  return ada;
}

function ambilEmailIdDitolak_(sheet) {
  const ada = {};
  const n = sheet.getLastRow() - 1;
  if (n < 1) return ada;
  sheet.getRange(2, 1, n, 1).getValues().forEach(r => {
    if (r[0] !== '') ada[String(r[0])] = true;
  });
  return ada;
}

/** Alasan penolakan yang mudah dibaca untuk log. */
function alasanTolak_(subject) {
  const s = subject.toLowerCase();
  if (s.indexOf('tidak berhasil') !== -1) return 'transaksi gagal';
  if (s.indexOf('dihubungkan') !== -1) return 'notifikasi akun, bukan transaksi';
  if (!tentukanTipe(subject)) return 'subject tidak dikenali';
  return 'nominal tidak ditemukan / format baru';
}

/** Pasang trigger sync tiap 6 jam. Jalankan SEKALI; aman diulang. */
function pasangTrigger() {
  // Hapus trigger lama supaya tidak dobel
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncLivin') ScriptApp.deleteTrigger(t);
    if (t.getHandlerFunction() === 'syncBibit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncLivin').timeBased().everyHours(6).create();
  ScriptApp.newTrigger('syncBibit').timeBased().everyHours(6).create();
  ScriptApp.newTrigger('syncWondr').timeBased().everyHours(6).create();
  Logger.log('Trigger terpasang: syncLivin + syncBibit + syncWondr tiap 6 jam.');
}

/* ============================================================
 *  PARSER BIBIT (TABUNGAN / REKSA DANA)
 *  Email: "Unit Reksa Dana Kamu Sudah Masuk" dari no-reply@bibit.id
 * ============================================================ */

const BIBIT_QUERY = 'from:no-reply@bibit.id newer_than:90d';

/** Tes: log semua pembelian Bibit yang berhasil di-parse. */
function tesParserBibit() {
  const threads = GmailApp.search(BIBIT_QUERY, 0, 30);
  let ok = 0, gagal = 0;
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const hasil = parseBibit(msg);
      if (hasil) { ok++; Logger.log(JSON.stringify(hasil)); }
      else { gagal++; Logger.log('DITOLAK: ' + msg.getSubject()); }
    });
  });
  Logger.log('Selesai. Berhasil: ' + ok + ' | Ditolak: ' + gagal);
}

/**
 * Parse satu pesan Bibit. Return objek tabungan, atau null kalau bukan
 * konfirmasi pembelian reksa dana.
 */
function parseBibit(msg) {
  const s = msg.getSubject().toLowerCase();
  if (s.indexOf('reksa dana') === -1) return null;
  if (s.indexOf('masuk') === -1 && s.indexOf('berhasil') === -1) return null;

  const teks = msg.getBody()
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ');

  const mOrder = /Nomor Order\s*([A-Za-z0-9]+)/.exec(teks);
  const mProduk = /Nama Reksa Dana\s*(.+?)\s*Nilai Pembelian/.exec(teks);
  const mNilai = /Nilai Pembelian\s*(Rp[\d.,]+)/i.exec(teks);
  const mUnit = /Jumlah Unit\s*([\d.,]+)/.exec(teks);
  if (!mOrder || !mNilai) return null;

  const jumlah = parseRupiahBibit(mNilai[1]);
  if (jumlah == null || jumlah <= 0) return null;

  return {
    id: mOrder[1],
    email_id: msg.getId(),
    tanggal: msg.getDate(),
    jumlah: jumlah,
    sumber: 'Bibit',
    produk: mProduk ? mProduk[1].trim() : '',
    unit: mUnit ? parseDesimal_(mUnit[1]) : null
  };
}

/**
 * "Rp200,000" / "Rp1,500,000" (koma = pemisah ribuan ala Bibit) -> angka.
 * Beda dengan parseRupiah Livin yang memakai titik sebagai pemisah ribuan.
 */
function parseRupiahBibit(teks) {
  const m = /Rp\s*([\d.,]+)/i.exec(teks);
  if (!m) return null;
  let angka = m[1].replace(/,/g, '');
  if ((angka.match(/\./g) || []).length > 1) angka = angka.replace(/\./g, '');
  const n = parseFloat(angka);
  return isNaN(n) ? null : Math.round(n);
}

/** "133.9976" -> 133.9976 (jumlah unit, boleh desimal). */
function parseDesimal_(teks) {
  const n = parseFloat(String(teks).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

/**
 * Sync utama Bibit: tarik email konfirmasi pembelian, parse, dedup,
 * tulis ke tab tabungan. Email lain dari Bibit dicatat di log_ditolak.
 */
function syncBibit(hariKebelakang, hariOffset) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetT = ss.getSheetByName(TAB_TABUNGAN);
  const sheetD = ss.getSheetByName(TAB_DITOLAK);
  if (!sheetT || !sheetD) throw new Error('Jalankan setup() dulu.');

  const props = PropertiesService.getScriptProperties();
  const terakhir = props.getProperty('BIBIT_SYNC_TERAKHIR');

  let query = 'from:no-reply@bibit.id';
  if (hariKebelakang) {
    query = tambahBatasWaktu_(query, hariKebelakang, hariOffset);
  } else if (terakhir) {
    const tgl = new Date(new Date(terakhir).getTime() - 24 * 3600 * 1000);
    query += ' after:' + Utilities.formatDate(tgl, 'Asia/Jakarta', 'yyyy/MM/dd');
  } else {
    query += ' newer_than:90d'; // sync pertama
  }

  const idAda = ambilIdAda_(sheetT);
  const emailDitolakAda = ambilEmailIdDitolak_(sheetD);

  const barisBaru = [];
  const barisDitolak = [];
  const sekarang = new Date();
  let dilewati = 0;

  const threads = cariSemuaThread_(query);
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const h = parseBibit(msg);

      if (!h) {
        const eid = msg.getId();
        if (!emailDitolakAda[eid]) {
          emailDitolakAda[eid] = true;
          barisDitolak.push([
            eid, 'bibit', msg.getSubject(),
            'bukan konfirmasi pembelian reksa dana', sekarang
          ]);
        }
        return;
      }

      const id = String(h.id);
      if (idAda[id]) { dilewati++; return; }
      idAda[id] = true;

      barisBaru.push([
        id, h.tanggal, h.jumlah, h.sumber, h.produk,
        h.unit == null ? '' : h.unit, h.email_id, sekarang
      ]);
    });
  });

  if (barisBaru.length) {
    sheetT.getRange(sheetT.getLastRow() + 1, 1, barisBaru.length, KOLOM_TABUNGAN.length)
          .setValues(barisBaru);
  }
  if (barisDitolak.length) {
    sheetD.getRange(sheetD.getLastRow() + 1, 1, barisDitolak.length, KOLOM_DITOLAK.length)
          .setValues(barisDitolak);
  }

  props.setProperty('BIBIT_SYNC_TERAKHIR', sekarang.toISOString());
  Logger.log('Sync Bibit selesai. Baru: ' + barisBaru.length +
             ' | Sudah ada: ' + dilewati +
             ' | Ditolak baru: ' + barisDitolak.length);
}

/** Backfill panjang: panggil manual, mis. syncBibit(365). */
function backfillBibit90() { syncBibit(90); }

/* ============================================================
 *  PARSER WONDR BY BNI
 *  Email: "Top-up berhasil!" / "Transaksi berhasil!" / "Transfer berhasil!"
 *  dari wondr@bni.co.id. Hasil ditulis ke tab transaksi yang sama
 *  dengan akun 'Wondr ****706', jadi dashboard langsung jalan.
 * ============================================================ */

const WONDR_QUERY = 'from:wondr@bni.co.id newer_than:90d';

const BULAN_WONDR = {
  Januari: 0, Februari: 1, Maret: 2, April: 3, Mei: 4, Juni: 5,
  Juli: 6, Agustus: 7, September: 8, Oktober: 9, November: 10, Desember: 11,
  Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6,
  Agu: 7, Aug: 7, Sep: 8, Okt: 9, Oct: 9, Nov: 10, Des: 11, Dec: 11
};

/** Tes: log semua transaksi wondr yang berhasil di-parse. */
function tesParserWondr() {
  const threads = GmailApp.search(WONDR_QUERY, 0, 30);
  let ok = 0, gagal = 0;
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const hasil = parseWondr(msg);
      if (hasil) { ok++; Logger.log(JSON.stringify(hasil)); }
      else { gagal++; Logger.log('DITOLAK: ' + msg.getSubject()); }
    });
  });
  Logger.log('Selesai. Berhasil: ' + ok + ' | Ditolak: ' + gagal);
}

/** Parse satu pesan wondr. Return objek transaksi, atau null kalau bukan transaksi. */
function parseWondr(msg) {
  const subject = msg.getSubject();
  const html = msg.getBody();
  const teksKasat = html.replace(/<[^>]+>/g, ' ');

  const tipe = tentukanTipeWondr(subject, teksKasat);
  if (!tipe) return null;

  const baris = ambilBaris(html);
  const penerima = ambilPihakWondr(html, '(?:Penerima|Tujuan)');
  const sumber = ambilPihakWondr(html, 'Sumber dana');

  const nominal = cariNominal(baris, ['Nominal']);
  const total = cariNominal(baris, ['Total']);
  const biaya = cariNominal(baris, ['Biaya transaksi', 'Biaya admin']) || 0;
  const jumlah = total != null ? total : nominal;
  if (jumlah == null) return null;

  const tanggal = parseTanggalWondr(baris['Tanggal'], baris['Waktu']);
  const akhirSumber = ((sumber.detail || '').match(/(\d+)[^\d]*$/) || [])[1] || '????';
  const merchant = (penerima.nama || 'Transaksi wondr').replace(/[\s\-–—]+$/, '');

  return {
    id: cariRefWondr(baris) || msg.getId(),
    email_id: msg.getId(),
    tanggal: tanggal,
    jumlah: jumlah,
    biaya: biaya,
    arah: 'out',
    akun: 'Wondr ****' + akhirSumber,
    merchant: merchant,
    detail_penerima: penerima.detail,
    kategori: null,
    tipe: koreksiTipe(tipe, penerima),
    sumber_subject: subject
  };
}

/**
 * Tipe dari kalimat greeting ("top-up e-Wallet" / "transfer uang" /
 * "pembayaran pakai Virtual Account"), fallback ke subject.
 */
function tentukanTipeWondr(subject, teks) {
  const t = teks.toLowerCase();
  if (t.indexOf('top-up e-wallet') !== -1 || t.indexOf('top up e-wallet') !== -1) return 'topup';
  if (t.indexOf('transfer uang') !== -1) return 'transfer';
  if (t.indexOf('virtual account') !== -1) return 'belanja';
  const s = subject.toLowerCase();
  if (s.indexOf('top-up') !== -1 || s.indexOf('top up') !== -1) return 'topup';
  if (s.indexOf('transfer') !== -1) return 'transfer';
  if (s.indexOf('transaksi') !== -1 && s.indexOf('berhasil') !== -1) return 'belanja';
  return null;
}

/** Blok "Penerima"/"Tujuan"/"Sumber dana": 1 baris judul + 2 baris isi. */
function ambilPihakWondr(html, label) {
  const rx = new RegExp(
    '<b>' + label + '<\\/b>\\s*<\\/td>\\s*<\\/tr>\\s*<tr>\\s*' +
    '<td[^>]*>([^<]*)<\\/td>\\s*<\\/tr>\\s*<tr>\\s*<td[^>]*>([^<]*)<\\/td>'
  );
  const m = rx.exec(html);
  if (!m) return { nama: null, detail: null };
  return { nama: bersihkan(m[1]), detail: bersihkan(m[2]) };
}

/** Nomor referensi: "Ref ID" (top-up) atau "Reference ID" (lainnya). */
function cariRefWondr(baris) {
  const kandidat = ['Ref ID', 'Reference ID'];
  for (let i = 0; i < kandidat.length; i++) {
    if (baris[kandidat[i]]) return baris[kandidat[i]];
  }
  return null;
}

/** "28 Agustus 2026" atau "24-Jul-2026" + "19:17:23 WIB" / "20:49 WIB" -> Date (WIB = UTC+7) */
function parseTanggalWondr(tgl, jam) {
  if (!tgl) return null;
  const t = /(\d{1,2})[-\s]([A-Za-z]+)[-\s](\d{4})/.exec(tgl.trim());
  if (!t || !(t[2] in BULAN_WONDR)) return null;
  const j = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec(jam || '') || [0, 0, 0, 0];
  return new Date(Date.UTC(+t[3], BULAN_WONDR[t[2]], +t[1], +j[1] - 7, +j[2], +(j[3] || 0)));
}

/**
 * Sync utama wondr: tarik email notifikasi, parse, dedup,
 * tulis ke tab transaksi. Email lain dicatat di log_ditolak.
 */
function syncWondr(hariKebelakang, hariOffset) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetT = ss.getSheetByName(TAB_TRANSAKSI);
  const sheetD = ss.getSheetByName(TAB_DITOLAK);
  if (!sheetT || !sheetD) throw new Error('Jalankan setup() dulu.');

  const props = PropertiesService.getScriptProperties();
  const terakhir = props.getProperty('WONDR_SYNC_TERAKHIR');

  let query = 'from:wondr@bni.co.id';
  if (hariKebelakang) {
    query = tambahBatasWaktu_(query, hariKebelakang, hariOffset);
  } else if (terakhir) {
    const tgl = new Date(new Date(terakhir).getTime() - 24 * 3600 * 1000);
    query += ' after:' + Utilities.formatDate(tgl, 'Asia/Jakarta', 'yyyy/MM/dd');
  } else {
    query += ' newer_than:90d'; // sync pertama
  }

  const idAda = ambilIdAda_(sheetT);
  const emailDitolakAda = ambilEmailIdDitolak_(sheetD);

  const barisBaru = [];
  const barisDitolak = [];
  const sekarang = new Date();
  let dilewati = 0;

  const threads = cariSemuaThread_(query);
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const h = parseWondr(msg);

      if (!h) {
        const eid = msg.getId();
        if (!emailDitolakAda[eid]) {
          emailDitolakAda[eid] = true;
          barisDitolak.push([
            eid, 'wondr', msg.getSubject(),
            'subject tidak dikenali / nominal tidak ditemukan', sekarang
          ]);
        }
        return;
      }

      const id = String(h.id);
      if (idAda[id]) { dilewati++; return; }
      idAda[id] = true;

      barisBaru.push([
        id, h.tanggal, h.jumlah, h.biaya, h.arah, h.akun, h.merchant,
        h.detail_penerima || '', h.kategori || '', h.tipe,
        h.sumber_subject, h.email_id, sekarang
      ]);
    });
  });

  if (barisBaru.length) {
    sheetT.getRange(sheetT.getLastRow() + 1, 1, barisBaru.length, KOLOM_TRANSAKSI.length)
          .setValues(barisBaru);
  }
  if (barisDitolak.length) {
    sheetD.getRange(sheetD.getLastRow() + 1, 1, barisDitolak.length, KOLOM_DITOLAK.length)
          .setValues(barisDitolak);
  }

  props.setProperty('WONDR_SYNC_TERAKHIR', sekarang.toISOString());
  Logger.log('Sync wondr selesai. Baru: ' + barisBaru.length +
             ' | Sudah ada: ' + dilewati +
             ' | Ditolak baru: ' + barisDitolak.length);
}

/** Backfill panjang: panggil manual, mis. syncWondr(365). */
function backfillWondr90() { syncWondr(90); }

/**
 * Backfill seluruh histori (±3 tahun sekaligus, paginasi otomatis).
 * Kalau kena timeout ("Exceeded maximum execution time"), bagi per tahun:
 * syncLivin(365) lalu syncLivin(365, 365) lalu syncLivin(365, 730), dst.
 * Dedup menjamin aman diulang / overlap.
 */
function backfillSemuaLivin() { syncLivin(1095); }
function backfillSemuaBibit() { syncBibit(1095); }
function backfillSemuaWondr() { syncWondr(1095); }