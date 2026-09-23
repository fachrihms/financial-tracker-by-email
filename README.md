# Finance Tracker by Email

Catat pengeluaran otomatis dari email notifikasi bank/e-wallet, simpan ke
Google Sheets, tampilkan sebagai dashboard.

**Alur:** Gmail → Apps Script (parse) → Spreadsheet → Dashboard web app.

**Sumber yang didukung:**

| Sumber | Email | Hasil |
|---|---|---|
| Livin' by Mandiri | `noreply.livin@bankmandiri.co.id` | belanja / top-up / transfer |
| wondr by BNI | `wondr@bni.co.id` | top-up / belanja (VA) / transfer |
| Bibit | `no-reply@bibit.id` | tabungan (reksa dana, terpisah dari pengeluaran) |

**File project:**

| File lokal | Jadi apa di Apps Script |
|---|---|
| `Parser.js` | file script `Parser` |
| `dashboard.js` | file script `dashboard` |
| `index.html` | file HTML `Index` (nama harus persis `Index`) |

---

## Prasyarat

- Akun Google + Gmail yang menerima notifikasi transaksi.
- Notifikasi email Livin' / wondr / Bibit dalam keadaan aktif.
- (Opsional, cara B) Node.js + `clasp` kalau mau push dari lokal.

## Cara pakai dari nol

### 1. Bikin spreadsheet kosong

Buat Google Sheets baru. Ini yang jadi database (semua tab dibikin otomatis
oleh `setup()`).

### 2. Buka Apps Script

Di spreadsheet: **Ekstensi → Apps Script**. Project ini **harus bound**
(menempel ke spreadsheet) karena kode memakai `getActiveSpreadsheet()`.
Jangan bikin project standalone dari script.google.com.

### 3. Masukkan 3 file project (pilih salah satu)

#### Cara A — copy-paste (paling gampang)

1. Ganti isi `Code.gs` dengan isi `Parser.js`, lalu rename file-nya jadi
   `Parser` (klik titik tiga → Rename).
2. Tombol `+` → **Script** → nama `dashboard` → paste isi `dashboard.js`.
3. Tombol `+` → **HTML** → nama `Index` → paste isi `index.html`.
4. Save (Ctrl+S).

#### Cara B — `clasp push` (kayak import, tanpa copy-paste)

```bash
npm i -g @google/clasp
clasp login
```

Ambil **Script ID** dari Project Settings di Apps Script, lalu:

```bash
clasp clone <SCRIPT_ID>
# copy Parser.js, dashboard.js, index.html ke folder itu, lalu:
clasp push
```

File `.js` otomatis jadi file script, `.html` jadi file HTML. Ke depannya
edit lokal → `clasp push` lagi.

### 4. Run `setup()` — sekali

Pilih fungsi `setup` → **Run**. Setujui akses yang diminta (Gmail read,
Sheets). Ini membuat 6 tab + format kolom + seed kategori:

- `transaksi` — semua transaksi Livin + wondr (kategori dihitung saat
  dashboard dibuka, bukan disimpan).
- `pemasukan` — diisi manual dari dashboard.
- `tabungan` — pembelian reksa dana Bibit.
- `config` — aturan `kategori, kata_kunci` (boleh diedit bebas).
- `koreksi` — koreksi kategori per merchant (diisi dari dashboard).
- `log_ditolak` — email yang gagal di-parse + alasannya.

### 5. Tes parser dulu

Run `tesParserLivin` / `tesParserWondr` / `tesParserBibit`, lihat hasilnya
di **Log**. Harusnya ada baris JSON transaksi + ringkasan
`Berhasil: X | Ditolak: Y`. Kalau 0 berhasil, jangan lanjut ke sync —
cek format email / kabarin yang bikin parser.

### 6. Tarik histori

| Cakupan | Fungsi |
|---|---|
| 90 hari terakhir | `backfill90`, `backfillWondr90`, `backfillBibit90` |
| ±3 tahun sekaligus | `backfillSemuaLivin`, `backfillSemuaWondr`, `backfillSemuaBibit` |

Dedup by id transaksi, jadi aman diulang. Kalau kena
`Exceeded maximum execution time` (histori ribuan email), bagi per tahun:
`syncLivin(365)`, lalu `syncLivin(365, 365)`, lalu `syncLivin(365, 730)`
(berlaku juga buat `syncWondr` / `syncBibit`).

### 7. Otomatisasi — run `pasangTrigger()` sekali

Sejak ini `syncLivin` + `syncBibit` + `syncWondr` jalan sendiri tiap 6 jam
(inkremental, cuma narik email baru). Run ulang kapan pun aman — trigger
lama dihapus dulu sebelum bikin baru, jadi nggak dobel.

### 8. Deploy dashboard

**Deploy → New deployment → Web app** → Execute as: *Me* → Who has access:
*Only myself* (dashboard pribadi) → Deploy → buka URL-nya.

> Tiap habis ubah kode: **Manage deployments → edit (ikon pensil) →
> New version**. Kalau nggak, dashboard tetap jalanin versi lama.

---

## Isi dashboard

- Dropdown **Tahun** (2024–tahun berjalan) + **Bulan** (Jan–Des).
- KPI: Pemasukan · Pengeluaran riil · **Arus kas bersih**
  (= pemasukan − pengeluaran riil − tabungan) · Jumlah transaksi ·
  Rata-rata & **Median per hari** (median tahan outlier, dihitung dari hari
  bertransaksi) · Top-up · Menabung.
- Grafik Per kategori & Per tipe + label persen, Top 10 merchant.
- Kartu **Top-up e-wallet**: breakdown DANA / GoPay / ShopeePay / OVO /
  LinkAja warna brand + catatan fungsi. **Bisa diklik** buat lihat detail
  transaksi per wallet.
- Tabel semua transaksi + filter + **koreksi kategori per merchant**
  (tersimpan di tab `koreksi`, berlaku ke semua transaksi merchant itu).
- Tambah / hapus pemasukan manual langsung dari dashboard.

## Kategori & keyword

Prioritas: **koreksi manual > aturan `config`**. Transaksi tipe `belanja`
yang tidak cocok keyword apa pun otomatis jadi `Belanja`.

Edit tab `config` sesukanya (`kata_kunci` dipisah koma, huruf kecil).
Urutan baris = urutan pengecekan, jadi taruh `Belanja Makan` di atas
`Belanja` supaya `martabak` tidak kemakan keyword `mart`.

## Cheat sheet fungsi

**Yang boleh di-run manual (tanpa parameter wajib):**
`setup` · `tesParser*` · `backfill*` / `sync*` · `pasangTrigger`

**Jangan di-run manual (helper, dipanggil fungsi lain):**
`parse*` · `ambil*` · `cari*` · `tentukan*` · apa pun berakhiran `_`
(mereka butuh parameter objek email).

## Troubleshooting

| Gejala | Cek |
|---|---|
| Transaksi tidak masuk | tab `log_ditolak`, kolom `alasan` |
| Kode diubah tapi dashboard sama | belum deploy **New version** |
| Trigger jalan dobel | run `pasangTrigger()` lagi (idempoten) |
| Email gagal/transaksi gagal | otomatis ditolak, bukan error |
| Butuh histori > 90 hari | `backfillSemua*` / chunk `syncX(365, offset)` |
