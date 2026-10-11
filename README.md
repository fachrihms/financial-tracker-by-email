# Finance Tracker by Email

Catat pengeluaran otomatis dari email notifikasi bank/e-wallet, simpan ke
Google Sheets, tampilkan sebagai dashboard.
![Dashboard Preview](assets/tutorial_11.png)

**Alur:** Gmail → Apps Script (parse) → Spreadsheet → Dashboard web app.

**Sumber yang didukung:**

| Sumber | Email | Hasil |
|---|---|---|
| Livin' by Mandiri | `noreply.livin@bankmandiri.co.id` | belanja / top-up / transfer |
| wondr by BNI | `wondr@bni.co.id` | top-up / belanja (VA) / transfer |
| Bibit | `no-reply@bibit.id` | tabungan (reksa dana, terpisah dari pengeluaran) |
| GoPay | `no-reply@customers.go-pay.co.id` | struk tagihan (mis. PLN Token → belanja, digabung dengan transfer bank yang nominal+tanggalnya sama) |

**File project:**

| File lokal | Jadi apa di Apps Script |
|---|---|
| `Parser.js` | file script `Parser` |
| `dashboard.js` | file script `dashboard` |
| `index.html` | file HTML `Index` (nama harus persis `Index`) |

---

## Prasyarat

- Akun Google + Gmail yang menerima notifikasi transaksi.
- Notifikasi email Livin' / wondr / Bibit / GoPay dalam keadaan aktif.
- (Opsional, cara B) Node.js + `clasp` kalau mau push dari lokal.

## Cara pakai dari nol

### 1. Bikin spreadsheet kosong

Buat Google Sheets baru. Ini yang jadi database (semua tab dibikin otomatis
oleh `setup()`).


### 2. Buka Apps Script

Di spreadsheet: **Ekstensi → Apps Script**. Project ini **harus bound**
(menempel ke spreadsheet) karena kode memakai `getActiveSpreadsheet()`.

Jangan bikin project standalone dari script.google.com!

![setup google sheets](assets/tutorial_1.png)

### 3. Masukkan 3 file project (pilih salah satu)

#### Cara A — copy-paste (paling gampang)

1. Ganti isi `Code.gs` dengan isi `Parser.js`, lalu rename file-nya jadi
   `Parser` (klik titik tiga → Rename).
2. Tombol `+` → **Script** → nama `dashboard` → paste isi `dashboard.js`.
3. Tombol `+` → **HTML** → nama `Index` → paste isi `index.html`.
4. Save (Ctrl+S).

![buat files](assets/tutorial_2.png)

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

- `transaksi` — semua transaksi Livin + wondr + GoPay (kategori dihitung saat
  dashboard dibuka, bukan disimpan).
- `pemasukan` — diisi manual dari dashboard.
- `manual` — transaksi cash manual (diisi + bisa dihapus dari dashboard).
- `catatan` — keterangan bebas per transaksi email (diisi dari dashboard).
- `tabungan` — pembelian reksa dana Bibit.
- `config` — aturan `kategori, kata_kunci` (boleh diedit bebas).
- `koreksi` — koreksi kategori per merchant (diisi dari dashboard).
- `log_ditolak` — email yang gagal di-parse + alasannya.

![setup data](assets/tutorial_3.png)

### 5. Tes parser dulu

Run `tesParserLivin` / `tesParserWondr` / `tesParserBibit` / `tesParserGoPay`,
lihat hasilnya
di **Log**. Harusnya ada baris JSON transaksi + ringkasan
`Berhasil: X | Ditolak: Y`. Kalau 0 berhasil, jangan lanjut ke sync —
cek format email / kabarin yang bikin parser.

### 6. Tarik histori

| Cakupan | Fungsi |
|---|---|
| 90 hari terakhir | `backfill90`, `backfillWondr90`, `backfillBibit90`, `backfillGoPay90` |
| ±3 tahun sekaligus | `backfillSemuaLivin`, `backfillSemuaWondr`, `backfillSemuaBibit`, `backfillSemuaGoPay` |

Dedup by id transaksi, jadi aman diulang. Struk GoPay yang nominal +
tanggalnya sama dengan transaksi bank digabung jadi satu (tidak dobel).
Kalau kena
`Exceeded maximum execution time` (histori ribuan email), bagi per tahun:
`syncLivin(365)`, lalu `syncLivin(365, 365)`, lalu `syncLivin(365, 730)`
(berlaku juga buat `syncWondr` / `syncBibit` / `syncGoPay`).

![tarik data historis](assets/tutorial_4.png)

### 7. Otomatisasi — run `pasangTrigger()` sekali

Sejak ini `syncLivin` + `syncBibit` + `syncWondr` + `syncGoPay` jalan sendiri tiap 6 jam
(inkremental, cuma narik email baru). Run ulang kapan pun aman — trigger
lama dihapus dulu sebelum bikin baru, jadi nggak dobel.

![set trigger](assets/tutorial_5.png)

### 8. Deploy dashboard

![deploy dashboard](assets/tutorial_6.png)
**Deploy → New deployment → Web app** → Execute as: *Me* → Who has access:
![deploy dashboard](assets/tutorial_7.png)
*Only myself* (dashboard pribadi) → Deploy → buka URL-nya.
![link dashboard](assets/tutorial_8.png)

> Tiap habis ubah kode: **Manage deployments → edit (ikon pensil) →
> New version**. Kalau nggak, dashboard tetap jalanin versi lama.


### 9. Bikin akun login (sekali)

Dashboard dilindungi halaman login (username + password). Kredensial
**tidak ditulis di kode maupun Sheets** — disimpan sebagai hash SHA-256 +
salt di ScriptProperties.

1. Di Apps Script, run fungsi **`buatAkun()`** → isi username, password
   (min 6 karakter), dan email pemulihan lewat prompt yang muncul.
> Note: ketika fungsi buatAkun() di run maka akan muncul pop up di spreadsheet untuk input username dan password
> isi input tersebut untuk keperluan login ke dashboard
![buat akun](assets/tutorial_9.png)
3. Buka URL dashboard → login.
![login](assets/tutorial_10.png)
4. Lupa password → klik **Lupa password** → kode 6 digit (berlaku
   15 menit) dikirim ke email pemulihan → tukar dengan password baru.

Sesi login berlaku 7 hari (tersimpan di browser). Tombol **Keluar** di
header untuk logout manual. Semua fungsi data menolak request tanpa token
sesi yang valid.

> Catatan: ini kunci level aplikasi buat pemakaian pribadi. Untuk lapis
> tambahan, biarkan Who has access = *Only myself* (login Google) saat
> deploy.

---

## Isi dashboard

- Dropdown **Tahun** (2024–tahun berjalan) + **Bulan** (Jan–Des).
- KPI: Pemasukan · Pengeluaran riil · **Arus kas bersih**
  (= pemasukan − pengeluaran riil) · Jumlah transaksi ·
  Rata-rata & **Median per hari** (median tahan outlier, dihitung dari hari
  bertransaksi) · Top-up · Menabung.
- Grafik Per kategori & Per tipe + label persen, Top 10 merchant.
- Kartu **Top-up e-wallet**: breakdown DANA / GoPay / ShopeePay / OVO /
  LinkAja warna brand + catatan fungsi. **Bisa diklik** buat lihat detail
  transaksi per wallet. Tiap transaksi bisa diubah tipenya:
  **Belanja/Transfer (pengeluaran)** vs **Top-up (disimpan)** + kategorinya
  + keterangan bebas — yang sudah diubah jadi pengeluaran tetap tampil di sini.
- Kartu **Riwayat menabung**: daftar nabung dari email yang masuk (mis. Bibit),
  dikelompokkan per produk reksa dana. **Bisa diklik** buat lihat detail
  tiap transaksi.
- Tabel semua transaksi + filter + **edit tipe per transaksi**
  (langsung ubah kolom `tipe` di tab `transaksi`, sync tidak menimpa)
  + **keterangan bebas per transaksi** (mis. top-up buat "bayar kos",
  tersimpan di tab `catatan`, bisa dicari lewat kolom filter)
  + **koreksi kategori per merchant**
  (tersimpan di tab `koreksi`, berlaku ke semua transaksi merchant itu).
  Di atas tabel ada keterangan **Terakhir diperbarui** per sumber
  (Livin / wondr / Bibit) — diambil dari waktu sync terakhir.
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

## Catatan builder
Buka link dashboard  di incognito mode kalau di browser ada 
akun gmail yang lebih dari satu 

## Troubleshooting

| Gejala | Cek |
|---|---|
| Transaksi tidak masuk | tab `log_ditolak`, kolom `alasan` |
| Kode diubah tapi dashboard sama | belum deploy **New version** |
| Trigger jalan dobel | run `pasangTrigger()` lagi (idempoten) |
| Email gagal/transaksi gagal | otomatis ditolak, bukan error |
| Butuh histori > 90 hari | `backfillSemua*` / chunk `syncX(365, offset)` |
