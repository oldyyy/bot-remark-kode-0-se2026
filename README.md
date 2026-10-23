# Bot Remark #0 NTB — FASIH SE2026

Memberi catatan **`#0`** pada rincian **27.a "Nilai produksi/pendapatan/penjualan barang dan jasa"**
(31.a bila usaha hanya mengisi bulanan) untuk usaha yang keputusan dashboard-nya #0 tetapi di FASIH
belum ada remark #0, lalu **Kirim**.

## Sasaran (dibaca otomatis dari Excel)

File: `rekap_remark_ntb_kode0_*.xlsx` (sheet `Remark #0`), yang **terbaru** di:
1. folder `input/` di sini, atau
2. `..\..\fasih-scraper\output\rasio_ntb\`, atau
3. `--file "path\ke\file.xlsx"`.

Satu baris = satu usaha (`kunci` = assignment_id|index1). Diproses bila:
- `remark_dashboard` = `#0` dan `remark_fasih` belum memuat `#0`, atau
- `remark_dashboard` = `#DC_00` dan `jumlah_remark_fasih` = 0.

## Alur per assignment

1. Buka `/edit`. Terkunci (tanpa tombol Kirim) → **Revoke** → buka ulang (hanya mode live).
2. Blok **SE2026 - L BLOK II** → kartu usaha yang namanya cocok dengan `nama_usaha`.
3. Rincian 27.a (atau 31.a) → ikon **Catatan** → bila `#0` sudah ada, dilewati; bila belum, tulis `#0` → Simpan.
4. Ada catatan baru → **Kirim** (galat "Tampilkan Anomali" diperbaiki otomatis; galat lain → dilewati & dicatat).

Jeda acak antar langkah & antar assignment, istirahat berkala.

## Persiapan (sekali per PC)

1. Pasang **Python 3.10+** (centang *Add to PATH*) dan **Google Chrome**.
2. Klik dua kali `setup.bat` (memasang `playwright` & `openpyxl`).

## Menjalankan

```bat
buka_chrome.bat                     :: Chrome port 9340, login SSO FASIH manual di jendela itu
py bot.py --dry-run --limit 3       :: uji: hanya memeriksa, tidak menulis/mengirim
py bot.py --live --limit 3          :: eksekusi 3 assignment dulu, cek hasilnya di FASIH
py bot.py --live                    :: semua sisa
py bot.py --status                  :: progres + logs\sisa_belum_selesai.csv
```

**Berhenti di tengah**: tekan `Ctrl+C`. Yang sudah terkirim tercatat; saat dijalankan lagi bot melanjutkan
sisa (usaha berstatus SELESAI / SUDAH_ADA dilewati, `--ulang` untuk memaksa).

**Paralel** (lebih cepat; **2 jendela** disarankan — 4 jendela terbukti membuat PC/server terlalu berat):
buka Chrome tiap port dan **login di tiap jendela** (sesi tidak disalin), lalu di jendela CMD terpisah:
```bat
buka_chrome.bat 9341
py bot.py --live --port 9340
py bot.py --live --port 9341 --mundur
```
Bot kedua mengerjakan dari akhir daftar. Tiap bot menandai assignment yang sedang dikerjakannya
(`logs\sedang_dikerjakan\`) dan memeriksa log sebelum tiap assignment, jadi tidak saling tabrak dan
bot yang berhenti bisa dijalankan lagi kapan saja. (`--shard K/N` juga tetap bisa dipakai.)

**Sesi login habis** di tengah jalan: bot berhenti dengan pesan "Sesi login habis" → login ulang di
jendela Chrome port itu → jalankan perintah yang sama lagi.

Opsi lain: `--ids a,b` (assignment tertentu), `--cepat` (jeda antar assignment diperpendek, untuk uji).

## Status di log (`logs\`)

| Status | Arti |
|---|---|
| `SELESAI` | `#0` ditulis & assignment terkirim |
| `SUDAH_ADA` | `#0` sudah ada di rincian itu, tidak ditulis ulang |
| `DRYRUN_OK` | (dry-run) siap ditulis |
| `PERLU_MANUAL` | kartu usaha tidak bisa dipastikan (nama tidak cocok) → kerjakan manual |
| `PERLU_REVOKE` | (dry-run) assignment terkunci, saat live akan di-revoke |
| `GALAT_TIDAK_DIKENAL` | kirim gagal karena galat lain di isian → screenshot `galat_<id>.png` |
| `ERROR` | gagal teknis → screenshot `error_<id>.png`; ikut diproses lagi saat bot dijalankan ulang |

Bot berhenti sendiri bila sesi login habis, muncul CAPTCHA/verifikasi, atau 3 error berturut-turut.

> Folder `logs\`, `input\` dan profil Chrome (`.profil_chrome_*`) berisi data/sesi — jangan dibagikan.
