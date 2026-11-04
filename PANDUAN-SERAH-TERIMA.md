# PANDUAN SERAH TERIMA — Layanan Surat Desa Marga Agung

Versi: 2 November 2026
Untuk: petugas desa / orang yang akan menjalankan layanan ini
Dibaca dari atas ke bawah: ±20 menit. Bagian 1, 4, 5, dan 13 cukup dibaca dulu
kalau waktunya sedikit.

---

## 1. Apa ini, dan siapa yang memakainya

Satu nomor **WhatsApp desa** yang melayani surat-menyurat warga secara otomatis,
dan satu menu khusus untuk petugas desa di nomor yang sama.

| Siapa | Bisa apa | Tidak bisa apa |
|---|---|---|
| **Warga** | Minta 12 jenis surat, cek status suratnya, minta dihubungkan ke petugas | Melihat data warga lain, melihat menu petugas, meminta surat di luar 12 daftar |
| **Petugas desa** (nomor terdaftar) | Lihat antrean, menyetujui surat, mengabari warga, rekap, mencari data warga | Mengubah database warga langsung dari chat |
| **Operator teknis** | Menjalankan/memperbaiki program, backup, membangun ulang database | — |

Yang menjawab warga adalah asisten otomatis (Hermes Agent dengan model DeepSeek,
profil `surat-desa`). Aturan gaya bahasa dan menu ada di berkas **SOUL** profil itu,
bukan di folder proyek ini.

Surat dibuat dari **blanko resmi desa** (12 berkas Word di `Template-surat/`), jadi
kop, logo, dan tata letaknya tidak berubah.

---

## 2. Peta folder — mana yang boleh disentuh, mana yang jangan

Folder proyek:
`C:\Users\Sulthonu Tora Aulia\OneDrive\Dokumen\Project Desa MargaAgung`

| Berkas / folder | Isinya | Boleh disentuh? |
|---|---|---|
| `Data-desa\database_kk.csv` | Database warga: 4.500 baris (4.498 warga asli + 2 warga uji) | **JANGAN** diedit manual. Ubah lewat antrean koreksi |
| `Data-desa\admin_wa.csv` | Daftar petugas: nomor, nama, jabatan, aktif | Boleh, kalau ada petugas baru. Hati-hati: ini penentu siapa yang dapat menu admin |
| `Data-desa\cari_warga.ps1` | Pencari data warga (wajib nama + NIK) | Tidak perlu diubah |
| `Data-desa\cek_admin.ps1` | Pemeriksa nomor petugas | Tidak perlu diubah |
| `Data-desa\status_surat.ps1` | Menyetujui / menandai siap diambil / sudah diambil | Dipakai petugas |
| `Data-desa\siapkan_kirim.ps1` | Menyiapkan dokumen surat untuk dikirim ke WhatsApp | Dipakai petugas |
| `Data-desa\kabari_warga.ps1` | Mengirim kabar ke warga bahwa surat siap diambil | Dipakai petugas |
| `Data-desa\kabari_antrean_admin.ps1` | Notifikasi antrean baru ke petugas | Berjalan otomatis |
| `Data-desa\ringkasan_antrean_pagi.ps1` | Ringkasan surat belum diproses (jam 09.00) | Berjalan otomatis |
| `Data-desa\dashboard_admin.ps1` | Rekap untuk menu admin | Dipakai petugas |
| `Data-desa\catat_koreksi.ps1`, `lihat_koreksi.ps1`, `setujui_koreksi.ps1`, `terapkan_koreksi.ps1` | Antrean perbaikan data NIK | Dipakai petugas |
| `Data-desa\bangun_database.bat` | Membangun ulang database dari file Excel | Operator teknis, sesekali |
| `Data-desa\_backup\` | Backup database bertanggal | **JANGAN dihapus** |
| `Data-desa\_arsip\sumber-excel\` | 6 file Excel resmi Dusun I–VI | **JANGAN dihapus** — bahan bangun ulang database |
| `Data-desa\_arsip\laporan\` | Laporan hasil build + daftar NIK perlu diperbaiki | Bahan pemeriksaan |
| `Template-surat\` | 12 blanko surat resmi | **JANGAN diubah** isinya |
| `Antrean-surat\antrean_surat.csv` | Antrean surat (jantung sistem) | Diubah lewat skrip, bukan manual |
| `Antrean-surat\antrean_ambil.csv` | Daftar surat yang menunggu diambil warga | Diubah lewat skrip |
| `Antrean-surat\notifikasi_antrean.csv` | Catatan siapa sudah dikabari | Otomatis |
| `Antrean-surat\notifikasi_antrean.txt` | Salinan persis pesan notifikasi yang dikirim | Bukti/jejak |
| `Antrean-surat\_kirim\` | Salinan dokumen bernama tanpa spasi (siap kirim WA) | Otomatis dibersihkan/diisi sendiri |
| `Antrean-surat\_backup\` | Backup antrean bertanggal | **JANGAN dihapus** |
| `ratings.csv` | Penilaian warga (1–10) per surat | Ditambah otomatis |
| `_backup\bersih-*.zip` | Arsip berkas lama sebelum dibersihkan | Simpan |

---

## 3. Fitur yang tersedia

### Untuk warga (menu 14 pilihan)

| No | Pilihan | No | Pilihan |
|---|---|---|---|
| 1 | Surat Keterangan Domisili | 8 | Surat Keterangan E-KTP |
| 2 | Surat Keterangan Tidak Mampu | 9 | Surat Keterangan Suami/Isteri |
| 3 | Surat Keterangan Usaha | 10 | Surat Keterangan Jual Beli |
| 4 | Surat Keterangan Kelahiran | 11 | Surat Mohon Izin Keramaian |
| 5 | Surat Keterangan Kematian | 12 | Surat Bepergian |
| 6 | Surat Keterangan Pindah | 13 | Cek status surat saya |
| 7 | Surat Keterangan Beda Nama | 14 | Tanya petugas desa |

Warga membalas angka, bot menanyakan **nama lengkap + NIK** (sekaligus, satu pesan),
lalu otomatis mengambil data dari database desa, membuat draf surat, meminta
konfirmasi (`KONFIRMASI`), meneruskan ke antrean petugas, dan memberi nomor surat
untuk cek status. Di akhir layanan warga diminta penilaian 1–10.

### Untuk petugas desa (menu 6 pilihan)

| No | Pilihan | Gunanya |
|---|---|---|
| 1 | Antrean menunggu | Daftar surat baru + dokumennya ikut dikirim |
| 2 | Setujui surat | Menyetujui surat (wajib dari nomor petugas terdaftar) |
| 3 | Kabari warga siap diambil | Mengirim WA ke warga bahwa suratnya bisa diambil |
| 4 | Antrean pengambilan | Siapa saja yang belum mengambil suratnya |
| 5 | Rekap hari ini | Jumlah surat per status hari ini |
| 6 | Cari data warga | Mencari data warga dari nama + NIK |

### Yang berjalan sendiri tanpa disuruh

| Fitur | Jalan kapan | Isinya |
|---|---|---|
| Notifikasi antrean baru | **tiap 5 menit, 24 jam** | 3 baris: nomor surat, jenis surat, nama + jam masuk. Tanpa NIK, tanpa alamat |
| Ringkasan pagi | **tiap hari jam 09.00** | Daftar semua surat yang belum disetujui petugas, dari yang paling lama menunggu |

Jam tenang 21.00–07.00 **dimatikan** atas permintaan pemilik (2 Nov 2026), jadi
notifikasi keluar kapan pun. Kalau nanti ingin dihidupkan lagi, jalankan skrip
notifikasi dengan tambahan `-PakaiJamTenang`.

---

## 4. Alur kerja lengkap

### Dari sisi warga

1. Warga kirim pesan → bot mengirim menu 14 pilihan.
2. Warga balas angka (mis. `1`) → bot minta nama lengkap + NIK.
3. Bot memeriksa data ke database desa dan menampilkan ringkasan data warga.
4. Bot menanyakan nomor HP dan keperluan surat.
5. Bot membuat draf surat dan meminta konfirmasi (`KONFIRMASI`).
6. Setelah dikonfirmasi, surat masuk antrean dan warga menerima **nomor surat**
   (contoh `SR-20261102-001`) untuk cek status.
7. Bot meminta penilaian 1–10.

Kalau data tidak cocok (nama beda dengan NIK, nama kembar, atau NIK di database
bukan 16 digit), bot **berhenti dan tidak membuat surat** — warga diarahkan ke
petugas desa. Ini disengaja: lebih baik menolak daripada menulis identitas yang salah
di surat resmi.

### Dari sisi petugas desa

1. Petugas menerima notifikasi antrean baru (3 baris) di WhatsApp.
2. Pilih menu `1` untuk melihat antrean beserta dokumen suratnya.
3. Setelah diperiksa: **setujui** surat (menu `2`), lalu tentukan bisa diambil atau belum.
4. Kalau sudah bisa: petugas menyebut **tanggal mulai bisa diambil**, lalu sistem
   mengabari warga otomatis (menu `3`).
5. Petugas desa mengisi nomor surat dan menandatangani suratnya.
6. Warga datang mengambil surat (bawa KTP asli), lalu status ditandai **sudah diambil**.

Contoh satu baris lengkap: warga mengajukan → antrean → disetujui petugas →
"siap diambil" + tanggal → warga dikabari → warga mengambil → "sudah diambil".

---

## 5. Peran dan tugas

| Peran | Yang harus dilakukan | Yang dilarang |
|---|---|---|
| **Petugas piket WhatsApp** | Menjawab pertanyaan yang diteruskan bot, memeriksa antrean tiap pagi, menyetujui surat | Menyetujui surat atas nama petugas lain; memberi dokumen surat ke warga lain |
| **Kepala desa / petugas tanda tangan** | Mengisi nomor surat, menandatangani, mengembalikan surat ke warga | — |
| **Operator teknis** | Memastikan program dan jembatan WhatsApp hidup, backup berkala, membangun ulang database saat data Excel berubah | Mengedit `database_kk.csv` langsung; menghapus folder `_backup` |

Siapa yang dianggap petugas ditentukan **hanya** oleh `Data-desa\admin_wa.csv`
(kolom `aktif` = 1). Nama tampilan di WhatsApp tidak berpengaruh — warga yang mengganti
namanya menjadi "Admin Desa" tetap warga.

---

## 6. Rutinitas yang harus dikerjakan

**Setiap hari**
- [ ] Baca ringkasan pagi jam 09.00, lalu proses surat yang menunggu.
- [ ] Setujui surat yang datanya sudah benar; isi nomor surat & tanda tangan.
- [ ] Kabari warga yang suratnya sudah bisa diambil.
- [ ] Kirim pesan percobaan ke nomor layanan (`halo`) untuk memastikan bot menjawab.

**Yang tidak boleh dilakukan**
- Jangan menyalakan aplikasi Hermes dua kali (dua jendela) — bisa berebut jalur
  WhatsApp dan membuat bot bisu (lihat bagian 9).

**Setiap minggu**
- [ ] Buka menu `4` (antrean pengambilan): tagih warga yang belum mengambil surat.
- [ ] Periksa menu `5` (rekap) untuk memastikan tidak ada surat "nyangkut" lama.

**Setiap bulan**
- [ ] Cadangkan folder proyek (salin ke flashdisk atau simpan zip).
- [ ] Periksa `Data-desa\_arsip\laporan\NIK_PERLU_DIPERBAIKI.csv` — NIK yang salah
      ketik di file Excel sebaiknya dibetulkan di sumbernya.

**Sekali waktu / saat data berubah**
- [ ] Kalau file Excel kependudukan diperbarui: bangun ulang database (bagian 11).

---

## 7. Aturan data & privasi (tidak bisa ditawar)

1. **Database warga tidak pernah diubah dari chat.** Perubahan hanya lewat antrean
   koreksi + persetujuan petugas.
2. **Satu warga, satu kali layanan.** Bot hanya menampilkan data warga yang sedang
   dilayani, tidak pernah menampilkan daftar warga lain.
3. **Notifikasi ke petugas tidak memuat NIK, alamat, keperluan, atau isi surat.**
   Skrip memeriksa sendiri pesannya sebelum dikirim; kalau data itu terbawa,
   pengiriman dibatalkan.
4. **Warga tidak pernah dikirimi dokumen surat.** Dokumen hanya untuk petugas.
5. **Jangan pernah meminta PIN, OTP, kata sandi, atau nomor rekening** ke warga.
6. **Di chat WhatsApp tidak boleh muncul**: nama berkas/folder, nama skrip, kode
   status mesin, nama penyetuju, dan waktu persetujuan. Semua itu hanya boleh ada di
   berkas, bukan di pesan.

---

## 8. Pemeriksaan cepat: apakah semuanya sehat?

| Yang diperiksa | Cara | Hasil yang benar |
|---|---|---|
| Program hidup? | `hermes gateway status` | running |
| Jembatan WhatsApp tersambung? | Cek apakah bot membalas pesan warga | bot membalas menu |
| Antrean & rekap | Balas `5` dari nomor petugas | jumlah per status |
| Notifikasi jalan? | `Antrean-surat\notifikasi_antrean.txt` | ada jejak pesan terakhir |
| Jadwal otomatis terdaftar? | minta operator: daftar jadwal Hermes | 2 jadwal: tiap 5 menit & jam 09.00 |
| Database utuh? | `Data-desa\database_kk.csv` | 4.501 baris (1 header + 4.500 warga) |

---

## 9. Kalau ada masalah (chat devloper saja kelompok 8)

| Gejala | Kemungkinan penyebab | Tindakan |
|---|---|---|
| Bot tidak membalas sama sekali | Layanan/jembatan WhatsApp berhenti | `hermes gateway status`, lalu `hermes gateway restart` |
| Bot bisu tepat setelah aplikasi dinyalakan (atau dinyalakan dua kali) | Dua jembatan WhatsApp berebut satu jalur (port 3001). Yang kalah gagal jalan, dan layanan berhenti mendengarkan WhatsApp. Di catatan muncul: `WhatsApp bridge process exited unexpectedly (code 1)` lalu `Fatal whatsapp adapter error` | 1) Tutup aplikasi Hermes sepenuhnya — pastikan tidak ada dua jendela yang jalan. 2) Jalankan `hermes gateway restart`. 3) Tunggu ±1 menit, lalu pastikan di `profiles\surat-desa\whatsapp\bridge.log` muncul `✅ WhatsApp connected!`. Catatan: yang mati hanya pesan **masuk**; pesan keluar (notifikasi) tetap bisa terkirim |
| WhatsApp terputus (minta scan) | Perangkat tertaut keluar | Jalankan `hermes whatsapp` dan scan QR dari HP desa |
| Bot tidak menemukan warga padahal ada | Nama/NIK salah ketik, atau NIK di database bukan 16 digit | Cek ulang; kalau tetap gagal, arahkan ke petugas. Jangan tambah data sendiri |
| Dokumen surat tidak terkirim ke petugas | Salinan dokumen belum dibuat | Minta petugas buka menu `1` lagi (sistem membuat salinan otomatis) |
| Notifikasi tidak masuk | Nomor petugas tidak aktif di `admin_wa.csv`, atau jadwal mati | Periksa kolom `aktif`; minta operator memeriksa jadwal |
| Surat tidak bisa disetujui | Persetujuan dilakukan dari nomor bukan petugas | Setujui dari nomor petugas terdaftar |
| Data warga di surat salah | Salah ketik di file Excel sumber | Catat lewat antrean koreksi, betulkan di Excel, bangun ulang database |

---

## 10. Backup dan pemulihan

- **Database**: `Data-desa\_backup\database_kk_*.csv` — dibuat otomatis sebelum
  perubahan data. Untuk memulihkan: salin berkas backup itu menimpa
  `Data-desa\database_kk.csv`.
- **Antrean**: `Antrean-surat\_backup\antrean_surat_*.csv` — dibuat otomatis setiap
  status surat berubah.
- **Cadangan menyeluruh**: `_backup\bersih-<tanggal>.zip` (2 Nov 2026) berisi berkas
  lama sebelum pembersihan. Simpan berkas ini.
- **Saran**: sebulan sekali salin seluruh folder proyek ke tempat lain (flashdisk /
  Google Drive desa). Jangan hanya mengandalkan OneDrive.

---

## 11. Membangun ulang database dari file Excel

Dilakukan kalau data kependudukan dari dusun diperbarui.

1. Letakkan file `.xls` resmi (Dusun I–VI) di `Data-desa\_arsip\sumber-excel\`
   (ganti yang lama).
2. Tutup Excel — file yang sedang dibuka tidak bisa ditimpa.
3. Dobel-klik `Data-desa\bangun_database.bat`.
4. Periksa `Data-desa\_arsip\laporan\laporan_bangun_database.txt`:
   harus 6 file resmi, dan jumlah warga wajar.
5. Hasil: `Data-desa\database_kk.csv` diperbarui. Skrip lain tidak perlu diubah.

Catatan: NIK yang salah ketik di Excel **tidak ditebak**. Data yang meragukan
dikarantina ke `tanggal_perlu_dicek.csv`, dan NIK bermasalah didaftar di
`NIK_PERLU_DIPERBAIKI.csv`.

---

## 12. Bagian teknis program (untuk operator)

- Program: **Hermes Agent**, profil bernama **`surat-desa`**
  (berkas profil: `%LOCALAPPDATA%\hermes\profiles\surat-desa`).
- Aturan bahasa & menu bot ada di berkas **`SOUL.md`** profil tersebut. Kalau menu
  diubah di SOUL, dokumen ini juga harus diperbarui (bagian 3).
- Dua jadwal otomatis:
  1. **Notifikasi antrean baru** — tiap 5 menit, 24 jam
     (menjalankan `Data-desa\kabari_antrean_admin.ps1`).
  2. **Ringkasan pagi** — tiap hari 09.00
     (menjalankan `Data-desa\ringkasan_antrean_pagi.ps1`).
- Perintah yang berguna: `hermes gateway status`, `hermes gateway restart`,
  `hermes whatsapp` (sambung ulang QR), `hermes --profile surat-desa send`.
- Catatan berkas: skrip `.ps1` yang memuat emoji wajib disimpan **UTF-8 dengan BOM**,
  kalau tidak Windows PowerShell gagal membacanya.

---

## 13. Kondisi saat serah terima (2 November 2026)

- **Antrean surat: 0** (sudah direset untuk serah terima). Berkas
  `antrean_surat.csv`, `antrean_ambil.csv`, dan `notifikasi_antrean.csv` hanya berisi
  baris judul.
- **Database: 4.500 baris** — 4.498 warga asli + **2 warga uji** yang sengaja
  disimpan untuk latihan:
  - `UJI BUDI SANTOSO` — NIK `1801999999990001`
  - `UJI SITI AMINAH` — NIK `1801999999990002`
  Keduanya ditandai `DATA UJI - HAPUS SETELAH PENGUJIAN` di kolom `Sumber`. Hapus dua
  baris ini kalau layanan sudah dinyatakan siap dipakai sungguhan.
- **Ratings**: `ratings.csv` masih memuat 3 penilaian lama dari masa pengujian.
- **Daftar petugas aktif**: RENDI (0812-xxxx-xxxx) dan TORA (0896-xxxx-xxxx) — sama
  dengan yang terdaftar di `admin_wa.csv`.

### Checklist serah terima (±15 menit)

- [ ] Nomor WhatsApp desa dibuka di HP desa, pastikan bot membalas `halo`.
- [ ] Nomor petugas terdaftar diuji: balas `5` (rekap) — harus keluar rekap, bukan menu warga.
- [ ] Coba satu pengajuan surat dari HP lain (boleh pakai warga uji di atas).
- [ ] Pastikan notifikasi antrean baru masuk ke HP petugas dalam ≤5 menit.
- [ ] Buka menu `1`, pastikan dokumen surat ikut terkirim.
- [ ] Setujui surat itu, tandai siap diambil, dan pastikan warga menerima kabar.
- [ ] Bersihkan uji coba: hapus baris surat uji + kembalikan antrean ke 0.
- [ ] Tunjukkan letak folder `_backup` dan `_arsip` kepada penerima.
- [ ] Serahkan dokumen ini + catat siapa operator teknis yang bisa dihubungi.
