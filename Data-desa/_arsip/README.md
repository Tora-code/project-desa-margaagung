# Folder _arsip

Isi folder ini **tidak dipakai** oleh alur layanan surat sehari-hari.
Semua file hanya dipindahkan (bukan dihapus), jadi masih bisa dipakai kapan saja.

## sumber-excel/
File sumber `.xls` (Daftar KK Prodeskel Dusun I–VI) untuk membangun ulang database.
- Dipakai oleh `..\bangun_database.bat` — script sudah diatur agar tetap mencari file di folder ini.
- File `...Dusun I acak2.xls` (file uji) sudah dibuang 2026-11-02; salinannya ada di `..\..\_backup\bersih-*.zip`.

## laporan/
- `laporan_bangun_database.txt` = log hasil bangun database (dibuat ulang tiap build).
- `NIK_PERLU_DIPERBAIKI.csv`, `tanggal_diperbaiki_otomatis.csv`, `tanggal_perlu_dicek.csv` = laporan audit build.

---

**Yang tetap di `Data-desa` (jangan dipindah):**

- Data & pencarian: `database_kk.csv`, `cari_warga.ps1`, `admin_wa.csv`, `cek_admin.ps1`
- Koreksi data: `catat_koreksi.ps1`, `lihat_koreksi.ps1`, `setujui_koreksi.ps1`, `terapkan_koreksi.ps1`
- Layanan surat: `status_surat.ps1`, `siapkan_kirim.ps1`, `kabari_warga.ps1`, `dashboard_admin.ps1`
- Notifikasi otomatis: `kabari_antrean_admin.ps1`, `ringkasan_antrean_pagi.ps1`
- Bangun ulang database: `bangun_database.bat`, `bangun_database.py`
- Backup: `_backup\`

**Panduan lengkap layanan** (fitur, alur, peran & tugas, cara pakai): lihat
`PANDUAN-SERAH-TERIMA.md` di folder utama proyek.

Dibuat: 2026-10-27 · Diperbarui: 2026-11-02 (pembersihan + panduan serah terima)
