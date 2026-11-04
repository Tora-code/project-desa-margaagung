# Project Desa MargaAgung

Sistem administrasi persuratan & data kependudukan Desa Marga Agung berbasis
PowerShell + Python (CSV sebagai basis data).

## Struktur

| Folder | Isi |
|---|---|
| `Data-desa/` | Script inti: pencarian warga, koreksi data, dashboard admin, notifikasi WA |
| `Antrean-surat/` | Antrean pengajuan surat (draft, status, pengiriman) |
| `Template-surat/` | 12 template `.docx` surat keterangan (kosong, siap isi) |
| `PANDUAN-SERAH-TERIMA.md` | Panduan operasional & serah terima ke petugas desa |

## ⚠️ Data warga TIDAK disimpan di repository ini

File berikut **sengaja di-`.gitignore`** karena memuat data pribadi warga
(NIK, KK, alamat, nomor HP) yang dilindungi UU PDP:

- `Data-desa/database_kk.csv`
- `Data-desa/admin_wa.csv`
- `Data-desa/antrean_koreksi.csv`
- `Antrean-surat/*.csv`, `*.docx`
- `Data-desa/_arsip/sumber-excel/` (Excel PRODESKEL)
- `ratings.csv`, `*.sql`

Data asli tetap tersimpan **lokal** di komputer desa / backup offline.
Jika clone repo ini di mesin baru, salin file data secara manual lewat media
terenkripsi — jangan lewat Git.

## Menjalankan

```powershell
cd Data-desa
.\dashboard_admin.ps1
```

Lihat `PANDUAN-SERAH-TERIMA.md` untuk alur lengkap.
# uji otomatis Rab, 04 Nov 2026 11.00.58
