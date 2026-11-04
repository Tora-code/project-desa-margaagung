# Cara Pakai Auto-Push ke GitHub

Repo: https://github.com/Tora-code/project-desa-margaagung (PRIVATE)

## Menyalakan

Dobel-klik **`1-NYALAKAN-AUTOPUSH.bat`**. Selesai. Watcher jalan di
belakang layar dan setiap perubahan file otomatis di-commit + push
ke GitHub dalam ~25 detik.

Nyalakan sekali saja setiap habis menyalakan komputer.

## Mematikan

Dobel-klik **`2-MATIKAN-AUTOPUSH.bat`**.

## Melihat aktivitas

Buka file **`_watcher_push.log`** — mencatat setiap push dan setiap
pemblokiran.

## Pengaman data warga

Sebelum setiap commit, watcher memeriksa **isi** semua file yang diubah.
Kalau menemukan NIK 16 digit atau nomor HP Indonesia, push dibatalkan
dan alasannya dicatat di log. Jadi walaupun suatu saat data warga
disimpan dengan nama file baru (yang belum ada di `.gitignore`),
datanya tetap tidak akan lolos ke GitHub.

Baris yang memuat kata `WARGA CONTOH`, `dummy`, `contoh`, atau `UJI `
diabaikan supaya contoh di dokumentasi tidak memblokir push.

## Batasan yang perlu diketahui

- Nomor HP format penulisan bebas (spasi/strip) juga terdeteksi.
- Kalau pengaman memblokir padahal isinya bukan data warga asli,
  tambahkan kata penanda seperti `contoh` pada baris itu.
- Watcher hanya memantau folder project ini.
- Kalau internet mati, push gagal dan dicatat di log. Watcher akan
  mencoba lagi saat ada perubahan berikutnya.
