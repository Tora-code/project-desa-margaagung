# -*- coding: utf-8 -*-
"""
watcher_push.py — Pantau folder Project Desa MargaAgung, auto-commit + push
ke GitHub setiap ada perubahan file.

PENGAMAN: sebelum commit, setiap file diperiksa isinya. Kalau terdeteksi
data pribadi warga (NIK 16 digit / nomor HP Indonesia), commit DIBATALKAN
dan ditulis ke log. Ini mencegah data warga lolos ke GitHub walaupun
namanya tidak ada di .gitignore.

Jalankan:  pythonw watcher_push.py      (tanpa jendela, untuk harian)
           python  watcher_push.py      (dengan jendela, untuk lihat log)
Hentikan : Ctrl+C, atau hapus/tutup prosesnya lewat Task Manager.
"""

import os
import re
import sys
import time
import subprocess
import datetime

# ---------------- Konfigurasi ----------------
REPO = r"C:\Users\Sulthonu Tora Aulia\OneDrive\Dokumen\Project Desa MargaAgung"
GIT = r"C:\Users\Sulthonu Tora Aulia\AppData\Local\hermes\tools\git-2.53.0+3-win32-x64\mingw64\bin\git.exe"
GH = r"C:\Program Files\GitHub CLI\gh.exe"
LOG = os.path.join(REPO, "_watcher_push.log")
DEBOUNCE = 20          # detik tenang sebelum push (biar tidak push tiap ketikan)
INTERVAL = 5           # detik antar pemeriksaan

# ---------------- Pola data pribadi ----------------
# NIK Indonesia = 16 digit. Nomor HP Indonesia = 62 / 08 + 9-13 digit.
POLA_NIK = re.compile(r"(?<!\d)(?!0{16})\d{16}(?!\d)")
POLA_HP = re.compile(r"(?<!\d)(?:628|628\d|08)\d{8,12}(?!\d)")

# File yang tidak pernah diperiksa (bukan teks)
EKSTENSI_LEWAT = {".docx", ".xlsx", ".xls", ".pdf", ".zip", ".png",
                  ".jpg", ".jpeg", ".gif", ".ico", ".exe", ".dll"}

# Baris berisi penanda ini diabaikan (contoh dummy di dokumentasi)
PENANDA_DUMMY = ("WARGA CONTOH", "dummy", "contoh", "UJI ", "xxxxx",
                 "1234567890", "1801999999990")


def tulis(pesan):
    waktu = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    baris = f"[{waktu}] {pesan}"
    print(baris, flush=True)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(baris + "\n")
    except OSError:
        pass


def jalankan(perintah, cwd=REPO):
    """Jalankan perintah, kembalikan (kode, stdout+stderr)."""
    try:
        p = subprocess.run(perintah, cwd=cwd, capture_output=True,
                           text=True, encoding="utf-8", errors="replace",
                           timeout=180)
        return p.returncode, (p.stdout or "") + (p.stderr or "")
    except subprocess.TimeoutExpired:
        return 1, "TIMEOUT"
    except OSError as e:
        return 1, f"GAGAL JALANKAN: {e}"


def snapshot():
    """Ambil keadaan folder: {path relatif: (mtime, ukuran)}"""
    hasil = {}
    for akar, folder, berkas in os.walk(REPO):
        # Lewati folder yang tidak perlu dipantau
        folder[:] = [f for f in folder
                     if f not in (".git", "node_modules", "__pycache__")]
        for nama in berkas:
            penuh = os.path.join(akar, nama)
            rel = os.path.relpath(penuh, REPO)
            try:
                st = os.stat(penuh)
                hasil[rel] = (st.st_mtime, st.st_size)
            except OSError:
                continue
    return hasil


def periksa_keamanan():
    """
    Periksa isi file yang sudah di-stage.
    Kembalikan daftar (file, alasan) jika ada data pribadi.
    """
    kode, keluaran = jalankan([GIT, "diff", "--cached", "--name-only", "-z"])
    if kode != 0:
        return [(None, f"tidak bisa baca daftar stage: {keluaran.strip()}")]

    berkas = [b for b in keluaran.split("\0") if b.strip()]
    temuan = []

    for rel in berkas:
        if os.path.splitext(rel)[1].lower() in EKSTENSI_LEWAT:
            continue
        penuh = os.path.join(REPO, rel)
        if not os.path.isfile(penuh):
            continue
        try:
            with open(penuh, "r", encoding="utf-8", errors="replace") as f:
                isi = f.read()
        except OSError:
            continue

        for nomor, baris in enumerate(isi.splitlines(), 1):
            # Lewati baris yang jelas-jelas contoh/dummy
            if any(p.lower() in baris.lower() for p in PENANDA_DUMMY):
                continue
            if POLA_NIK.search(baris):
                temuan.append((rel, f"baris {nomor}: NIK 16 digit"))
                break
            if POLA_HP.search(baris):
                temuan.append((rel, f"baris {nomor}: nomor HP"))
                break
    return temuan


def sinkron():
    """Commit + push kalau ada perubahan. Return True kalau push sukses."""
    kode, _ = jalankan([GIT, "add", "-A"])
    if kode != 0:
        tulis("GAGAL: git add")
        return False

    kode, keluaran = jalankan([GIT, "diff", "--cached", "--quiet"])
    if kode == 0:
        return False                      # tidak ada perubahan
    if kode not in (0, 1):
        tulis(f"GAGAL: git diff --cached ({keluaran.strip()})")
        return False

    # === PENGAMAN: periksa isi sebelum commit ===
    temuan = periksa_keamanan()
    if temuan:
        jalankan([GIT, "reset"])          # batalkan stage
        tulis("=" * 60)
        tulis("!! PUSH DIBATALKAN — TERDETEKSI DATA PRIBADI WARGA !!")
        for berkas, alasan in temuan:
            tulis(f"   {berkas}  ->  {alasan}")
        tulis("   Pindahkan data itu ke luar folder project, lalu ulangi.")
        tulis("=" * 60)
        return False

    waktu = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    kode, keluaran = jalankan([GIT, "commit", "-m", f"Auto-sync {waktu}"])
    if kode != 0:
        tulis(f"GAGAL: git commit ({keluaran.strip()[:300]})")
        return False

    kode, keluaran = jalankan([GIT, "push", "origin", "HEAD"])
    if kode != 0:
        tulis(f"GAGAL: git push ({keluaran.strip()[:300]})")
        return False

    kode, keluaran = jalankan([GIT, "rev-parse", "--short", "HEAD"])
    tulis(f"PUSH BERHASIL — commit {keluaran.strip()} ke GitHub")
    return True


def main():
    if not os.path.isdir(os.path.join(REPO, ".git")):
        tulis(f"BATAL: {REPO} bukan repo git.")
        return 1
    os.chdir(REPO)
    tulis(f"Watcher aktif — memantau {REPO}")

    sebelumnya = snapshot()
    waktu_berubah = None

    while True:
        time.sleep(INTERVAL)
        sekarang = snapshot()
        if sekarang != sebelumnya:
            sebelumnya = sekarang
            waktu_berubah = time.time()
            continue
        # Ada perubahan dan sudah tenang selama DEBOUNCE detik -> sinkron
        if waktu_berubah and (time.time() - waktu_berubah) >= DEBOUNCE:
            waktu_berubah = None
            try:
                sinkron()
            except Exception as e:                      # noqa: BLE001
                tulis(f"ERROR tak terduga: {e}")


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        tulis("Watcher dihentikan.")
