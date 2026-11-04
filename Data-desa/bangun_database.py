# -*- coding: utf-8 -*-
"""
bangun_database.py - Membangun database_kk.csv dari 6 FILE RESMI Dusun I-VI.

Aturan:
  * HANYA file "DAFTAR KK_PRODESKEL BARU Dusun I..VI.xls" (file 'acak2' TIDAK dipakai).
  * Kolom "Tahun Lahir" DITURUNKAN dari "Tanggal Lahir" (tidak ada kolom tahun di Excel).
  * Pemetaan header eksplisit; kalau kolom kritis hilang -> GAGAL KERAS (exit 1).
  * Setiap kolom yang dipetakan dicatat di laporan_header.txt untuk audit.

Jalankan: bangun_database.bat  /  py bangun_database.py
"""
import csv, glob, os, re, sys
from collections import Counter
import xlrd

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "database_kk.csv")
LAPORAN = os.path.join(BASE, "laporan_bangun_database.txt")

# (kolom baku CSV, wajib?)
COLS = ["NIK", "Nama Anggota Keluarga", "Tempat Lahir", "Tanggal Lahir",
        "Tahun Lahir", "Jenis Kelamin", "Hubungan", "Status", "Agama",
        "Pendidikan", "Pekerjaan", "Alamat", "RT", "RW", "Dusun",
        "Nama Kepala Keluarga", "Kode Keluarga", "Gol Darah",
        "Kewarganegaraan", "Etnis", "Sumber"]

# urutan baku kolom Excel setelah "Usia" (dipakai kalau header rusak/ganda)
TAIL = ["Usia", "Status", "Agama", "GDarah", "Kewarganegaraan",
        "Etnis / Suku", "Pendidikan", "Pekerjaan"]

ALIAS = {
    "rw": "rw", "rt": "rt", "dusun": "dusun", "alamat": "alamat",
    "kodekeluarga": "kode", "nomorkk": "kode", "kk": "kode",
    "namakepalakeluarga": "namakk", "kepalakeluarga": "namakk",
    "no": "no", "nomor": "no", "nik": "nik",
    "namaanggota": "nama", "namaanggotakeluarga": "nama", "nama": "nama",
    "jeniskelamin": "jk", "jk": "jk", "lp": "jk",
    "hubungan": "hubungan", "statusdalamkeluarga": "hubungan",
    "tempatlahir": "tempat", "tgllahir": "tgl", "tanggallahir": "tgl",
    "usia": "usia", "umur": "usia",
    "status": "status", "statusperkawinan": "status", "statuskawin": "status",
    "agama": "agama",
    "gd": "gd", "gdarah": "gd", "golongandarah": "gd", "goldarah": "gd",
    "kewarganegaraan": "wni", "wn": "wni",
    "etnis": "etnis", "etnissuku": "etnis", "suku": "etnis",
    "pendidikan": "pendidikan", "pekerjaan": "pekerjaan",
}

FIELDS_21 = ["rw", "rt", "dusun", "alamat", "kode", "namakk", "no", "nik",
             "nama", "jk", "hubungan", "tempat", "tgl", "usia", "status",
             "agama", "gd", "wni", "etnis", "pendidikan", "pekerjaan"]
SKEMA_B = ["rt", "kode", "namakk", "nik", "nama", "jk", "hubungan",
           "tempat", "tgl", "usia", "status", "pendidikan", "pekerjaan"]

log = []
karantina = []
perbaikan = []


def catat(s):
    log.append(s)


def key(h):
    return re.sub(r"[^a-z0-9]", "", str(h).lower())


def norm_hdr(h):
    return ALIAS.get(key(h))


def digits(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        v = ("%.0f" % v) if abs(v - round(v)) < 1e-6 else repr(v)
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return re.sub(r"\D", "", s)


def teks(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if abs(v - round(v)) < 1e-6:
            return str(int(round(v)))
        return repr(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def tanggal(v, datemode=0):
    """-> (dd/mm/yyyy, tahun, catatan). tahun "" = dikarantina (ambigu)."""
    if v is None:
        return "", ""
    if isinstance(v, float) and v == v:
        if v <= 0:
            return "", "", "tanggal kosong"
        try:
            d = xlrd.xldate.xldate_as_datetime(v, datemode)
            return d.strftime("%d/%m/%Y"), str(d.year), ""
        except Exception:
            return "", "", "serial Excel tidak valid"
    s = teks(v)
    if not s:
        return "", "", "tanggal kosong"
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$", s)
    if m:
        return "%02d/%02d/%s" % (int(m.group(1)), int(m.group(2)), m.group(3)), m.group(3), ""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return "%s/%s/%s" % (m.group(3), m.group(2), m.group(1)), m.group(1), ""
    # tak-ambigu: garis miring kedua hilang, mis. 20/051965
    m = re.match(r"^(\d{1,2})[/\-.](\d{2})(\d{4})$", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        if 1 <= d <= 31 and 1 <= mo <= 12:
            return ("%02d/%02d/%s" % (d, mo, y)), y, "diperbaiki otomatis (pemisah tanggal hilang)"
    # ambigu: tahun 2 digit / tahun tidak wajar -> karantina, jangan ditebak
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2})$", s)
    if m:
        return "", "", "tahun 2 digit (%s) - perlu dicek" % s
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{1,3})$", s)
    if m:
        return "", "", "tahun tidak wajar (%s) - perlu dicek" % s
    return "", "", "format tidak dikenal (%s)" % s


def jk(v):
    s = teks(v).upper()
    if s in ("L", "LAKI-LAKI", "LAKI LAKI", "LAKILAKI", "M"):
        return "LAKI-LAKI"
    if s in ("P", "PEREMPUAN", "W", "F"):
        return "PEREMPUAN"
    return s


ROMAWI = {"1": "I", "2": "II", "3": "III", "4": "IV", "5": "V", "6": "VI"}


def dusun_rapi(nilai, hint):
    for s in (teks(nilai).upper(), teks(hint).upper()):
        m = re.search(r"DUSUN\s*(VI|IV|V|III|II|I|\d+)", s)
        if m:
            d = m.group(1)
            return "DUSUN " + ROMAWI.get(d, d)
    return teks(nilai).upper()


def rapi(v):
    s = teks(v)
    if s in ("0", "0.0", "-", "~", "#N/A"):
        return ""
    return s


def peta_kolom(header_row):
    names = [norm_hdr(h) for h in header_row]
    while names and names[-1] is None:
        names.pop()
    if not names or "nik" not in names:
        return None, False
    diperbaiki = False
    if "usia" in names:
        i0 = names.index("usia")
        ekor = names[i0:]
        if len(ekor) != len(set(ekor)):
            diperbaiki = True
            for i, f in enumerate(TAIL):
                pos = i0 + i
                if pos < len(names):
                    names[pos] = norm_hdr(f)
    peta, sudah = {}, set()
    for i, f in enumerate(names):
        if f and f not in sudah:
            peta[i] = f
            sudah.add(f)
    return peta, diperbaiki


def proses_sheet(sh, fname, gagal):
    """Kembalikan (records, status_sheet)."""
    nrows, ncols = sh.nrows, sh.ncols
    if nrows < 2 or ncols < 5:
        return [], "kosong"
    peta, mulai, ket = None, None, ""
    for r in range(min(12, nrows)):
        p, diperbaiki = peta_kolom(sh.row_values(r))
        if p:
            peta, mulai = p, r + 1
            ket = "header baris %d" % r + (" (kolom bergeser diperbaiki)" if diperbaiki else "")
            break
    if peta is None:
        v = sh.row_values(0)
        if ncols >= 20 and len(digits(v[6])) >= 12:
            peta = {i: f for i, f in enumerate(FIELDS_21)}
            mulai, ket = 0, "TANPA header - susunan ditebak dari 21 kolom"
        elif ncols == 13 and len(digits(v[3])) >= 12:
            peta = {i: f for i, f in enumerate(SKEMA_B)}
            mulai, ket = 0, "TANPA header - susunan ditebak dari 13 kolom"
        else:
            return [], "dilewati (tidak ada header & bentuk kolom tidak dikenal)"

    # sheet hanya dipakai kalau punya NIK + nama anggota + tanggal lahir
    if "nik" not in peta.values():
        return [], "dilewati (tanpa kolom NIK)"
    if "nama" not in peta.values():
        return [], "dilewati (sheet pendukung: tanpa kolom Nama Anggota Keluarga)"
    if "tgl" not in peta.values():
        gagal.append("%s / %s : kolom Tanggal Lahir tidak ada (tahun lahir wajib)" % (fname, sh.name))
        return [], "GAGAL (tanpa Tanggal Lahir)"

    hasil = []
    for r in range(mulai, nrows):
        vals = sh.row_values(r)
        d = {f: vals[i] for i, f in peta.items() if i < len(vals)}
        nik = digits(d.get("nik"))
        nama = rapi(d.get("nama"))
        if len(nik) < 10 or not nama:
            continue
        tgl, tahun, catatan_tgl = tanggal(d.get("tgl"), sh.book.datemode)
        if not tahun:
            karantina.append({
                "NIK": nik, "Nama": nama.upper(), "Dusun": dusun_rapi(d.get("dusun"), fname),
                "Tanggal_Lahir_di_Excel": teks(d.get("tgl")), "Alasan": catatan_tgl,
                "Sumber": "%s / %s" % (fname, sh.name),
            })
        elif catatan_tgl:
            perbaikan.append({"NIK": nik, "Nama": nama.upper(), "Nilai": teks(d.get("tgl")),
                              "Hasil": tgl, "Catatan": catatan_tgl,
                              "Sumber": "%s / %s" % (fname, sh.name)})
        rec = {
            "NIK": nik,
            "Nama Anggota Keluarga": nama.upper(),
            "Tempat Lahir": rapi(d.get("tempat")).upper(),
            "Tanggal Lahir": tgl,
            "Tahun Lahir": tahun,
            "Jenis Kelamin": jk(d.get("jk")),
            "Hubungan": rapi(d.get("hubungan")),
            "Status": rapi(d.get("status")),
            "Agama": rapi(d.get("agama")),
            "Pendidikan": rapi(d.get("pendidikan")),
            "Pekerjaan": rapi(d.get("pekerjaan")),
            "Alamat": rapi(d.get("alamat")).upper(),
            "RT": rapi(d.get("rt")),
            "RW": rapi(d.get("rw")),
            "Dusun": dusun_rapi(d.get("dusun"), fname),
            "Nama Kepala Keluarga": rapi(d.get("namakk")).upper(),
            "Kode Keluarga": re.sub(r"\D", "", teks(d.get("kode"))),
            "Gol Darah": rapi(d.get("gd")),
            "Kewarganegaraan": rapi(d.get("wni")),
            "Etnis": rapi(d.get("etnis")),
            "Sumber": "%s / %s" % (fname, sh.name),
        }
        for k in ("RT", "RW"):
            if re.match(r"^\d+\.0$", rec[k] or ""):
                rec[k] = rec[k][:-2]
        hasil.append(rec)
    return hasil, ket


def main():
    # Cari file .xls di folder script DAN di _arsip/sumber-excel (lokasi arsip).
    _pola = "DAFTAR KK_PRODESKEL BARU Dusun*.xls"
    semua_file = []
    for _dir in (BASE, os.path.join(BASE, "_arsip", "sumber-excel")):
        semua_file += glob.glob(os.path.join(_dir, _pola))
    semua_file = sorted(semua_file)
    resmi = [f for f in semua_file if "acak" not in os.path.basename(f).lower()]
    dibuang = [os.path.basename(f) for f in semua_file if f not in resmi]

    catat("FILE RESMI DIPAKAI (%d):" % len(resmi))
    for f in resmi:
        catat("  - %s" % os.path.basename(f))
    catat("FILE TIDAK DIPAKAI: %s" % (", ".join(dibuang) if dibuang else "-"))
    catat("")

    if len(resmi) != 6:
        catat("GAGAL: file resmi harus 6 (Dusun I-VI), ditemukan %d." % len(resmi))
        simpan_log()
        print(open(LAPORAN, encoding="utf-8").read())
        return 1

    semua, gagal, stat = [], [], []
    for f in resmi:
        fname = os.path.basename(f)
        wb = xlrd.open_workbook(f)
        n = 0
        catat("FILE: %s" % fname)
        for sh in wb.sheets():
            rows, ket = proses_sheet(sh, fname, gagal)
            n += len(rows)
            semua.extend(rows)
            catat("  %-28s %-58s %5d baris" % (sh.name[:28], ket[:58], len(rows)))
        stat.append((fname, n))
        catat("  TOTAL %d baris\n" % n)
        if n == 0:
            gagal.append("%s : tidak ada satu pun baris terbaca" % fname)

    def skor(r):
        return sum(1 for c in COLS if r.get(c))

    terbaik = {}
    konflik = []
    for r in semua:
        k = r["NIK"]
        if k not in terbaik:
            terbaik[k] = r
        else:
            lama = terbaik[k]
            if lama["Nama Anggota Keluarga"] != r["Nama Anggota Keluarga"]:
                konflik.append((k, lama["Nama Anggota Keluarga"], r["Nama Anggota Keluarga"]))
            if skor(r) > skor(lama):
                terbaik[k] = r
    data = sorted(terbaik.values(), key=lambda r: (r["Dusun"], r["Nama Anggota Keluarga"]))

    # ---- validasi wajib ----
    tanpa_tahun = [r for r in data if not r["Tahun Lahir"]]
    tanpa_dusun = [r for r in data if not r["Dusun"]]
    if not data:
        gagal.append("database kosong")

    sementara = OUT + ".tmp"
    with open(sementara, "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=COLS, quoting=csv.QUOTE_MINIMAL)
        w.writeheader()
        for r in data:
            w.writerow({c: r.get(c, "") for c in COLS})

    catat("")
    catat("HASIL")
    catat("  baris terbaca        : %d" % len(semua))
    catat("  warga unik (NIK)     : %d" % len(data))
    catat("  duplikat dibuang     : %d" % (len(semua) - len(data)))
    catat("  tanpa Tahun Lahir    : %d" % len(tanpa_tahun))
    catat("  tanpa Dusun          : %d" % len(tanpa_dusun))
    catat("  NIK tidak 16 digit   : %d" % len([r for r in data if len(r["NIK"]) != 16]))
    catat("  NIK bentrok nama     : %d (nama beda untuk NIK sama)" % len(konflik))
    catat("  tanggal dikarantina  : %d (tahun lahir tidak bisa dipastikan)" % len(tanpa_tahun))
    catat("  tanggal diperbaiki   : %d (pemisah hilang, otomatis)" % len(perbaikan))
    pjg = Counter(len(r["NIK"]) for r in data)
    catat("  panjang NIK          : %s" % dict(sorted(pjg.items())))
    catat("")
    if gagal:
        catat("VALIDASI: GAGAL (%d masalah) - database TIDAK diperbarui" % len(gagal))
        for g in gagal[:20]:
            catat("  ! %s" % g)
        simpan_log()
        print(open(LAPORAN, encoding="utf-8").read())
        return 1

    tulis_csv("tanggal_perlu_dicek.csv", karantina,
              ["NIK", "Nama", "Dusun", "Tanggal_Lahir_di_Excel", "Alasan", "Sumber"])
    tulis_csv("tanggal_diperbaiki_otomatis.csv", perbaikan,
              ["NIK", "Nama", "Nilai", "Hasil", "Catatan", "Sumber"])
    if konflik:
        catat("  contoh NIK bentrok nama:")
        for k, a, b in konflik[:5]:
            catat("    %s : %r vs %r" % (k, a, b))

    try:
        os.replace(sementara, OUT)
    except PermissionError:
        catat("GAGAL: database_kk.csv sedang dipakai program lain (Excel/OneDrive). Tutup lalu ulangi.")
        simpan_log()
        print(open(LAPORAN, encoding="utf-8").read())
        return 1
    catat("VALIDASI: LULUS - database_kk.csv diperbarui")
    catat("  %s" % OUT)
    simpan_log()
    print(open(LAPORAN, encoding="utf-8").read())
    return 0


def tulis_csv(nama, rows, kolom):
    p = os.path.join(BASE, nama)
    try:
        with open(p, "w", encoding="utf-8-sig", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=kolom, extrasaction="ignore")
            w.writeheader()
            for r in rows:
                w.writerow(r)
        catat("  file laporan        : %s (%d baris)" % (nama, len(rows)))
    except PermissionError:
        catat("  file laporan        : %s TERKUNCI (buka di Excel?) - dilewati" % nama)


def simpan_log():
    txt = "\n".join(log)
    with open(LAPORAN, "w", encoding="utf-8") as fh:
        fh.write(txt + "\n")


if __name__ == "__main__":
    sys.exit(main())
