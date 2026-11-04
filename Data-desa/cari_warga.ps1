# ============================================================================
# cari_warga.ps1 - Pencarian data warga untuk layanan surat Desa Marga Agung
#
# ALUR RESMI (wajib): nama lengkap + NIK 16 digit diminta sekaligus.
# Script ini HANYA MEMBACA. Tidak pernah mengubah database warga.
#
# Pemakaian:
#   .\cari_warga.ps1 -nik "1801130507820001" -nama "WARGA CONTOH"
#   .\cari_warga.ps1 -nik "1801130507820001" -nama "WARGA CONTOH" -tahun 1982
#   .\cari_warga.ps1 -nama "WARGA CONTOH" -tahun 1982
#
# STATUS yang dikeluarkan (baris pertama):
#   NAMA_COCOK            -> Kasus A: NIK ketemu & nama cocok. Data boleh dipakai.
#   NAMA_TIDAK_COCOK      -> Kasus B: NIK ketemu tapi nama beda.  STOP.
#   NIK_TIDAK_ADA         -> Kasus C awal: minta konfirmasi NIK + tahun lahir.
#   COCOK_SATU            -> Kasus C: nama+tahun cocok TEPAT SATU. Pakai NIK ketikan.
#   KEMBAR                -> nama+tahun cocok >1 data. STOP (tanpa menampilkan siapa pun).
#   TIDAK_COCOK           -> nama+tahun tidak cocok. STOP.
#   NIK_DB_BUKAN_16       -> Kasus D: NIK di database bukan 16 digit. STOP + catat koreksi.
#   NIK_BUKAN_16          -> NIK ketikan bukan 16 digit.
#   NAMA_KURANG           -> nama belum dikirim (alur wajib nama + NIK).
#   TAHUN_KURANG          -> pencarian nama butuh tahun lahir.
#
# WAJIB_CATAT_KOREKSI: YA|<jenis>  -> bot wajib memanggil catat_koreksi.ps1
# ============================================================================

param(
    [string]$nik,
    [string]$nama,
    [int]$tahun = 0,
    [string]$DataFile = ""
)

$ErrorActionPreference = "Stop"

if (-not $DataFile) { $DataFile = Join-Path $PSScriptRoot "database_kk.csv" }
if (-not (Test-Path $DataFile)) {
    Write-Output "STATUS: ERROR"
    Write-Output "PESAN: file database tidak ditemukan"
    exit 1
}

$semua = @(Import-Csv -Path $DataFile -Encoding UTF8)
if ($semua.Count -eq 0) {
    Write-Output "STATUS: ERROR"
    Write-Output "PESAN: database kosong"
    exit 1
}

function Angka($v) { if ($null -eq $v) { return "" } return ("$v" -replace '\D', '') }

function NormNama($v) {
    if ($null -eq $v) { return "" }
    $s = "$v".ToUpper()
    $s = $s -replace '[^A-Z0-9 ]', ' '
    $s = $s -replace '\s+', ' '
    return $s.Trim()
}

function TahunDari($r) {
    $t = "$($r.'Tahun Lahir')".Trim()
    if ($t) { return $t }
    $m = [regex]::Match("$($r.'Tanggal Lahir')", '(\d{4})')
    if ($m.Success) { return $m.Groups[1].Value }
    return ""
}

function TampilData($r, $nikTerpakai) {
    Write-Output ("NIK: " + (Angka $r.NIK))
    Write-Output ("NIK_TERPAKAI: " + $nikTerpakai)
    Write-Output ("Nama: " + "$($r.'Nama Anggota Keluarga')".Trim())
    Write-Output ("Tempat/Tgl Lahir: " + "$($r.'Tempat Lahir')".Trim() + ", " + "$($r.'Tanggal Lahir')".Trim())
    Write-Output ("Tahun Lahir: " + (TahunDari $r))
    Write-Output ("Jenis Kelamin: " + "$($r.'Jenis Kelamin')".Trim())
    Write-Output ("Agama: " + "$($r.Agama)".Trim())
    Write-Output ("Status: " + "$($r.Status)".Trim())
    Write-Output ("Pekerjaan: " + "$($r.Pekerjaan)".Trim())
    Write-Output ("Pendidikan: " + "$($r.Pendidikan)".Trim())
    Write-Output ("Alamat: " + "$($r.Alamat)".Trim())
    Write-Output ("RT/Dusun: " + "$($r.RT)".Trim() + " / " + "$($r.Dusun)".Trim())
    Write-Output ("Kepala Keluarga: " + "$($r.'Nama Kepala Keluarga')".Trim())
    Write-Output ("Hubungan: " + "$($r.Hubungan)".Trim())
}

function CariNamaTahun($kunci, $thn) {
    return @($semua | Where-Object {
        (NormNama $_.'Nama Anggota Keluarga') -eq $kunci -and (TahunDari $_) -eq "$thn"
    })
}

# ---------------------------------------------------------------- validasi awal
if (-not $nik -and -not $nama) {
    Write-Output "STATUS: NAMA_KURANG"
    Write-Output "PESAN: minta nama lengkap dan NIK 16 digit dulu"
    exit
}

if (-not $nama) {
    Write-Output "STATUS: NAMA_KURANG"
    Write-Output "PESAN: alur wajib minta NAMA LENGKAP dan NIK sekaligus"
    if ($nik) { Write-Output ("NIK_DITERIMA: " + (Angka $nik)) }
    exit
}

$kunciNama = NormNama $nama
$bersih = Angka $nik

if ($nik -and $bersih.Length -ne 16) {
    Write-Output "STATUS: NIK_BUKAN_16"
    Write-Output ("JUMLAH_DIGIT: " + $bersih.Length)
    Write-Output "PESAN: NIK harus tepat 16 digit angka"
    exit
}

# ------------------------------------------------------- 1) pencarian lewat NIK
if ($nik) {
    $row = $semua | Where-Object { (Angka $_.NIK) -eq $bersih } | Select-Object -First 1

    if ($row) {
        $nikDb = Angka $row.NIK
        if ($nikDb.Length -ne 16) {
            Write-Output "STATUS: NIK_DB_BUKAN_16"
            Write-Output ("NIK_KETIKAN: " + $bersih)
            Write-Output ("NIK_DATABASE: " + $nikDb)
            Write-Output ("JUMLAH_DIGIT_DATABASE: " + $nikDb.Length)
            Write-Output "WAJIB_CATAT_KOREKSI: YA|NIK_DB_BUKAN_16"
            Write-Output "PESAN: NIK di database desa bukan 16 digit. JANGAN terbitkan surat. Catat ke antrean koreksi, arahkan ke admin desa."
            TampilData $row $nikDb
            exit
        }

        $namaDb = NormNama $row.'Nama Anggota Keluarga'
        if ($namaDb -eq $kunciNama) {
            Write-Output "STATUS: NAMA_COCOK"
            Write-Output "PESAN: NIK dan nama cocok dengan data desa. Data boleh dipakai."
            TampilData $row $nikDb
            exit
        }

        # Kasus B - nama tidak cocok: jangan tampilkan data pribadi, STOP.
        Write-Output "STATUS: NAMA_TIDAK_COCOK"
        Write-Output ("NAMA_DATABASE: " + "$($row.'Nama Anggota Keluarga')".Trim())
        Write-Output ("DUSUN_DATABASE: " + "$($row.Dusun)".Trim())
        Write-Output "PESAN: NIK ini terdaftar atas nama lain. JANGAN terbitkan surat. Minta warga cek ulang NIK dan nama, tawarkan pencarian pakai nama + tahun lahir."
        exit
    }

    # NIK belum ketemu -> lanjut jalur nama + tahun lahir (Kasus C)
    if ($tahun -le 0) {
        Write-Output "STATUS: NIK_TIDAK_ADA"
        Write-Output "PESAN: NIK tidak ada di data desa. Minta warga konfirmasi ulang NIK (ketik sekali lagi), lalu tanyakan tahun lahir."
        exit
    }
}

# --------------------------------------- 2) pencarian nama + tahun lahir (Kasus C)
if ($tahun -le 0) {
    Write-Output "STATUS: TAHUN_KURANG"
    Write-Output "PESAN: pencarian pakai nama membutuhkan tahun lahir warga"
    exit
}

$cocok = CariNamaTahun $kunciNama $tahun

if ($cocok.Count -eq 0) {
    Write-Output "STATUS: TIDAK_COCOK"
    Write-Output ("NAMA_DICARI: " + $kunciNama)
    Write-Output ("TAHUN_DICARI: " + $tahun)
    Write-Output "PESAN: nama dan tahun lahir tidak cocok dengan data desa. JANGAN terbitkan surat, arahkan ke admin desa."
    exit
}

if ($cocok.Count -gt 1) {
    Write-Output "STATUS: KEMBAR"
    Write-Output ("JUMLAH_COCOK: " + $cocok.Count)
    Write-Output "PESAN: nama dan tahun lahir cocok dengan lebih dari satu data. JANGAN terbitkan surat dan JANGAN tampilkan siapa pun. Arahkan ke admin desa."
    exit
}

$row = $cocok[0]
$nikDb = Angka $row.NIK

# Kasus D lewat jalur nama
if ($nikDb.Length -ne 16) {
    Write-Output "STATUS: NIK_DB_BUKAN_16"
    if ($bersih) { Write-Output ("NIK_KETIKAN: " + $bersih) }
    Write-Output ("NIK_DATABASE: " + $nikDb)
    Write-Output ("JUMLAH_DIGIT_DATABASE: " + $nikDb.Length)
    Write-Output "WAJIB_CATAT_KOREKSI: YA|NIK_DB_BUKAN_16"
    Write-Output "PESAN: nama dan tahun lahir cocok, TAPI NIK di database bukan 16 digit. JANGAN terbitkan surat. Catat ke antrean koreksi, arahkan ke admin desa."
    TampilData $row $nikDb
    exit
}

# Kasus C sah: nama + tahun cocok tepat satu, NIK ketikan 16 digit
Write-Output "STATUS: COCOK_SATU"
Write-Output ("NIK_DATABASE: " + $nikDb)
if ($bersih -and $bersih -ne $nikDb) {
    Write-Output "WAJIB_CATAT_KOREKSI: YA|NIK_BEDA"
    Write-Output "PESAN: nama dan tahun lahir cocok tepat satu data, NIK ketikan dipakai. NIK ketikan BERBEDA dengan database - wajib dicatat ke antrean koreksi dan ditandai di konfirmasi admin."
    TampilData $row $bersih
} else {
    Write-Output "PESAN: nama dan tahun lahir cocok tepat satu data."
    TampilData $row $nikDb
}
exit
