# ============================================================================
# cek_admin.ps1 - Periksa apakah sebuah nomor WhatsApp terdaftar sebagai ADMIN.
#
# Hanya MEMBACA Data-desa\admin_wa.csv. Tidak pernah mengubah apa pun.
# Menormalkan format nomor: 0856... / +62 856... / 62856... / 856... -> 62856...
#
# Pemakaian:
#   .\cek_admin.ps1 -Nomor "085658361716"
#   .\cek_admin.ps1 -Nomor "+62 856-5836-1716" -Ringkas
#
# Keluaran (baris pertama selalu STATUS):
#   STATUS: ADMIN          -> nomor terdaftar & aktif
#   STATUS: ADMIN_NONAKTIF -> terdaftar tapi aktif != 1
#   STATUS: WARGA          -> nomor valid tapi tidak terdaftar
#   STATUS: NOMOR_TIDAK_VALID -> bukan nomor HP Indonesia yang wajar
#
# Exit code: 0 ADMIN, 1 WARGA, 2 NOMOR_TIDAK_VALID, 3 ADMIN_NONAKTIF
# ============================================================================

param(
    [Parameter(Mandatory = $true)][string]$Nomor,
    [string]$DaftarAdmin = "",
    [switch]$Ringkas
)

$ErrorActionPreference = "Stop"

if (-not $DaftarAdmin) { $DaftarAdmin = Join-Path $PSScriptRoot "admin_wa.csv" }

function Normalisasi-Nomor {
    param([string]$n)
    if (-not $n) { return "" }
    $d = ($n -replace '[^0-9]', '')          # buang +, spasi, strip, kurung
    if ($d -eq "") { return "" }
    if ($d.StartsWith("00")) { $d = $d.Substring(2) }
    if ($d.StartsWith("62")) { return $d }
    if ($d.StartsWith("0")) { return "62" + $d.Substring(1) }
    if ($d.StartsWith("8")) { return "62" + $d }
    return $d
}

$bersih = Normalisasi-Nomor $Nomor

function Selesai {
    param([string]$Status, [string]$Nama = "", [string]$Jabatan = "", [int]$Kode)
    if ($Ringkas) {
        Write-Output $Status
    }
    else {
        Write-Output ("STATUS: " + $Status)
        Write-Output ("NOMOR: " + $bersih)
        if ($Nama) { Write-Output ("NAMA: " + $Nama) }
        if ($Jabatan) { Write-Output ("JABATAN: " + $Jabatan) }
    }
    exit $Kode
}

# --- validasi bentuk nomor HP Indonesia (62 8xx, 10-15 digit) --------------
if ($bersih -notmatch '^628[0-9]{7,12}$') {
    Selesai -Status "NOMOR_TIDAK_VALID" -Kode 2
}

if (-not (Test-Path $DaftarAdmin)) {
    Selesai -Status "DAFTAR_ADMIN_TIDAK_ADA" -Kode 2
}

$baris = @(Import-Csv -Path $DaftarAdmin -Encoding UTF8)
$cocok = $baris | Where-Object { (Normalisasi-Nomor "$($_.nomor)") -eq $bersih } | Select-Object -First 1

if (-not $cocok) {
    Selesai -Status "WARGA" -Kode 1
}

$nama = "$($cocok.nama)".Trim()
$jabatan = "$($cocok.jabatan)".Trim()
$aktif = "$($cocok.aktif)".Trim()

if ($aktif -ne "1") {
    Selesai -Status "ADMIN_NONAKTIF" -Nama $nama -Jabatan $jabatan -Kode 3
}

Selesai -Status "ADMIN" -Nama $nama -Jabatan $jabatan -Kode 0
