# ============================================================================
# kabari_warga.ps1 - Kirim pesan WhatsApp ke warga bahwa suratnya siap diambil.
#
# Membaca data dari Antrean-surat (nama, kontak, jenis surat) lalu mengirim
# pesan lewat kanal WhatsApp desa yang sudah tersambung. Setelah berhasil,
# mencatat waktu pengabaran di antrean_ambil.csv.
#
# Pakai:
#   .\kabari_warga.ps1 -Id "SR-20261027-001" -DryRun      # hanya lihat pesan
#   .\kabari_warga.ps1 -Id "SR-20261027-001"
#   .\kabari_warga.ps1 -Id "SR-..." -Tanggal "28 Oktober 2026" -Tambahan "Bawa fotokopi KK"
#
# Keluaran JSON: { ok, id, nama, nomor, pesan, terkirim }
# ============================================================================

param(
    [Parameter(Mandatory = $true)][string]$Id,
    [string]$Tanggal = "",
    [string]$Tambahan = "",
    [string]$AntreanDir = "",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

if (-not $AntreanDir) {
    $AntreanDir = Join-Path (Split-Path -Parent $PSScriptRoot) "Antrean-surat"
}
$csv = Join-Path $AntreanDir "antrean_surat.csv"
$csvAmbil = Join-Path $AntreanDir "antrean_ambil.csv"

function Keluar-Json($obj) {
    $obj | ConvertTo-Json -Depth 4 -Compress
    exit 0
}

function Normalisasi-Nomor($n) {
    $d = ($n -replace '[^0-9]', '')
    if (-not $d) { return "" }
    if ($d.StartsWith("00")) { $d = $d.Substring(2) }
    if ($d.StartsWith("62")) { return $d }
    if ($d.StartsWith("0")) { return "62" + $d.Substring(1) }
    if ($d.StartsWith("8")) { return "62" + $d }
    return $d
}

$idBersih = ($Id -replace '\s', '').ToUpper()
if ($idBersih -notmatch '^SR-\d{8}-\d{3}$') {
    Keluar-Json @{ ok = $false; error = "FORMAT_ID_TIDAK_VALID"; id = $Id }
}
if (-not (Test-Path $csv)) {
    Keluar-Json @{ ok = $false; error = "DATA_ANTREAN_TIDAK_DITEMUKAN" }
}

$baris = @(Import-Csv -Path $csv -Encoding UTF8 |
    Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $idBersih }) | Select-Object -First 1
if (-not $baris) {
    Keluar-Json @{ ok = $false; error = "ID_TIDAK_ADA_DI_ANTREAN"; id = $idBersih }
}

$nomor = Normalisasi-Nomor "$($baris.kontak)"
if ($nomor -notmatch '^628[0-9]{7,12}$') {
    Keluar-Json @{ ok = $false; error = "KONTAK_WARGA_TIDAK_VALID"; id = $idBersih; kontak = "$($baris.kontak)" }
}

# --- tanggal siap: dari parameter, atau dari antrean_ambil ------------------
if (-not $Tanggal.Trim() -and (Test-Path $csvAmbil)) {
    $a = @(Import-Csv -Path $csvAmbil -Encoding UTF8 |
        Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $idBersih }) | Select-Object -First 1
    if ($a) { $Tanggal = "$($a.tanggal_siap)".Trim() }
}

$nama = "$($baris.nama)".Trim()
$jenis = "$($baris.jenis_surat)".Trim()
$keperluan = "$($baris.keperluan)".Trim()

$pesan = @()
$pesan += "Halo $nama, surat Anda sudah selesai dan siap diambil."
$pesan += ""
$pesan += "Nomor surat: $idBersih"
$pesan += "Jenis: $jenis"
if ($keperluan) { $pesan += "Keperluan: $keperluan" }
if ($Tanggal.Trim()) { $pesan += "Bisa diambil mulai: $($Tanggal.Trim())" }
$pesan += "Tempat: Kantor Desa Marga Agung"
$pesan += "Mohon bawa KTP asli."
if ($Tambahan.Trim()) { $pesan += ""; $pesan += $Tambahan.Trim() }
$pesan += ""
$pesan += "Balas pesan ini kalau ada yang ingin ditanyakan."
$teks = ($pesan -join "`n")

if ($DryRun) {
    Keluar-Json @{ ok = $true; id = $idBersih; nama = $nama; nomor = $nomor; pesan = $teks; terkirim = $false; dry_run = $true }
}

# --- kirim lewat kanal WhatsApp desa --------------------------------------
$hermes = Join-Path $env:LOCALAPPDATA "hermes\bin\hermes.exe"
if (-not (Test-Path $hermes)) { $hermes = "hermes" }

$hasil = & $hermes --profile surat-desa send -t "whatsapp:$nomor" --json $teks 2>&1
$teksHasil = ($hasil | Out-String)
$sukses = $teksHasil -match '"success"\s*:\s*true'

if (-not $sukses) {
    Keluar-Json @{ ok = $false; error = "GAGAL_KIRIM"; id = $idBersih; nomor = $nomor; detail = $teksHasil.Trim() }
}

# --- catat waktu pengabaran ----------------------------------------------
if (Test-Path $csvAmbil) {
    $ambil = @(Import-Csv -Path $csvAmbil -Encoding UTF8)
    $ada = $ambil | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $idBersih } | Select-Object -First 1
    if ($ada) {
        $ada.dikabari_pada = (Get-Date).ToString("yyyy-MM-dd HH:mm")
        $ambil = @($ambil | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -ne $idBersih }) + @($ada)
        $ambil | Sort-Object id_surat | Export-Csv -Path $csvAmbil -NoTypeInformation -Encoding UTF8
    }
}

Keluar-Json @{ ok = $true; id = $idBersih; nama = $nama; nomor = $nomor; pesan = $teks; terkirim = $true }
