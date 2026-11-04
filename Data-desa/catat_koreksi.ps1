# ============================================================================
# catat_koreksi.ps1 - Menambah satu catatan ke ANTREAN KOREKSI (bukan database).
#
# TIDAK PERNAH mengubah database_kk.csv. Hanya menulis ke antrean_koreksi.csv.
# Dipakai bot setiap kali NIK ketikan berbeda dengan database, atau NIK database
# bukan 16 digit.
#
# Pemakaian:
#   .\catat_koreksi.ps1 -nikKetikan "1801130507820001" -nikDatabase "1801130507820001" `
#       -nama "ABDUL MUIN" -dusun "DUSUN I" -jenis NIK_BEDA -catatan "NIK ketikan 16 digit"
# ============================================================================

param(
    [Parameter(Mandatory=$true)][string]$nikKetikan,
    [string]$nikDatabase = "",
    [Parameter(Mandatory=$true)][string]$nama,
    [string]$namaDatabase = "",
    [string]$dusun = "",
    [Parameter(Mandatory=$true)][ValidateSet("NIK_BEDA", "NIK_DB_BUKAN_16", "NAMA_TIDAK_COCOK")][string]$jenis,
    [string]$catatan = "",
    [string]$QueueFile = "",
    [string]$Sumber = "whastapp-layanan-surat"
)

$ErrorActionPreference = "Stop"
if (-not $QueueFile) { $QueueFile = Join-Path $PSScriptRoot "antrean_koreksi.csv" }

$kolom = @("id_koreksi", "waktu", "jenis", "nik_ketikan", "nik_database",
           "nama", "nama_database", "dusun", "sumber", "status",
           "disetujui_oleh", "waktu_persetujuan", "catatan")

$waktu = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$prefix = "KR-" + (Get-Date).ToString("yyyyMMdd")

$lama = @()
if (Test-Path $QueueFile) {
    $lama = @(Import-Csv -Path $QueueFile -Encoding UTF8)
}

$urut = 1
$hariIni = @($lama | Where-Object { "$($_.id_koreksi)" -like "$prefix-*" })
if ($hariIni.Count -gt 0) {
    $urut = ($hariIni | ForEach-Object { [int]("$($_.id_koreksi)".Split("-")[-1]) } | Measure-Object -Maximum).Maximum + 1
}
$id = "$prefix-" + $urut.ToString("000")

$baris = [pscustomobject]@{
    id_koreksi         = $id
    waktu              = $waktu
    jenis              = $jenis
    nik_ketikan        = ($nikKetikan -replace '\D', '')
    nik_database       = ($nikDatabase -replace '\D', '')
    nama               = "$nama".Trim().ToUpper()
    nama_database      = "$namaDatabase".Trim().ToUpper()
    dusun              = "$dusun".Trim()
    sumber             = $Sumber
    status             = "MENUNGGU"
    disetujui_oleh     = ""
    waktu_persetujuan  = ""
    catatan            = "$catatan".Trim()
}

if ($lama.Count -eq 0) {
    $baris | Export-Csv -Path $QueueFile -NoTypeInformation -Encoding UTF8 -Delimiter ","
} else {
    $baris | Export-Csv -Path $QueueFile -NoTypeInformation -Encoding UTF8 -Append -Delimiter ","
}

Write-Output ("KOREKSI_DICATAT: " + $id)
Write-Output ("STATUS_ANTREAN: MENUNGGU")
Write-Output ("JENIS: " + $jenis)
Write-Output ("NIK_KETIKAN: " + $baris.nik_ketikan)
Write-Output ("NIK_DATABASE: " + $baris.nik_database)
Write-Output "PESAN: tercatat untuk admin. Database TIDAK diubah."
