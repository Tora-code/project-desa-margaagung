# ============================================================================
# terapkan_koreksi.ps1 - MENERAPKAN koreksi NIK ke database warga.
#
# Gerbang keras (semua harus terpenuhi, kalau tidak -> TOLAK):
#   1. catatan antrean berstatus DISETUJUI (artinya admin sudah menyetujui)
#   2. switch -setuju dikirim
#   3. nama admin diisi
# Setiap penerapan:
#   - membuat backup bertanggal database_kk.csv ke folder _backup\
#   - mencatat perubahan ke log_perubahan_database.csv (apa, siapa, kapan)
#   - mengubah status antrean menjadi DITERAPKAN
#
#   .\terapkan_koreksi.ps1 -id "KR-20261027-001" -admin "Suharno" -setuju
# ============================================================================

param(
    [Parameter(Mandatory=$true)][string]$id,
    [Parameter(Mandatory=$true)][string]$admin,
    [switch]$setuju,
    [string]$QueueFile = "",
    [string]$DataFile = "",
    [string]$LogFile = "",
    [string]$BackupDir = ""
)

$ErrorActionPreference = "Stop"
if (-not $QueueFile) { $QueueFile = Join-Path $PSScriptRoot "antrean_koreksi.csv" }
if (-not $DataFile)  { $DataFile  = Join-Path $PSScriptRoot "database_kk.csv" }
if (-not $LogFile)   { $LogFile   = Join-Path $PSScriptRoot "log_perubahan_database.csv" }
if (-not $BackupDir) { $BackupDir = Join-Path $PSScriptRoot "_backup" }

if (-not $setuju) {
    Write-Output "DITOLAK: perlu switch -setuju (bukti admin menyetujui perubahan)."
    exit 1
}
if (-not "$admin".Trim()) {
    Write-Output "DITOLAK: nama admin wajib diisi."
    exit 1
}
if (-not (Test-Path $QueueFile)) { Write-Output "DITOLAK: antrean koreksi tidak ada."; exit 1 }
if (-not (Test-Path $DataFile))  { Write-Output "DITOLAK: database tidak ada."; exit 1 }

$rows = @(Import-Csv -Path $QueueFile -Encoding UTF8)
$row = $rows | Where-Object { "$($_.id_koreksi)" -eq $id } | Select-Object -First 1
if (-not $row) { Write-Output "DITOLAK: ID $id tidak ada di antrean."; exit 1 }

if ("$($row.status)".ToUpper() -ne "DISETUJUI") {
    Write-Output ("DITOLAK: status catatan " + $row.status + ". Admin harus menyetujui dulu (setujui_koreksi.ps1).")
    exit 1
}

$nikLama = ($row.nik_database -replace '\D', '')
$nikBaru = ($row.nik_ketikan -replace '\D', '')
if ($nikBaru.Length -ne 16) {
    Write-Output "DITOLAK: NIK baru bukan 16 digit, tidak boleh diterapkan."
    exit 1
}

# ---- backup bertanggal ----
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
$stamp = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
$backup = Join-Path $BackupDir ("database_kk_" + $stamp + ".csv")
Copy-Item -LiteralPath $DataFile -Destination $backup

# ---- terapkan ----
$data = @(Import-Csv -Path $DataFile -Encoding UTF8)
$sasaran = @($data | Where-Object {
    (($_.NIK -replace '\D', '') -eq $nikLama) -and
    ($row.nama -eq "" -or ("$($_.'Nama Anggota Keluarga')".Trim().ToUpper() -eq $row.nama))
})

if ($sasaran.Count -ne 1) {
    Write-Output ("DITOLAK: sasaran di database tidak tunggal (" + $sasaran.Count + " baris cocok). Tidak ada yang diubah.")
    exit 1
}

$sasaran[0].NIK = $nikBaru
$data | Export-Csv -Path $DataFile -NoTypeInformation -Encoding UTF8 -Delimiter ","

# ---- catat perubahan (apa, siapa, kapan) ----
$log = [pscustomobject]@{
    waktu            = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    id_koreksi       = $id
    nik_lama         = $nikLama
    nik_baru         = $nikBaru
    nama             = "$($sasaran[0].'Nama Anggota Keluarga')".Trim()
    dusun            = "$($sasaran[0].Dusun)".Trim()
    disetujui_oleh   = "$admin".Trim()
    waktu_setuju     = "$($row.waktu_persetujuan)"
    file_backup      = $backup
}
if (Test-Path $LogFile) {
    $log | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8 -Append -Delimiter ","
} else {
    $log | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8 -Delimiter ","
}

# ---- tandai antrean ----
$row.status = "DITERAPKAN"
$rows | Export-Csv -Path $QueueFile -NoTypeInformation -Encoding UTF8 -Delimiter ","

Write-Output ("DITERAPKAN: " + $id)
Write-Output ("NIK_LAMA: " + $nikLama)
Write-Output ("NIK_BARU: " + $nikBaru)
Write-Output ("NAMA: " + $log.nama)
Write-Output ("DISETUJUI_OLEH: " + $log.disetujui_oleh)
Write-Output ("WAKTU: " + $log.waktu)
Write-Output ("BACKUP: " + $backup)
Write-Output ("LOG: " + $LogFile)
