# ============================================================================
# setujui_koreksi.ps1 - Admin MENYETUJUI satu catatan koreksi.
#
# Menyetujui BELUM mengubah database. Perubahan hanya terjadi lewat
# terapkan_koreksi.ps1 (butuh status DISETUJUI + switch -setuju).
#
#   .\setujui_koreksi.ps1 -id "KR-20261027-001" -admin "Suharno"
# ============================================================================

param(
    [Parameter(Mandatory=$true)][string]$id,
    [Parameter(Mandatory=$true)][string]$admin,
    [string]$QueueFile = ""
)

$ErrorActionPreference = "Stop"
if (-not $QueueFile) { $QueueFile = Join-Path $PSScriptRoot "antrean_koreksi.csv" }
if (-not (Test-Path $QueueFile)) { Write-Output "GAGAL: antrean koreksi belum ada"; exit 1 }

$rows = @(Import-Csv -Path $QueueFile -Encoding UTF8)
$row = $rows | Where-Object { "$($_.id_koreksi)" -eq $id } | Select-Object -First 1
if (-not $row) { Write-Output "GAGAL: ID $id tidak ditemukan di antrean"; exit 1 }

if ("$($row.status)".ToUpper() -ne "MENUNGGU") {
    Write-Output ("GAGAL: status sekarang " + $row.status + " (hanya MENUNGGU yang bisa disetujui)")
    exit 1
}

$row.status = "DISETUJUI"
$row.disetujui_oleh = "$admin".Trim()
$row.waktu_persetujuan = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

$rows | Export-Csv -Path $QueueFile -NoTypeInformation -Encoding UTF8 -Delimiter ","

Write-Output ("DISETUJUI: " + $id)
Write-Output ("ADMIN: " + $row.disetujui_oleh)
Write-Output ("WAKTU: " + $row.waktu_persetujuan)
Write-Output "PESAN: database BELUM diubah. Jalankan terapkan_koreksi.ps1 -id <id> -admin <nama> -setuju"
