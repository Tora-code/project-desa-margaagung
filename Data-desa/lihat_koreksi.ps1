# ============================================================================
# lihat_koreksi.ps1 - Melihat antrean koreksi NIK (khusus admin).
#
#   .\lihat_koreksi.ps1                    -> semua
#   .\lihat_koreksi.ps1 -status MENUNGGU   -> hanya yang menunggu
# ============================================================================

param(
    [string]$status = "",
    [string]$QueueFile = ""
)

if (-not $QueueFile) { $QueueFile = Join-Path $PSScriptRoot "antrean_koreksi.csv" }

if (-not (Test-Path $QueueFile)) {
    Write-Output "ANTREAN: kosong"
    exit
}

$rows = @(Import-Csv -Path $QueueFile -Encoding UTF8)
if ($status) { $rows = @($rows | Where-Object { "$($_.status)".ToUpper() -eq $status.ToUpper() }) }

if ($rows.Count -eq 0) {
    Write-Output "ANTREAN: kosong"
    exit
}

Write-Output ("ANTREAN: " + $rows.Count + " catatan")
foreach ($r in $rows) {
    Write-Output ("- " + $r.id_koreksi + " | " + $r.waktu + " | " + $r.jenis + " | NIK ketikan " + $r.nik_ketikan + " vs database " + $r.nik_database + " | " + $r.nama + " | " + $r.dusun + " | " + $r.status + $(if ($r.disetujui_oleh) { " oleh " + $r.disetujui_oleh } else { "" }))
}
