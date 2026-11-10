# ============================================================================
# lihat_masukan.ps1 - menampilkan masukan warga soal layanan (khusus admin).
#
#   .\lihat_masukan.ps1            -> 5 masukan terbaru + jumlah total
#   .\lihat_masukan.ps1 -Maks 10   -> 10 terbaru
#   .\lihat_masukan.ps1 -Maks 0    -> semua
#
# Keluaran: JSON satu baris {"jumlah":N,"items":[{"id_masukan","tanggal","isi"}]}
# ============================================================================

param(
    [int]$Maks = 5,
    [string]$Berkas = ""
)

[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$ErrorActionPreference = "Stop"

if (-not $Berkas) { $Berkas = Join-Path $PSScriptRoot "..\masukan_warga.csv" }

$items = New-Object System.Collections.ArrayList
$total = 0

if (Test-Path $Berkas) {
    $rows = @(Import-Csv -Path $Berkas -Encoding UTF8)
    $total = $rows.Count
    [array]::Reverse($rows)
    if ($Maks -gt 0 -and $rows.Count -gt $Maks) { $rows = @($rows[0..($Maks - 1)]) }
    foreach ($r in $rows) {
        $null = $items.Add([pscustomobject]@{
            id_masukan = [string]$r.id_masukan
            tanggal    = [string]$r.tanggal
            isi        = [string]$r.isi
        })
    }
}

$hasil = [pscustomobject]@{
    jumlah = $total
    items  = @($items.ToArray())
}
Write-Output ($hasil | ConvertTo-Json -Compress -Depth 5)