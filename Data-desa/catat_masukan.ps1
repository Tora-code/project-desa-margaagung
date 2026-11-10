# ============================================================================
# catat_masukan.ps1 - mencatat masukan/saran warga soal layanan (ANONIM).
#
#   .\catat_masukan.ps1 -Isi "Menu suratnya sudah jelas, tapi balasan agak lambat"
#
# Tidak menyimpan nama maupun nomor pengirim. Keluaran: JSON satu baris.
#   {"ok":true,"id_masukan":"MS-20261110-001","jumlah":1,"pesan":""}
# ============================================================================

param(
    [string]$Isi = "",
    [string]$Berkas = "",
    [string]$Sekarang = ""
)

[Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
$ErrorActionPreference = "Stop"

if (-not $Berkas) { $Berkas = Join-Path $PSScriptRoot "..\masukan_warga.csv" }

function Keluar($obj) {
    Write-Output ($obj | ConvertTo-Json -Compress)
    exit 0
}

function Kutip($s) { '"' + ([string]$s).Replace('"', '""') + '"' }

$teks = ([string]$Isi) -replace "\s+", " "
$teks = $teks.Trim()
if (-not $teks) { Keluar @{ ok = $false; id_masukan = ""; jumlah = 0; pesan = "MASUKAN_KOSONG" } }
if ($teks.Length -gt 1000) { $teks = $teks.Substring(0, 1000).Trim() }

if ($Sekarang) {
    $waktu = [datetime]::ParseExact($Sekarang.Trim(), "yyyy-MM-dd HH:mm", $null)
} else {
    $waktu = Get-Date
}
$tanggal = $waktu.ToString("yyyy-MM-dd HH\:mm")
$prefix = "MS-" + $waktu.ToString("yyyyMMdd") + "-"

$header = '"tanggal","id_masukan","isi"'
if (-not (Test-Path $Berkas)) {
    [IO.File]::WriteAllText($Berkas, $header + "`r`n", (New-Object Text.UTF8Encoding($true)))
}

$rows = @(Import-Csv -Path $Berkas -Encoding UTF8)
$urut = 0
foreach ($r in $rows) {
    $m = [regex]::Match(([string]$r.id_masukan).Trim(), '^MS-(\d{8})-(\d+)$')
    if ($m.Success) {
        $prefBaris = "MS-" + $m.Groups[1].Value + "-"
        if ($prefBaris -eq $prefix) {
            $n = [int]$m.Groups[2].Value
            if ($n -gt $urut) { $urut = $n }
        }
    }
}
$id = $prefix + ("{0:D3}" -f ($urut + 1))

$baris = (Kutip $tanggal) + "," + (Kutip $id) + "," + (Kutip $teks) + "`r`n"
[IO.File]::AppendAllText($Berkas, $baris, (New-Object Text.UTF8Encoding($false)))

# buktikan barisnya benar-benar ada sebelum melapor ke warga
$cek = @(Import-Csv -Path $Berkas -Encoding UTF8)
$ada = $false
foreach ($r in $cek) { if (([string]$r.id_masukan) -eq $id) { $ada = $true } }

Keluar @{ ok = $ada; id_masukan = $id; jumlah = $cek.Count; pesan = "" }