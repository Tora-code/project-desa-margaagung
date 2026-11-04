# ============================================================================
# ringkasan_antrean_pagi.ps1 - Ringkasan surat yang BELUM DIPROSES, dikirim ke
# petugas desa saat jam kerja dimulai (dijadwalkan 09.00 oleh job cron).
#
# Ganti dari cara lama: dulu antrean yang masuk malam ditahan (jam tenang) lalu
# dikirim sekaligus keesokan paginya. Sejak 2026-11-02 jam tenang dimatikan dan
# notifikasi antrean baru keluar segera, jadi yang dibutuhkan pagi bukan lagi
# "kiriman tertahan" tapi DAFTAR surat yang masih menunggu persetujuan petugas -
# termasuk yang sudah menunggu sejak kemarin.
#
# Isi pesan: hanya nomor surat, jenis, nama pemohon, dan kapan masuk.
# TIDAK ADA NIK, alamat, keperluan, atau isi surat; skrip memeriksa sendiri
# pesannya sebelum dikirim dan membatalkan pengiriman (BOCORAN_DITOLAK) kalau
# data itu ikut terbawa.
#
# Tidak menyentuh notifikasi_antrean.csv: ini ringkasan, bukan kabar per-surat,
# jadi surat yang masih menunggu tetap muncul tiap pagi sampai diproses petugas.
#
# Pakai:
#   .\ringkasan_antrean_pagi.ps1                      # susun & kirim
#   .\ringkasan_antrean_pagi.ps1 -DryRun              # lihat isi pesan, jangan kirim
#   .\ringkasan_antrean_pagi.ps1 -Sekarang "2026-11-03 09:00"
#
# Keluaran JSON: { ok, jumlah, pesan, dikirim_ke, ada_pesan_tambahan }
# Kalau tidak ada surat menunggu: { ok: true, jumlah: 0, pesan: "" } (diam).
# ============================================================================

param(
    [string]$AntreanDir = "",
    [string]$AdminFile = "",
    [string]$Sekarang = "",
    [int]$Maks = 15,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

# Keluaran JSON memuat emoji; tanpa ini Windows PowerShell 5.1 menulis stdout
# pakai codepage OEM dan emoji jadi "??" saat dibaca pemanggil.
try { [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false) } catch { }

if (-not $AntreanDir) {
    $AntreanDir = Join-Path (Split-Path -Parent $PSScriptRoot) "Antrean-surat"
}
if (-not $AdminFile) { $AdminFile = Join-Path $PSScriptRoot "admin_wa.csv" }
$csv = Join-Path $AntreanDir "antrean_surat.csv"
$logKirim = Join-Path $AntreanDir "notifikasi_antrean.txt"

function Keluar-Json($obj) {
    $obj | ConvertTo-Json -Depth 5 -Compress
    exit 0
}

$waktu = if ($Sekarang.Trim()) { [datetime]::ParseExact($Sekarang.Trim(), "yyyy-MM-dd HH\:mm", $null) } else { Get-Date }

function Samarkan-Nomor($n) {
    $d = ($n -replace '[^0-9]', '')
    if ($d.Length -lt 8) { return "****" }
    return $d.Substring(0, 4) + "****" + $d.Substring($d.Length - 4)
}

$BULAN = @("Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
    "Agustus", "September", "Oktober", "November", "Desember")

# "2026-11-02 15:40" -> "2 November 2026, 15.40"
function Tanggal-Manusiawi($t) {
    $x = "$t".Trim()
    if ($x -match '^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})') {
        return ("{0} {1} {2}, {3}.{4}" -f [int]$Matches[3], $BULAN[[int]$Matches[2] - 1], [int]$Matches[1], $Matches[4], $Matches[5])
    }
    if ($x -match '^(\d{4})-(\d{2})-(\d{2})') {
        return ("{0} {1} {2}" -f [int]$Matches[3], $BULAN[[int]$Matches[2] - 1], [int]$Matches[1])
    }
    return $x
}

if (-not (Test-Path $csv)) {
    Keluar-Json @{ ok = $false; error = "DATA_ANTREAN_TIDAK_DITEMUKAN" }
}

$semua = @(Import-Csv -Path $csv -Encoding UTF8)
$menunggu = @($semua |
    Where-Object { "$($_.status)".Trim().ToUpper() -like "MENUNGGU*" } |
    Sort-Object { "$($_.tanggal_pengajuan)".Trim() })

if ($menunggu.Count -eq 0) {
    Keluar-Json @{ ok = $true; jumlah = 0; pesan = ""; dikirim_ke = @(); ada_pesan_tambahan = $false }
}

# --- susun pesan ------------------------------------------------------------
$baris = @()
$baris += "📋 Antrean pagi - $($menunggu.Count) surat belum diproses"
$baris += ""
$nomor = 0
foreach ($b in ($menunggu | Select-Object -First $Maks)) {
    $nomor++
    $baris += "$nomor. $("$($b.id_surat)".Trim()) · $("$($b.jenis_surat)".Trim())"
    $baris += "   a.n. $("$($b.nama)".Trim()) · masuk $(Tanggal-Manusiawi "$($b.tanggal_pengajuan)")"
}
$sisa = $menunggu.Count - $nomor
if ($sisa -gt 0) { $baris += "... dan $sisa surat lainnya." }
$baris += ""
$baris += "Balas 1 untuk membuka antrean, lalu pilih surat yang mau diproses."
$teks = ($baris -join "`n")

# --- penjaga kebocoran: NIK / alamat / keperluan / isi surat tidak boleh ikut --
function Periksa-Bersih($pesan, $barisAntrean) {
    foreach ($r in $barisAntrean) {
        $nik = ("$($r.nik)" -replace '[^0-9]', '')
        if ($nik -and $pesan -match [regex]::Escape($nik)) { return $false }
        if ($nik -and $nik.Length -ge 8 -and $pesan -match [regex]::Escape($nik.Substring($nik.Length - 8))) { return $false }
        foreach ($k in @("data_inti", "keperluan", "kontak")) {
            $v = "$($r.$k)".Trim()
            if ($v.Length -ge 8 -and $pesan -match [regex]::Escape($v.Substring(0, 8))) { return $false }
        }
        if ($pesan -match 'Dusun|RT\s?[0-9]|RW\s?[0-9]') { return $false }
    }
    return $true
}

if (-not (Periksa-Bersih $teks $menunggu)) {
    Keluar-Json @{ ok = $false; error = "BOCORAN_DITOLAK"; pesan = "Ringkasan memuat data yang dilarang (NIK/alamat/isi surat) - tidak dikirim." }
}

if ($DryRun) {
    Keluar-Json @{ ok = $true; dry_run = $true; jumlah = $menunggu.Count; pesan = $teks; dikirim_ke = @(); ada_pesan_tambahan = ($sisa -gt 0) }
}

# --- penerima: semua petugas aktif -----------------------------------------
$penerima = @()
if (Test-Path $AdminFile) {
    $penerima = @(Import-Csv -Path $AdminFile -Encoding UTF8 |
        Where-Object { "$($_.aktif)".Trim() -eq "1" } |
        ForEach-Object {
            $d = ("$($_.nomor)" -replace '[^0-9]', '')
            if ($d.StartsWith("0")) { $d = "62" + $d.Substring(1) }
            elseif ($d.StartsWith("8")) { $d = "62" + $d }
            $d
        } | Where-Object { $_ -match '^628[0-9]{7,12}$' })
}
if ($penerima.Count -eq 0) {
    Keluar-Json @{ ok = $false; error = "TIDAK_ADA_ADMIN_AKTIF" }
}

$hermes = Join-Path $env:LOCALAPPDATA "hermes\bin\hermes.exe"
if (-not (Test-Path $hermes)) { $hermes = "hermes" }

$terkirimKe = @()
$gagal = @()
foreach ($nomor in $penerima) {
    $hasil = & $hermes --profile surat-desa send -t "whatsapp:$nomor" --json $teks 2>&1
    $teksHasil = ($hasil | Out-String)
    if ($teksHasil -match '"success"\s*:\s*true') { $terkirimKe += (Samarkan-Nomor $nomor) }
    else { $gagal += (Samarkan-Nomor $nomor) }
    $hasil = $null
}

# --- jejak audit: pesan yang benar-benar dikirim (UTF-8 tanpa BOM) ----------
$infoKirim = "GAGAL"
if ($terkirimKe.Count -gt 0) { $infoKirim = "TERKIRIM ke " + (($terkirimKe | Select-Object -Unique) -join ", ") }
$catatLog = "=== {0} | RINGKASAN PAGI | {1} | {2} surat menunggu ===`n{3}`n`n" -f `
    $waktu.ToString("yyyy-MM-dd HH\:mm"), $infoKirim, $menunggu.Count, $teks
[IO.File]::AppendAllText($logKirim, $catatLog, (New-Object Text.UTF8Encoding($false)))

Keluar-Json @{
    ok                 = $true
    jumlah             = $menunggu.Count
    pesan              = $teks
    dikirim_ke         = @($terkirimKe | Select-Object -Unique)
    gagal              = @($gagal | Select-Object -Unique)
    ada_pesan_tambahan = ($sisa -gt 0)
}
