# ============================================================================
# status_surat.ps1 - Ubah status surat di antrean + catat antrean pengambilan.
#
# Aksi (status ditulis dalam bahasa biasa, sama dengan yang tampil di WhatsApp):
#   setujui  -> "sudah disetujui"  (WAJIB -Nomor admin terdaftar)
#   tunda    -> "sudah disetujui"  (WAJIB -Nomor admin terdaftar; belum bisa diambil)
#   siap     -> "siap diambil"     (+ catat tanggal siap diambil)
#   ambil    -> "sudah diambil"    (warga sudah mengambil)
#
# Penyetuju dicatat di file antrean, tepat saat persetujuan terjadi:
#   Disetujui_Oleh     : NAMA admin dari Data-desa\admin_wa.csv (hasil cek_admin.ps1),
#                        BUKAN nama yang diketik pengirim
#   Waktu_Persetujuan  : tanggal & jam persetujuan (yyyy-MM-dd HH:mm)
#   status             : sudah disetujui
# Nama penyetuju dan waktu persetujuan TIDAK pernah dikembalikan ke chat.
#
# Gerbang nomor: setujui/tunda DITOLAK kalau -Nomor bukan nomor admin aktif
# (dicek lewat cek_admin.ps1). Aksi siap/ambil boleh tanpa -Nomor, tapi kalau
# -Nomor diberikan maka harus admin juga.
#
# File yang disentuh:
#   Antrean-surat\antrean_surat.csv    (status, Disetujui_Oleh, Waktu_Persetujuan)
#   Antrean-surat\antrean_ambil.csv    (tanggal siap, kapan dikabari)
#   Antrean-surat\log_persetujuan.csv  (jejak audit: siapa, kapan, dari status apa)
# Backup bertanggal antrean_surat.csv dibuat SEBELUM perubahan diterapkan.
#
# Pakai:
#   .\status_surat.ps1 -Id "SR-20261027-001" -Aksi setujui -Nomor "6281234567890"
#   .\status_surat.ps1 -Id "SR-20261027-001" -Aksi siap -Tanggal "28 Oktober 2026"
#   .\status_surat.ps1 -Id "SR-20261027-001" -Aksi ambil
# ============================================================================

param(
    [Parameter(Mandatory = $true)][string]$Id,
    [Parameter(Mandatory = $true)][ValidateSet("setujui", "tunda", "siap", "ambil")][string]$Aksi,
    [string]$Nomor = "",
    [string]$Tanggal = "",
    [string]$Catatan = "",
    [string]$AntreanDir = "",
    [string]$LogFile = "",
    [string]$Admin = ""   # dipertahankan agar pemanggil lama tidak error; NILAI INI DIIABAIKAN
)

$ErrorActionPreference = "Stop"

if (-not $AntreanDir) {
    $AntreanDir = Join-Path (Split-Path -Parent $PSScriptRoot) "Antrean-surat"
}
$csv = Join-Path $AntreanDir "antrean_surat.csv"
$csvAmbil = Join-Path $AntreanDir "antrean_ambil.csv"
$backupDir = Join-Path $AntreanDir "_backup"
if (-not $LogFile) { $LogFile = Join-Path $AntreanDir "log_persetujuan.csv" }

function Keluar-Json($obj) {
    $obj | ConvertTo-Json -Depth 4 -Compress
    exit 0
}

# Label bahasa biasa untuk ditampilkan ke WhatsApp. Kode internal lama tetap
# dipetakan supaya baris lama tidak pernah bocor mentah ke chat.
function Label-Status($s) {
    switch ("$s".Trim().ToUpper()) {
        "MENUNGGU_VERIFIKASI_ADMIN" { "menunggu" }
        "MENUNGGU"                  { "menunggu" }
        "DISETUJUI_ADMIN"           { "sudah disetujui" }
        "DISETUJUI"                 { "sudah disetujui" }
        "SUDAH DISETUJUI"           { "sudah disetujui" }
        "DITUNDA"                   { "sudah disetujui" }
        "SIAP_DIAMBIL"              { "siap diambil" }
        "SIAP DIAMBIL"              { "siap diambil" }
        "DIAMBIL"                   { "sudah diambil" }
        "SUDAH DIAMBIL"             { "sudah diambil" }
        default                     { "perlu dicek petugas" }
    }
}

# Jalankan cek_admin.ps1 di proses terpisah: skrip itu memanggil `exit`, jadi tidak
# boleh dijalankan in-process (akan mengakhiri skrip ini).
function Cek-NomorAdmin($nomor) {
    $skrip = Join-Path $PSScriptRoot "cek_admin.ps1"
    if (-not (Test-Path $skrip)) {
        return @{ status = "PEMERIKSA_TIDAK_ADA"; nama = "" }
    }
    $lama = $ErrorActionPreference
    $ErrorActionPreference = "Continue"   # stderr dari native tidak boleh mematikan skrip
    try {
        $hasil = & powershell -NoProfile -ExecutionPolicy Bypass -File $skrip -Nomor "$nomor" 2>$null
    }
    finally {
        $ErrorActionPreference = $lama
    }
    $baris = @($hasil | ForEach-Object { "$_".Trim() })
    $statusBaris = @($baris | Where-Object { $_ -like "STATUS:*" })
    $namaBaris = @($baris | Where-Object { $_ -like "NAMA:*" })
    # JEBAKAN: `$null -replace 'x',''` mengembalikan Object[] KOSONG, bukan string, dan
    # memanggil .Trim() langsung di situ melempar "does not contain a method named 'Trim'"
    # (persis kasus nomor bukan admin, yang justru harus ditolak dengan rapi).
    $status = if ($statusBaris.Count -gt 0) { ([string]($statusBaris[0] -replace '^STATUS:\s*', '')).Trim() } else { "" }
    $nama = if ($namaBaris.Count -gt 0) { ([string]($namaBaris[0] -replace '^NAMA:\s*', '')).Trim() } else { "" }
    return @{ status = $status; nama = $nama }
}

$idBersih = ($Id -replace '\s', '').ToUpper()
if ($idBersih -notmatch '^SR-\d{8}-\d{3}$') {
    Keluar-Json @{ ok = $false; error = "FORMAT_ID_TIDAK_VALID"; id = $Id }
}
if (-not (Test-Path $csv)) {
    Keluar-Json @{ ok = $false; error = "DATA_ANTREAN_TIDAK_DITEMUKAN" }
}

# --- gerbang nomor admin ---------------------------------------------------
$wajibNomor = ($Aksi -eq "setujui" -or $Aksi -eq "tunda")
$namaAdmin = ""
$nomorBersih = ($Nomor -replace '[^0-9]', '')

if ($wajibNomor -and -not $nomorBersih) {
    Keluar-Json @{ ok = $false; error = "PERSETUJUAN_TANPA_NOMOR"; id = $idBersih; pesan = "Persetujuan hanya dari nomor WhatsApp petugas desa yang terdaftar." }
}
if ($nomorBersih) {
    $cek = Cek-NomorAdmin $Nomor
    if ($cek.status -ne "ADMIN") {
        Keluar-Json @{ ok = $false; error = "PERSETUJUAN_DITOLAK"; id = $idBersih; alasan = "NOMOR_BUKAN_ADMIN_TERDAFTAR"; status_nomor = $cek.status }
    }
    $namaAdmin = $cek.nama
    if (-not $namaAdmin) { $namaAdmin = "PETUGAS" }
}

$statusBaru = switch ($Aksi) {
    "setujui" { "sudah disetujui" }
    "tunda"   { "sudah disetujui" }   # sudah disetujui, tapi belum bisa diambil
    "siap"    { "siap diambil" }
    "ambil"   { "sudah diambil" }
}

$semua = @(Import-Csv -Path $csv -Encoding UTF8)
if ($semua.Count -eq 0) {
    Keluar-Json @{ ok = $false; error = "ANTREAN_KOSONG" }
}

# kolom pencatatan penyetuju mungkin belum ada di file lama -> tambahkan ke semua baris
$butuhKolom = @("Disetujui_Oleh", "Waktu_Persetujuan") | Where-Object { -not ($semua[0].PSObject.Properties[$_]) }
if ($butuhKolom.Count -gt 0) {
    foreach ($r in $semua) {
        foreach ($k in $butuhKolom) {
            $r | Add-Member -NotePropertyName $k -NotePropertyValue "" -Force
        }
    }
}

$baris = $semua | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $idBersih } | Select-Object -First 1
if (-not $baris) {
    Keluar-Json @{ ok = $false; error = "ID_TIDAK_ADA_DI_ANTREAN"; id = $idBersih }
}
if ($Aksi -eq "siap" -and -not $Tanggal.Trim()) {
    Keluar-Json @{ ok = $false; error = "TANGGAL_SIAP_KOSONG"; id = $idBersih }
}

# --- backup dulu (jejak audit kondisi SEBELUM) -----------------------------
if (-not (Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
$cap = Get-Date -Format "yyyy-MM-dd_HHmmss"
$backupPath = Join-Path $backupDir "antrean_surat.csv.$cap.bak"
Copy-Item -LiteralPath $csv -Destination $backupPath -Force

# --- ubah status + catat penyetuju ----------------------------------------
$statusLama = "$($baris.status)".Trim()
$waktuAksi = (Get-Date).ToString("yyyy-MM-dd HH\:mm")
$baris.status = $statusBaru
if ($Aksi -eq "setujui" -or $Aksi -eq "tunda") {
    $baris.Disetujui_Oleh = $namaAdmin
    $baris.Waktu_Persetujuan = $waktuAksi
}
$semua | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

# --- catat antrean pengambilan -------------------------------------------
if ($Aksi -eq "siap" -or $Aksi -eq "ambil" -or $Aksi -eq "tunda") {
    $ambil = @()
    if (Test-Path $csvAmbil) { $ambil = @(Import-Csv -Path $csvAmbil -Encoding UTF8) }
    $ada = $ambil | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $idBersih } | Select-Object -First 1
    if ($ada) {
        if ($Aksi -eq "siap") { $ada.tanggal_siap = $Tanggal.Trim() }
        if ($Catatan.Trim()) { $ada.catatan = $Catatan.Trim() }
        $ambil = @($ambil | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -ne $idBersih }) + @($ada)
    }
    else {
        $ambil += [pscustomobject]@{
            id_surat      = $idBersih
            nama          = "$($baris.nama)".Trim()
            kontak        = "$($baris.kontak)".Trim()
            tanggal_siap  = if ($Aksi -eq "siap") { $Tanggal.Trim() } else { "" }
            dikabari_pada = ""
            catatan       = $Catatan.Trim()
        }
    }
    $ambil | Sort-Object id_surat | Export-Csv -Path $csvAmbil -NoTypeInformation -Encoding UTF8
}

# --- jejak audit di log (detail teknis; bukan untuk chat) ------------------
try {
    $logAda = Test-Path $LogFile
    $logBaris = [pscustomobject]@{
        waktu          = (Get-Date).ToString("yyyy-MM-dd HH\:mm\:ss")
        id_surat       = $idBersih
        aksi           = $Aksi
        status_lama    = $statusLama
        status_baru    = $statusBaru
        disetujui_oleh = $namaAdmin
        nomor_admin    = $nomorBersih
        backup_sebelum = (Split-Path $backupPath -Leaf)
        catatan        = $Catatan.Trim()
    }
    if ($logAda) {
        $logBaris | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8 -Append
    }
    else {
        $logBaris | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8
    }
}
catch {
    # log tidak boleh menggagalkan perubahan status yang sudah sah
}

# Keluaran untuk chat: hanya bahasa biasa. Nama penyetuju & waktu TIDAK dikembalikan.
Keluar-Json @{
    ok               = $true
    id               = $idBersih
    nama             = "$($baris.nama)".Trim()
    status_lama      = $statusLama
    status_baru      = $statusBaru
    status_teks_lama = (Label-Status $statusLama)
    status_teks_baru = (Label-Status $statusBaru)
    tercatat         = [bool]($Aksi -eq "setujui" -or $Aksi -eq "tunda")
    tanggal          = $Tanggal.Trim()
    kontak           = "$($baris.kontak)".Trim()
}
