# ============================================================================
# siapkan_kirim.ps1 - Siapkan dokumen surat agar bisa dikirim ke WhatsApp.
#
# Kenapa perlu: Hermes hanya mengenali lampiran dari penanda `MEDIA:<path>` dan
# path itu TIDAK boleh mengandung spasi. Nama file surat selalu berspasi
# ("SR-... - Surat Keterangan Domisili - NAMA.docx"), jadi file disalin dulu ke
# folder Antrean-surat\_kirim dengan nama tanpa spasi, lalu path-nya dicetak
# sebagai JSON.
#
# Hanya MEMBACA antrean_surat.csv + menyalin file. Tidak mengubah antrean.
#
# Pakai:
#   .\siapkan_kirim.ps1 -Id "SR-20261027-002"          # satu surat
#   .\siapkan_kirim.ps1 -Status MENUNGGU                # daftar antrean + file
#   .\siapkan_kirim.ps1 -Status MENUNGGU -Maks 5
#   .\siapkan_kirim.ps1 -Id "SR-..." -TanpaData        # hanya path file
#
# Keluaran satu surat: { ok, id, nama, jenis, status, tanggal, keperluan,
#                        kontak, catatan, file }
# Keluaran daftar   : { ok, jumlah, sisa, total, status, items: [ ... ] }
# Gagal             : { ok: false, error: "..." }
# ============================================================================

param(
    [string]$Id = "",
    [string]$Status = "",
    [int]$Maks = 3,
    [string]$AntreanDir = "",
    [switch]$TanpaData
)

$ErrorActionPreference = "Stop"

if (-not $AntreanDir) {
    $AntreanDir = Join-Path (Split-Path -Parent $PSScriptRoot) "Antrean-surat"
}
$csv = Join-Path $AntreanDir "antrean_surat.csv"
$tujuanDir = Join-Path $AntreanDir "_kirim"

function Keluar-Json($obj) {
    $obj | ConvertTo-Json -Depth 5 -Compress
    exit 0
}

function Label-Status($s) {
    # Nilai ini yang dibaca model untuk ditampilkan ke WhatsApp: SELALU bahasa biasa,
    # tidak pernah kode internal. Nilai tak dikenal pun jangan diteruskan apa adanya.
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
        "SELESAI"                   { "sudah diambil" }
        "DITOLAK"                   { "ditolak" }
        default                     { "perlu dicek petugas" }
    }
}

# Catatan mentah di antrean memuat detail internal (kode pengecekan NIK, nama penyetuju,
# waktu persetujuan, nama file). Yang boleh tampil ke chat hanya versi bersih ini;
# catatan mentah tetap ada di file antrean sebagai jejak audit.
function Bersihkan-Catatan($t) {
    $x = "$t"
    if (-not $x.Trim()) { return "" }
    # (1) buang klausa yang memuat kode pengecekan internal
    $x = $x -replace '(?i)[^.]*\b(NAMA_COCOK|COCOK_SATU|NAMA_TIDAK_COCOK|TIDAK_COCOK|KEMBAR|NIK_BEDA|NIK_DB_BUKAN_16|NIK_BUKAN_16)\b[^.]*\.?', ''
    $x = $x -replace '(?i)NIK sama dengan database[^.]*\.?', ''
    $x = $x -replace '(?i)Nomor surat & TTD diisi admin[^.]*\.?', ''
    # (2) buang kalimat yang menyebut backup/jejak audit, nama penyetuju, atau nama file antrean
    $bagian = @([regex]::Split($x, '(?<=\.)\s+') | Where-Object {
            $_ -and ($_ -notmatch '(?i)backup|jejak audit|disetujui admin|antrean_(surat|ambil|koreksi)')
        })
    $x = ($bagian -join ' ')
    $x = $x -replace '\s{2,}', ' '
    return $x.Trim(' ', '.', ',', ';')
}

function Ada-NikBeda($t) {
    return ("$t" -match '(?i)NIK_BEDA|berbeda dengan database|NIK dipakai berbeda|NIK beda')
}

# Siapkan satu baris antrean -> objek siap kirim (atau $null kalau file hilang)
function Siapkan-Satu($baris) {
    $idb = "$($baris.id_surat)".Trim().ToUpper()
    $kandidat = @()
    if ("$($baris.file_draft)".Trim()) {
        $kandidat += (Join-Path $AntreanDir "$($baris.file_draft)".Trim())
    }
    $kandidat += @(Get-ChildItem -Path $AntreanDir -Filter "$idb*.docx" -File -ErrorAction SilentlyContinue |
        ForEach-Object { $_.FullName })

    $sumber = $kandidat | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $sumber) { return $null }

    if (-not (Test-Path $tujuanDir)) { New-Item -ItemType Directory -Path $tujuanDir -Force | Out-Null }
    $namaAman = ([IO.Path]::GetFileNameWithoutExtension($sumber) -replace '[^A-Za-z0-9\-_]', '-') -replace '-{2,}', '-'
    $tujuan = Join-Path $tujuanDir ($namaAman + ".docx")
    Copy-Item -LiteralPath $sumber -Destination $tujuan -Force

    if ($TanpaData) { return @{ ok = $true; id = $idb; file = $tujuan } }

    $catatan = "$($baris.catatan)".Trim()
    return @{
        ok           = $true
        id           = $idb
        nama         = "$($baris.nama)".Trim()
        jenis        = "$($baris.jenis_surat)".Trim()
        status       = (Label-Status "$($baris.status)")
        status_teks  = (Label-Status "$($baris.status)")
        tanggal      = "$($baris.tanggal_pengajuan)".Trim()
        keperluan    = "$($baris.keperluan)".Trim()
        kontak       = "$($baris.kontak)".Trim()
        catatan      = $catatan
        catatan_chat = (Bersihkan-Catatan $catatan)
        nik_beda     = (Ada-NikBeda $catatan)
        file         = $tujuan
    }
}

# --- Mode daftar: semua surat dengan status tertentu -----------------------
if ($Status) {
    if (-not (Test-Path $csv)) {
        Keluar-Json @{ ok = $false; error = "DATA_ANTREAN_TIDAK_DITEMUKAN" }
    }
    $cocok = @(Import-Csv -Path $csv -Encoding UTF8 | Where-Object {
            $s = "$($_.status)".Trim().ToUpper()
            if ("$Status".Trim().ToUpper() -eq "MENUNGGU") { $s -like "MENUNGGU*" } else { $s -eq "$Status".Trim().ToUpper() }
        })

    if ($cocok.Count -eq 0) {
        Keluar-Json @{ ok = $true; jumlah = 0; sisa = 0; total = 0; status = (Label-Status $Status); items = @() }
    }

    if ($Maks -lt 1) { $Maks = 1 }
    $items = @()
    foreach ($b in ($cocok | Select-Object -First $Maks)) {
        $o = Siapkan-Satu $b
        if ($o) { $items += $o }
    }
    Keluar-Json @{
        ok     = $true
        jumlah = $items.Count
        sisa   = [Math]::Max(0, $cocok.Count - $items.Count)
        total  = $cocok.Count
        status = (Label-Status $Status)
        items  = $items
    }
}

# --- Mode satu surat -------------------------------------------------------
if (-not $Id) {
    Keluar-Json @{ ok = $false; error = "PERLU_ID_ATAU_STATUS" }
}

$idBersih = ($Id -replace '\s', '').ToUpper()
if ($idBersih -notmatch '^SR-\d{8}-\d{3}$') {
    Keluar-Json @{ ok = $false; error = "FORMAT_ID_TIDAK_VALID"; id = $Id }
}

$baris = $null
if (Test-Path $csv) {
    $baris = @(Import-Csv -Path $csv -Encoding UTF8 |
        Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $idBersih }) | Select-Object -First 1
}
if (-not $baris) {
    Keluar-Json @{ ok = $false; error = "ID_TIDAK_ADA_DI_ANTREAN"; id = $idBersih }
}

$hasil = Siapkan-Satu $baris
if (-not $hasil) {
    Keluar-Json @{ ok = $false; error = "FILE_DOKUMEN_TIDAK_DITEMUKAN"; id = $idBersih }
}
Keluar-Json $hasil
