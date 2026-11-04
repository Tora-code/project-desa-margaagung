# ============================================================================
# kabari_antrean_admin.ps1 - Notifikasi antrean baru ke petugas desa (WhatsApp).
#
# Format notifikasi (TIGA baris, persis):
#   📩 Antrean baru masuk
#   <ID SR> · <Jenis surat>
#   a.n. <Nama pemohon> · <Tanggal, jam masuk>
#
# Yang TIDAK boleh ada di notifikasi: NIK, alamat lengkap, isi surat, keperluan.
# Skrip memeriksa sendiri pesannya sebelum dikirim; kalau ada NIK/isi surat yang
# ikut terbawa, pengiriman DIBATALKAN (error BOCORAN_DITOLAK).
#
# Jam tenang 21.00-07.00 NONAKTIF sejak 2026-11-02 (permintaan pemilik): antrean
# baru langsung dikirim kapan pun, siang maupun malam. Jalur lama masih ada di
# bawah dan bisa dihidupkan lagi dengan -PakaiJamTenang (antrean malam dicatat
# "menunggu pagi", lalu dikirim sekaligus begitu jam tenang selesai).
#
# Ringkasan surat yang belum diproses tiap pagi (09.00) ada di skrip terpisah:
# ringkasan_antrean_pagi.ps1, dijalankan job cron-nya sendiri.
#
# Notifikasi ini hanya kabar, BUKAN persetujuan. Persetujuan tetap lewat
# status_surat.ps1 -Aksi setujui -Nomor <nomor admin>.
#
# Pakai:
#   .\kabari_antrean_admin.ps1                       # periksa antrean, kirim yang baru
#   .\kabari_antrean_admin.ps1 -Id "SR-20261028-001" # kabari satu surat saja
#   .\kabari_antrean_admin.ps1 -DryRun               # susun pesan, jangan kirim
#   .\kabari_antrean_admin.ps1 -DryRun -Sekarang "2026-11-02 23:30"   # uji kirim malam
#   .\kabari_antrean_admin.ps1 -PakaiJamTenang -Sekarang "2026-11-02 23:30"  # uji jam tenang
#
# Keluaran JSON: { ok, mode, jumlah, tertahan, dikirim_ke, pesan }
# ============================================================================

param(
    [string]$Id = "",
    [string]$AntreanDir = "",
    [string]$AdminFile = "",
    [string]$StateFile = "",
    [string]$Sekarang = "",
    [switch]$DryRun,
    [switch]$PakaiJamTenang   # jam tenang NONAKTIF secara bawaan sejak 2026-11-02
)

$ErrorActionPreference = "Stop"

# Keluaran JSON berisi emoji/tanda baca non-ASCII; tanpa ini Windows PowerShell 5.1
# menulis stdout pakai codepage OEM dan emoji jadi "??" saat dibaca pemanggil.
try { [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false) } catch { }

if (-not $AntreanDir) {
    $AntreanDir = Join-Path (Split-Path -Parent $PSScriptRoot) "Antrean-surat"
}
if (-not $AdminFile) { $AdminFile = Join-Path $PSScriptRoot "admin_wa.csv" }
$csv = Join-Path $AntreanDir "antrean_surat.csv"
if (-not $StateFile) { $StateFile = Join-Path $AntreanDir "notifikasi_antrean.csv" }

function Keluar-Json($obj) {
    $obj | ConvertTo-Json -Depth 5 -Compress
    exit 0
}

function Samarkan-Nomor($n) {
    $d = ($n -replace '[^0-9]', '')
    if ($d.Length -lt 8) { return "****" }
    return $d.Substring(0, 4) + "****" + $d.Substring($d.Length - 4)
}

$BULAN = @("Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
    "Agustus", "September", "Oktober", "November", "Desember")

# "2026-10-28 11:38" -> "28 Oktober 2026, 11.38"
function Tanggal-Manusiawi($t) {
    $x = "$t".Trim()
    if ($x -match '^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})') {
        $th = [int]$Matches[1]; $bl = [int]$Matches[2]; $tg = [int]$Matches[3]
        $jm = $Matches[4]; $mn = $Matches[5]
        return ("{0} {1} {2}, {3}.{4}" -f $tg, $BULAN[$bl - 1], $th, $jm, $mn)
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
$idBersih = ($Id -replace '\s', '').ToUpper()
if ($idBersih) {
    $semua = @($semua | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $idBersih })
    if ($semua.Count -eq 0) {
        Keluar-Json @{ ok = $false; error = "ID_TIDAK_ADA_DI_ANTREAN"; id = $idBersih }
    }
}

# --- waktu sekarang & jam tenang ------------------------------------------
$waktu = if ($Sekarang.Trim()) { [datetime]::ParseExact($Sekarang.Trim(), "yyyy-MM-dd HH\:mm", $null) } else { Get-Date }
$jam = $waktu.Hour
# Jam tenang 21.00-07.00 NONAKTIF secara bawaan sejak 2026-11-02: pemilik meminta
# notifikasi antrean baru keluar segera, tanpa ditahan sampai pagi. Jalur jam tenang
# sengaja TIDAK dihapus, hanya dimatikan - hidupkan lagi dengan -PakaiJamTenang.
$jamTenang = ($PakaiJamTenang -and ($jam -ge 21 -or $jam -lt 7))

# --- catatan notifikasi yang sudah dikirim ---------------------------------
$catat = @()
if (Test-Path $StateFile) { $catat = @(Import-Csv -Path $StateFile -Encoding UTF8) }
function Status-Catat($sid) {
    $r = $catat | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq $sid } | Select-Object -First 1
    if (-not $r) { return "" }
    return "$($r.status)".Trim().ToLower()
}
function Tulis-Catat($sid, $masuk, $status, $tujuan) {
    $barisBaru = [pscustomobject]@{
        id_surat          = $sid
        tanggal_masuk     = $masuk
        waktu_notifikasi  = if ($status -eq "terkirim") { $waktu.ToString("yyyy-MM-dd HH\:mm") } else { "" }
        status            = $status
        tujuan            = $tujuan
    }
    $script:catat = @($script:catat | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -ne $sid }) + @($barisBaru)
}

$belum = @($semua | Where-Object {
        $s = Status-Catat ("$($_.id_surat)".Trim().ToUpper())
        ($s -ne "terkirim")
    })
$tertahan = @($semua | Where-Object { (Status-Catat ("$($_.id_surat)".Trim().ToUpper())) -eq "menunggu_pagi" })
$baru = @($belum | Where-Object { (Status-Catat ("$($_.id_surat)".Trim().ToUpper())) -eq "" })

# --- jam tenang: catat dulu, jangan kirim ---------------------------------
if ($jamTenang) {
    foreach ($b in $belum) {
        Tulis-Catat ("$($b.id_surat)".Trim().ToUpper()) "$($b.tanggal_pengajuan)".Trim() "menunggu_pagi" ""
    }
    if ($DryRun) {
        Keluar-Json @{ ok = $true; dry_run = $true; mode = "jam tenang"; jumlah = $belum.Count; tertahan = @($belum | ForEach-Object { "$($_.id_surat)".Trim() }); pesan = "" }
    }
    if (-not (Test-Path $StateFile) -or $catat.Count -gt 0) {
        $catat | Sort-Object id_surat | Export-Csv -Path $StateFile -NoTypeInformation -Encoding UTF8
    }
    Keluar-Json @{
        ok        = $true
        mode      = "jam tenang (21.00-07.00)"
        jumlah    = $belum.Count
        tertahan  = @($belum | ForEach-Object { "$($_.id_surat)".Trim() })
        keterangan = if ($belum.Count -gt 0) { "Notifikasi ditahan sampai pagi, lalu dikirim sekaligus." } else { "Tidak ada antrean baru." }
    }
}

# --- susun pesan ----------------------------------------------------------
function Pesan-Satu($baris) {
    $tigaBaris = @(
        "📩 Antrean baru masuk",
        ("{0} · {1}" -f "$($baris.id_surat)".Trim().ToUpper(), "$($baris.jenis_surat)".Trim()),
        ("a.n. {0} · {1}" -f "$($baris.nama)".Trim(), (Tanggal-Manusiawi "$($baris.tanggal_pengajuan)"))
    )
    return ($tigaBaris -join "`n")
}

# penjaga: notifikasi tidak boleh membawa NIK, alamat, isi surat, atau keperluan
function Periksa-Bersih($pesan, $baris) {
    $nik = ("$($baris.nik)" -replace '[^0-9]', '')
    if ($nik -and $pesan -match [regex]::Escape($nik)) { return $false }
    if ($nik -and $nik.Length -ge 8 -and $pesan -match [regex]::Escape($nik.Substring($nik.Length - 8))) { return $false }
    foreach ($k in @("data_inti", "keperluan", "kontak")) {
        $v = "$($baris.$k)".Trim()
        if ($v.Length -ge 8 -and $pesan -match [regex]::Escape($v.Substring(0, 8))) { return $false }
    }
    if ($pesan -match 'Dusun|RT\s?[0-9]|RW\s?[0-9]') { return $false }
    return $true
}

$pesanKirim = @()
if ($tertahan.Count -gt 0) {
    # yang tertahan jam tenang -> satu pesan sekaligus
    $blok = @($tertahan | ForEach-Object { Pesan-Satu $_ })
    $pesanKirim += @{ ids = @($tertahan | ForEach-Object { "$($_.id_surat)".Trim() }); teks = ($blok -join "`n`n"); jenis = "sekali-sekaligus" }
}
if ($baru.Count -gt 0) {
    # yang baru masuk di jam kerja -> satu pesan per antrean
    foreach ($b in $baru) {
        $pesanKirim += @{ ids = @("$($b.id_surat)".Trim()); teks = (Pesan-Satu $b); jenis = "langsung" }
    }
}

if ($pesanKirim.Count -eq 0) {
    Keluar-Json @{ ok = $true; mode = "tidak ada antrean baru"; jumlah = 0; tertahan = @(); dikirim_ke = @(); pesan = "" }
}

# --- penjaga kebocoran ----------------------------------------------------
foreach ($p in $pesanKirim) {
    foreach ($sid in $p.ids) {
        $b = $semua | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq "$sid".ToUpper() } | Select-Object -First 1
        if ($b -and -not (Periksa-Bersih $p.teks $b)) {
            Keluar-Json @{ ok = $false; error = "BOCORAN_DITOLAK"; id = $sid; pesan = "Notifikasi memuat data yang dilarang (NIK/alamat/isi surat) - tidak dikirim." }
        }
    }
}

if ($DryRun) {
    $logUji = Join-Path $AntreanDir "notifikasi_antrean.txt"
    foreach ($p in $pesanKirim) {
        $catatLog = "=== {0} | RENCANA (belum dikirim) | {1} ===`n{2}`n`n" -f $waktu.ToString("yyyy-MM-dd HH\:mm"), (($p.ids) -join ","), $p.teks
        [IO.File]::AppendAllText($logUji, $catatLog, (New-Object Text.UTF8Encoding($false)))
    }
    Keluar-Json @{
        ok        = $true
        dry_run   = $true
        mode      = "siap kirim"
        jumlah    = $pesanKirim.Count
        tertahan  = @($tertahan | ForEach-Object { "$($_.id_surat)".Trim() })
        pesan     = ($pesanKirim | ForEach-Object { $_.teks }) -join "`n---`n"
    }
}

# --- penerima: semua petugas aktif ----------------------------------------
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
foreach ($p in $pesanKirim) {
    $suksesSatu = $false
    foreach ($nomor in $penerima) {
        $hasil = & $hermes --profile surat-desa send -t "whatsapp:$nomor" --json $p.teks 2>&1
        $teksHasil = ($hasil | Out-String)
        if ($teksHasil -match '"success"\s*:\s*true') { $suksesSatu = $true; $terkirimKe += (Samarkan-Nomor $nomor) }
        else { $gagal += (Samarkan-Nomor $nomor) }
    }
    # CATATAN: ekspresi `if` TIDAK boleh dipakai sebagai argumen fungsi di PS 5.1
    # (jalan di hashtable, gagal di posisi argumen) — hitung dulu ke variabel.
    $statusCatat = "gagal"
    if ($suksesSatu) { $statusCatat = "terkirim" }
    $tujuanTeks = ($penerima | ForEach-Object { Samarkan-Nomor $_ }) -join "; "
    foreach ($sid in $p.ids) {
        $b = $semua | Where-Object { ("$($_.id_surat)".Trim().ToUpper()) -eq "$sid".ToUpper() } | Select-Object -First 1
        $masuk = ""
        if ($b) { $masuk = "$($b.tanggal_pengajuan)".Trim() }
        Tulis-Catat "$sid".ToUpper() $masuk $statusCatat $tujuanTeks
    }
    # jejak audit: simpan pesan yang benar-benar dikirim (UTF-8, tanpa BOM)
    $infoKirim = "GAGAL"
    if ($suksesSatu) { $infoKirim = "TERKIRIM ke " + (($penerima | ForEach-Object { Samarkan-Nomor $_ }) -join ", ") }
    $logKirim = Join-Path $AntreanDir "notifikasi_antrean.txt"
    $catatLog = "=== {0} | {1} | {2} ===`n{3}`n`n" -f $waktu.ToString("yyyy-MM-dd HH\:mm"), $infoKirim, (($p.ids) -join ","), $p.teks
    [IO.File]::AppendAllText($logKirim, $catatLog, (New-Object Text.UTF8Encoding($false)))
}
$catat | Sort-Object id_surat | Export-Csv -Path $StateFile -NoTypeInformation -Encoding UTF8

Keluar-Json @{
    ok          = $true
    mode        = if ($tertahan.Count -gt 0) { "lanjutan jam tenang (sekali-sekaligus)" } else { "langsung" }
    jumlah      = $pesanKirim.Count
    tertahan    = @($tertahan | ForEach-Object { "$($_.id_surat)".Trim() })
    dikirim_ke  = @($terkirimKe | Select-Object -Unique)
    gagal       = @($gagal | Select-Object -Unique)
    pesan       = (($pesanKirim | ForEach-Object { $_.teks }) -join "`n---`n")
}
