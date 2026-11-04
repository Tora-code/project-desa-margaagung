# ============================================================================
# dashboard_admin.ps1 - Ringkasan untuk menu admin di WhatsApp.
#
# Bagian:
#   rekap  -> jumlah surat per status + daftar ID yang menunggu
#   ambil  -> isi antrean pengambilan (surat disetujui, belum/sudah dikabari)
#
# Pakai:
#   .\dashboard_admin.ps1 -Bagian rekap
#   .\dashboard_admin.ps1 -Bagian ambil
# ============================================================================

param(
    [Parameter(Mandatory = $true)][ValidateSet("rekap", "ambil")][string]$Bagian,
    [string]$AntreanDir = ""
)

$ErrorActionPreference = "Stop"

if (-not $AntreanDir) {
    $AntreanDir = Join-Path (Split-Path -Parent $PSScriptRoot) "Antrean-surat"
}
$csv = Join-Path $AntreanDir "antrean_surat.csv"
$csvAmbil = Join-Path $AntreanDir "antrean_ambil.csv"

function Keluar-Json($obj) {
    $obj | ConvertTo-Json -Depth 5 -Compress
    exit 0
}

if ($Bagian -eq "rekap") {
    if (-not (Test-Path $csv)) {
        Keluar-Json @{ ok = $false; error = "DATA_ANTREAN_TIDAK_DITEMUKAN" }
    }
    $semua = @(Import-Csv -Path $csv -Encoding UTF8)

    # Kosakata baru (bahasa biasa) + kode lama, supaya baris lama tetap terhitung.
    $menunggu = @($semua | Where-Object { "$($_.status)".Trim().ToUpper() -like "MENUNGGU*" })
    $disetujui = @($semua | Where-Object { "$($_.status)".Trim().ToUpper() -in @("DISETUJUI", "DISETUJUI_ADMIN", "SUDAH DISETUJUI") })
    $siap = @($semua | Where-Object { "$($_.status)".Trim().ToUpper() -in @("SIAP_DIAMBIL", "SIAP DIAMBIL") })
    $diambil = @($semua | Where-Object { "$($_.status)".Trim().ToUpper() -in @("DIAMBIL", "SUDAH DIAMBIL") })

    $hariIni = (Get-Date).ToString("yyyy-MM-dd")
    $baruHariIni = @($semua | Where-Object { "$($_.tanggal_pengajuan)".StartsWith($hariIni) })

    Keluar-Json @{
        ok           = $true
        hari_ini     = $baruHariIni.Count
        menunggu     = $menunggu.Count
        disetujui    = $disetujui.Count
        siap_diambil = $siap.Count
        sudah_diambil = $diambil.Count
        total        = $semua.Count
        daftar_menunggu = @($menunggu | ForEach-Object {
                @{ id = "$($_.id_surat)".Trim(); nama = "$($_.nama)".Trim(); jenis = "$($_.jenis_surat)".Trim() }
            })
        daftar_siap = @($siap | ForEach-Object {
                @{ id = "$($_.id_surat)".Trim(); nama = "$($_.nama)".Trim(); jenis = "$($_.jenis_surat)".Trim() }
            })
    }
}

# --- bagian: ambil --------------------------------------------------------
if (-not (Test-Path $csvAmbil)) {
    Keluar-Json @{ ok = $true; jumlah = 0; items = @() }
}

$ambil = @(Import-Csv -Path $csvAmbil -Encoding UTF8)
$items = @($ambil | ForEach-Object {
        @{
            id            = "$($_.id_surat)".Trim()
            nama          = "$($_.nama)".Trim()
            tanggal_siap  = "$($_.tanggal_siap)".Trim()
            dikabari_pada = "$($_.dikabari_pada)".Trim()
            catatan       = "$($_.catatan)".Trim()
        }
    })

Keluar-Json @{
    ok     = $true
    jumlah = $items.Count
    belum_dikabari = @($items | Where-Object { -not $_.dikabari_pada }).Count
    items  = $items
}
