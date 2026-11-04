@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================
echo  Membangun ulang database_kk.csv dari file .xls
echo ================================================
echo.

set "PY="
if exist "C:\Python314\python.exe" set "PY=C:\Python314\python.exe"
if not defined PY if exist "C:\Python313\python.exe" set "PY=C:\Python313\python.exe"
if not defined PY if exist "C:\Python312\python.exe" set "PY=C:\Python312\python.exe"
if not defined PY where py >nul 2>nul && set "PY=py"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY goto nopython

"%PY%" -c "import xlrd" >nul 2>nul
if errorlevel 1 (
  echo Memasang pustaka pembaca Excel ^(xlrd^) sekali saja...
  "%PY%" -m pip install --quiet xlrd
)

"%PY%" "%~dp0bangun_database.py"
if errorlevel 1 (
  echo.
  echo GAGAL. Catat pesan error di atas.
  pause
  exit /b 1
)

echo.
echo Selesai. database_kk.csv sudah diperbarui.
echo Coba cek:  cari_warga.ps1 -nik "NIK 16 digit" -nama "NAMA LENGKAP"
pause
exit /b 0

:nopython
echo Python tidak ditemukan di laptop ini.
echo Pasang Python dari https://python.org (centang "Add python to PATH"), lalu jalankan file ini lagi.
pause
exit /b 1
