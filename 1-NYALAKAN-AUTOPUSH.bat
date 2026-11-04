@echo off
REM ============================================================
REM  Nyalakan watcher auto-push ke GitHub (tanpa jendela hitam)
REM  Dobel-klik file ini sekali. Watcher jalan di belakang layar.
REM ============================================================
cd /d "%~dp0"
start "" /min "C:\Python314\pythonw.exe" "_watcher_push.py"
echo.
echo   Watcher auto-push SUDAH AKTIF.
echo   Setiap perubahan file akan otomatis di-push ke GitHub.
echo.
echo   Cek log aktivitas :  lihat file  _watcher_push.log
echo   Matikan           :  Task Manager - cari "pythonw.exe" - End task
echo.
timeout /t 8 >nul
