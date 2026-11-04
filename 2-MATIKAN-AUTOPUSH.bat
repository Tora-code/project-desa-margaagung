@echo off
REM ============================================================
REM  Matikan watcher auto-push
REM ============================================================
echo Menghentikan watcher auto-push...
taskkill /f /im pythonw.exe 2>nul
if %errorlevel%==0 (
  echo   Watcher DIHENTIKAN.
) else (
  echo   Tidak ada watcher yang berjalan.
)
timeout /t 5 >nul
