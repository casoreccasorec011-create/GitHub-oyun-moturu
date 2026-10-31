@echo off
setlocal

set "FILE_NAME=AI_x_3D_Oyun_Motoru_Master_Plan.xlsx"
set "DOWNLOAD_URL=https://raw.githubusercontent.com/casoreccasorec011-create/GitHub-oyun-moturu/main/%FILE_NAME%"
set "DOWNLOAD_PATH=%USERPROFILE%\Downloads\%FILE_NAME%"

if not exist "%USERPROFILE%\Downloads" mkdir "%USERPROFILE%\Downloads"

echo Excel dosyasi indiriliyor...
powershell.exe -NoProfile -Command "$ErrorActionPreference='Stop'; try { Invoke-WebRequest -UseBasicParsing -Uri $env:DOWNLOAD_URL -OutFile $env:DOWNLOAD_PATH } catch { Write-Error $_; exit 1 }"
if errorlevel 1 (
    echo Indirme basarisiz oldu. Internet baglantinizi kontrol edip tekrar deneyin.
    pause
    exit /b 1
)

echo Excel dosyasi indirildi: %DOWNLOAD_PATH%
start "" "%DOWNLOAD_PATH%"
