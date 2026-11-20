# AlifboOtkazgich Setup Builder
# Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
# Ishlatish: powershell -File make_setup.ps1

$ErrorActionPreference = 'Stop'
$src = $PSScriptRoot; if (-not $src) { $src = Split-Path -Parent $MyInvocation.MyCommand.Path }
$dist = "$src\dist"
$nsis = "C:\Program Files (x86)\NSIS\makensis.exe"

Write-Host "=== AlifboOtkazgich Setup Builder v1.0.4 ===" -ForegroundColor Cyan
Write-Host ""

# dist papkasini tozalash va yaratish
if (Test-Path $dist) { Remove-Item $dist -Recurse -Force }
New-Item -ItemType Directory -Path $dist | Out-Null
New-Item -ItemType Directory -Path "$dist\tools" | Out-Null
New-Item -ItemType Directory -Path "$dist\dotm" | Out-Null
New-Item -ItemType Directory -Path "$dist\xlam" | Out-Null
New-Item -ItemType Directory -Path "$dist\ppam" | Out-Null

Write-Host "[1/5] Fayllar ko'chirilmoqda..." -ForegroundColor Yellow

# Asosiy fayllar
Copy-Item "$src\alifbo-yordamchi.exe" "$dist\alifbo-yordamchi.exe" -Force
Copy-Item "$src\alifbo.ico" "$dist\alifbo.ico" -Force
Copy-Item "$src\version.json" "$dist\version.json" -Force

# Tools
Copy-Item "$src\tools\alifbo-bolish.ps1" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-lic.ps1" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-tools.ps1" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-core.cs" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-pdf2docx.ps1" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-sssr.ps1" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-sssr-ocr.py" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-sssr-gui.ps1" "$dist\tools\" -Force
Copy-Item "$src\tools\alifbo-pdf2word-gui.ps1" "$dist\tools\" -Force

# Office add-inlar
$dotm = "$env:APPDATA\Microsoft\Word\STARTUP\AlifboOtkazgich.dotm"
$xlam = "$env:APPDATA\Microsoft\Excel\XLSTART\AlifboOtkazgich.xlam"
$ppam = "$env:APPDATA\Microsoft\AddIns\AlifboOtkazgich.ppam"

if (-not (Test-Path $dotm)) { Write-Error "AlifboOtkazgich.dotm topilmadi: $dotm"; exit 1 }
if (-not (Test-Path $xlam)) { Write-Error "AlifboOtkazgich.xlam topilmadi: $xlam"; exit 1 }
if (-not (Test-Path $ppam)) { New-Item -ItemType File -Path $ppam -Force | Out-Null }

Copy-Item $dotm "$dist\dotm\AlifboOtkazgich.dotm" -Force
Copy-Item $xlam "$dist\xlam\AlifboOtkazgich.xlam" -Force
Copy-Item $ppam "$dist\ppam\AlifboOtkazgich.ppam" -Force

Write-Host "  dotm: $([math]::Round((Get-Item $dotm).Length/1024,1)) KB" -ForegroundColor Gray
Write-Host "  xlam: $([math]::Round((Get-Item $xlam).Length/1024,1)) KB" -ForegroundColor Gray
Write-Host "  ppam: $([math]::Round((Get-Item $ppam).Length/1024,1)) KB" -ForegroundColor Gray

Write-Host "[2/5] NSIS skripti yaratilmoqda..." -ForegroundColor Yellow

$nsiContent = @'
; AlifboOtkazgich Setup
; Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473

!define APPNAME "AlifboOtkazgich"
!define APPVERSION "1.0.4"
!define PUBLISHER "Fozliddinov Safarmurod"
!define APPDIR "$LOCALAPPDATA\AlifboOtkazgich"

Name "${APPNAME} ${APPVERSION}"
OutFile "AlifboOtkazgich-Setup.exe"
InstallDir "${APPDIR}"
RequestExecutionLevel user
SetCompressor /SOLID lzma
Unicode True

!include "MUI2.nsh"

!define MUI_ICON "alifbo.ico"
!define MUI_UNICON "alifbo.ico"
!define MUI_WELCOMEPAGE_TITLE "AlifboOtkazgich ${APPVERSION} o'rnatish"
!define MUI_WELCOMEPAGE_TEXT "Ushbu dastur yangi o'zbek alifbosiga matnlarni o'tkazish uchun Word, Excel va PowerPoint qo'shimchalarini o'rnatadi.$\n$\nMuallif: Fozliddinov Safarmurod$\nTelegram: @Safar_master$\nTel: +998939530473"
!define MUI_FINISHPAGE_TITLE "O'rnatish tugadi!"
!define MUI_FINISHPAGE_TEXT "AlifboOtkazgich ${APPVERSION} muvaffaqiyatli o'rnatildi.$\n$\nWord, Excel yoki PowerPoint ni qayta oching - Alifbo paneli paydo bo'ladi."
!define MUI_FINISHPAGE_RUN "$LOCALAPPDATA\AlifboOtkazgich\alifbo-yordamchi.exe"
!define MUI_FINISHPAGE_RUN_TEXT "Yordamchi dasturni hozir ishga tushirish"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Russian"

Section "Asosiy"
  ; Ishlab turgan yordamchini to'xtatamiz (fayl qulflanib qolmasligi uchun)
  DetailPrint "Eski jarayonlar tekshirilmoqda..."
  ExecWait 'taskkill /F /IM alifbo-yordamchi.exe'
  Sleep 1000

  SetOutPath "${APPDIR}"
  File "alifbo-yordamchi.exe"
  File "alifbo.ico"
  File "version.json"

  SetOutPath "${APPDIR}\tools"
  File "tools\alifbo-bolish.ps1"
  File "tools\alifbo-lic.ps1"
  File "tools\alifbo-tools.ps1"
  File "tools\alifbo-core.cs"
  File "tools\alifbo-pdf2docx.ps1"
  File "tools\alifbo-sssr.ps1"
  File "tools\alifbo-sssr-ocr.py"
  File "tools\alifbo-sssr-gui.ps1"
  File "tools\alifbo-pdf2word-gui.ps1"

  SetOutPath "${APPDIR}"
  WriteUninstaller "${APPDIR}\Uninstall.exe"

  SetOutPath "$APPDATA\Microsoft\Word\STARTUP"
  File /oname=AlifboOtkazgich.dotm "dotm\AlifboOtkazgich.dotm"

  SetOutPath "$APPDATA\Microsoft\Excel\XLSTART"
  File /oname=AlifboOtkazgich.xlam "xlam\AlifboOtkazgich.xlam"

  SetOutPath "$APPDATA\Microsoft\AddIns"
  File /oname=AlifboOtkazgich.ppam "ppam\AlifboOtkazgich.ppam"

  ; PowerPoint Add-in avtomatik yuklanishi uchun registr yozuvlari
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\PowerPoint\AddIns\AlifboOtkazgich" "Path" "$APPDATA\Microsoft\AddIns\AlifboOtkazgich.ppam"
  WriteRegDWORD HKCU "Software\Microsoft\Office\16.0\PowerPoint\AddIns\AlifboOtkazgich" "AutoLoad" 1
  WriteRegStr HKCU "Software\Microsoft\Office\15.0\PowerPoint\AddIns\AlifboOtkazgich" "Path" "$APPDATA\Microsoft\AddIns\AlifboOtkazgich.ppam"
  WriteRegDWORD HKCU "Software\Microsoft\Office\15.0\PowerPoint\AddIns\AlifboOtkazgich" "AutoLoad" 1

  ; Word, Excel va PowerPoint bloklangan (DisabledItems) ro'yxatini tozalash
  DeleteRegKey HKCU "Software\Microsoft\Office\16.0\Word\Resiliency\DisabledItems"
  DeleteRegKey HKCU "Software\Microsoft\Office\15.0\Word\Resiliency\DisabledItems"
  DeleteRegKey HKCU "Software\Microsoft\Office\14.0\Word\Resiliency\DisabledItems"
  DeleteRegKey HKCU "Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems"
  DeleteRegKey HKCU "Software\Microsoft\Office\16.0\PowerPoint\Resiliency\DisabledItems"

  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "AlifboYordamchi" '"${APPDIR}\alifbo-yordamchi.exe"'

  ; Kuchsiz (4GB RAM) kompyuterlarda klaviatura uzilib qolmasligi uchun Windows kutilma vaqtini 2000 ms ga oshiramiz
  WriteRegStr HKCU "Control Panel\Desktop" "LowLevelHooksTimeout" "2000"



  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME} ${APPVERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" '"${APPDIR}\Uninstall.exe"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" "${APPDIR}\alifbo.ico"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "${PUBLISHER}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${APPVERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoModify" 1
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "NoRepair" 1

  Exec '"${APPDIR}\alifbo-yordamchi.exe"'
SectionEnd

Section "Uninstall"
  ExecWait 'taskkill /F /IM alifbo-yordamchi.exe'
  Sleep 1000

  DeleteRegValue HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "AlifboYordamchi"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
  DeleteRegValue HKCU "Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableInternetFilesInPV"
  DeleteRegValue HKCU "Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableUnsafeLocationsInPV"
  DeleteRegValue HKCU "Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableAttachmentsInPV"
  DeleteRegKey HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDownloads"
  DeleteRegKey HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDesktop"
  Delete "$APPDATA\Microsoft\Word\STARTUP\AlifboOtkazgich.dotm"
  Delete "$APPDATA\Microsoft\Excel\XLSTART\AlifboOtkazgich.xlam"
  Delete "$APPDATA\Microsoft\AddIns\AlifboOtkazgich.ppam"
  Delete "${APPDIR}\alifbo-yordamchi.exe"
  Delete "${APPDIR}\alifbo.ico"
  Delete "${APPDIR}\version.json"
  Delete "${APPDIR}\tools\*.*"
  RMDir "${APPDIR}\tools"
  Delete "${APPDIR}\Uninstall.exe"
  RMDir "${APPDIR}"
SectionEnd
'@

$nsiContent | Set-Content "$dist\setup.nsi" -Encoding UTF8

Write-Host "[3/5] NSIS compile qilinmoqda..." -ForegroundColor Yellow
Push-Location $dist
& $nsis "setup.nsi"
$exitCode = $LASTEXITCODE
Pop-Location
if (Test-Path "$dist\AlifboOtkazgich-Setup.exe") {
    Move-Item "$dist\AlifboOtkazgich-Setup.exe" "$src\AlifboOtkazgich-Setup.exe" -Force
}

if ($exitCode -ne 0) {
    Write-Host "XATO: NSIS compile xato! Kod: $exitCode" -ForegroundColor Red
    exit 1
}

$setupPath = "$src\AlifboOtkazgich-Setup.exe"
if (Test-Path $setupPath) {
    $size = [math]::Round((Get-Item $setupPath).Length / 1MB, 2)
    Write-Host "[4/5] Setup.exe yaratildi: $size MB" -ForegroundColor Green
    Write-Host "      $setupPath" -ForegroundColor Gray
} else {
    Write-Host "XATO: Setup.exe yaratilmadi!" -ForegroundColor Red
    exit 1
}

Write-Host "[5/5] dist papkasi tozalanmoqda..." -ForegroundColor Yellow
Remove-Item $dist -Recurse -Force

Write-Host ""
Write-Host "=== TAYYOR! ===" -ForegroundColor Green
Write-Host "Setup fayil: C:\Users\User\Downloads\AlifboOtkazgich-Setup.exe ($size MB)" -ForegroundColor Cyan
