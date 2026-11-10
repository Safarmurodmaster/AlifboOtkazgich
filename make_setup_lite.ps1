# AlifboOtkazgich Lite Setup Builder
# Kuchsiz va 4 GB RAM kompyuterlar uchun maxsus yengil versiya
# Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473

$ErrorActionPreference = 'Stop'
$src = "C:\Users\User\Downloads\AlifboOtkazgich-kodlar"
$dist = "$src\dist_lite"
$nsis = "C:\Program Files (x86)\NSIS\makensis.exe"

Write-Host "=== AlifboOtkazgich Lite Setup Builder v1.0.3 ===" -ForegroundColor Cyan
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

# Office add-inlar
$dotm = "$env:APPDATA\Microsoft\Word\STARTUP\AlifboOtkazgich.dotm"
$xlam = "$env:APPDATA\Microsoft\Excel\XLSTART\AlifboOtkazgich.xlam"
$ppam = "$env:APPDATA\Microsoft\AddIns\AlifboOtkazgich.ppam"

if (-not (Test-Path $dotm)) { Write-Error "AlifboOtkazgich.dotm topilmadi: $dotm"; exit 1 }
if (-not (Test-Path $xlam)) { Write-Error "AlifboOtkazgich.xlam topilmadi: $xlam"; exit 1 }
if (-not (Test-Path $ppam)) { Write-Error "AlifboOtkazgich.ppam topilmadi: $ppam"; exit 1 }

Copy-Item $dotm "$dist\dotm\AlifboOtkazgich.dotm" -Force
Copy-Item $xlam "$dist\xlam\AlifboOtkazgich.xlam" -Force
Copy-Item $ppam "$dist\ppam\AlifboOtkazgich.ppam" -Force

Write-Host "  dotm: $([math]::Round((Get-Item $dotm).Length/1024,1)) KB" -ForegroundColor Gray
Write-Host "  xlam: $([math]::Round((Get-Item $xlam).Length/1024,1)) KB" -ForegroundColor Gray
Write-Host "  ppam: $([math]::Round((Get-Item $ppam).Length/1024,1)) KB" -ForegroundColor Gray

Write-Host "[2/5] NSIS skripti yaratilmoqda..." -ForegroundColor Yellow

$nsiContent = @'
; AlifboOtkazgich Lite Setup (Kuchsiz kompyuterlar uchun)
; Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473

!define APPNAME "AlifboOtkazgich Lite"
!define APPVERSION "1.0.3"
!define PUBLISHER "Fozliddinov Safarmurod"
!define APPDIR "$LOCALAPPDATA\AlifboOtkazgich"

Name "${APPNAME} ${APPVERSION}"
OutFile "C:\Users\User\Downloads\AlifboOtkazgich-Lite-Setup.exe"
InstallDir "${APPDIR}"
RequestExecutionLevel user
SetCompressor /SOLID lzma
Unicode True

!include "MUI2.nsh"

!define MUI_ICON "alifbo.ico"
!define MUI_UNICON "alifbo.ico"
!define MUI_WELCOMEPAGE_TITLE "AlifboOtkazgich Lite ${APPVERSION} (Yengil versiya)"
!define MUI_WELCOMEPAGE_TEXT "Ushbu dastur yangi o'zbek alifbosiga matnlarni o'tkazish uchun Word, Excel va PowerPoint qo'shimchalarini o'rnatadi.$\n$\nBu versiya operativ xotirasi 2GB-4GB bo'lgan kuchsiz kompyuterlar uchun maxsus optimallashtirilgan.$\n$\nMuallif: Fozliddinov Safarmurod$\nTelegram: @Safar_master$\nTel: +998939530473"
!define MUI_FINISHPAGE_TITLE "O'rnatish tugadi!"
!define MUI_FINISHPAGE_TEXT "AlifboOtkazgich Lite ${APPVERSION} muvaffaqiyatli o'rnatildi.$\n$\nWord, Excel yoki PowerPoint ni oching - Alifbo paneli paydo bo'ladi."
!define MUI_FINISHPAGE_RUN "$LOCALAPPDATA\AlifboOtkazgich\alifbo-yordamchi.exe"
!define MUI_FINISHPAGE_RUN_TEXT "Yordamchi dasturni hozir ishga tushirish"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Russian"

Section "Asosiy"
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

  SetOutPath "${APPDIR}"
  WriteUninstaller "${APPDIR}\Uninstall.exe"

  SetOutPath "$APPDATA\Microsoft\Word\STARTUP"
  File /oname=AlifboOtkazgich.dotm "dotm\AlifboOtkazgich.dotm"

  SetOutPath "$APPDATA\Microsoft\Excel\XLSTART"
  File /oname=AlifboOtkazgich.xlam "xlam\AlifboOtkazgich.xlam"

  SetOutPath "$APPDATA\Microsoft\AddIns"
  File /oname=AlifboOtkazgich.ppam "ppam\AlifboOtkazgich.ppam"

  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "AlifboYordamchi" '"${APPDIR}\alifbo-yordamchi.exe"'

  ; Kuchsiz (4GB RAM) kompyuterlarda klaviatura uzilib qolmasligi uchun Windows kutilma vaqtini 2000 ms ga oshiramiz
  DetailPrint "Klaviatura kutilma vaqti (LowLevelHooksTimeout) sozlanmoqda..."
  WriteRegStr HKCU "Control Panel\Desktop" "LowLevelHooksTimeout" "2000"

  ; Word Protected View o'chirish
  DetailPrint "Word Protected View sozlanmoqda..."
  WriteRegDWORD HKCU "Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableInternetFilesInPV" 1
  WriteRegDWORD HKCU "Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableUnsafeLocationsInPV" 1
  WriteRegDWORD HKCU "Software\Microsoft\Office\16.0\Word\Security\ProtectedView" "DisableAttachmentsInPV" 1

  ; Downloads va Desktop papkalarini ishonchli joylarga qo'shish
  ExpandEnvStrings $R0 "%USERPROFILE%"
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDownloads" "Path" "$R0\Downloads"
  WriteRegDWORD HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDownloads" "AllowSubFolders" 1
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDownloads" "Description" "AlifboOtkazgich - Downloads"
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDesktop" "Path" "$R0\Desktop"
  WriteRegDWORD HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDesktop" "AllowSubFolders" 1
  WriteRegStr HKCU "Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\AlifboDesktop" "Description" "AlifboOtkazgich - Desktop"

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

$nsiContent | Set-Content "$dist\setup_lite.nsi" -Encoding UTF8

Write-Host "[3/5] NSIS compile qilinmoqda..." -ForegroundColor Yellow
Push-Location $dist
& $nsis "setup_lite.nsi"
$exitCode = $LASTEXITCODE
Pop-Location

if ($exitCode -ne 0) {
    Write-Host "XATO: NSIS compile xato! Kod: $exitCode" -ForegroundColor Red
    exit 1
}

$setupPath = "C:\Users\User\Downloads\AlifboOtkazgich-Lite-Setup.exe"
if (Test-Path $setupPath) {
    $size = [math]::Round((Get-Item $setupPath).Length / 1MB, 2)
    Write-Host "[4/5] Setup.exe yaratildi: $size MB" -ForegroundColor Green
    Write-Host "      $setupPath" -ForegroundColor Gray
} else {
    Write-Host "XATO: Setup.exe yaratilmadi!" -ForegroundColor Red
    exit 1
}

Write-Host "[5/5] dist_lite papkasi tozalanmoqda..." -ForegroundColor Yellow
Remove-Item $dist -Recurse -Force

Write-Host ""
Write-Host "=== TAYYOR! ===" -ForegroundColor Green
Write-Host "Lite Setup fayil: C:\Users\User\Downloads\AlifboOtkazgich-Lite-Setup.exe ($size MB)" -ForegroundColor Cyan
