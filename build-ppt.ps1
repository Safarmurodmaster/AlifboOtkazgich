# build-ppt.ps1 - PowerPoint PPAM builder (PowerShell)
# Muallif: Fozliddinov Safarmurod | @Safar_master

$ErrorActionPreference = 'Continue'
$logDir = "$env:APPDATA\Alifbo"
$logPath = "$logDir\kurulish-log.txt"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }

function Lg {
    param([string]$msg)
    $line = "$(Get-Date -Format 'dd.MM.yyyy H:mm:ss')  [PPT-PS] $msg"
    Add-Content -Path $logPath -Value $line -Encoding UTF8
    Write-Host $line
}

$dir = $PSScriptRoot
if (-not $dir) { $dir = Split-Path -Parent $MyInvocation.MyCommand.Path }

# PowerPoint exe qidirish
$pptExe = ""
$candidates = @(
    "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
    "C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
    "C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE"
)
foreach ($c in $candidates) {
    if (Test-Path $c) { $pptExe = $c; break }
}

if ($pptExe -eq "") {
    Lg "POWERPNT.EXE topilmadi - o'tkazib yuborildi"
    exit 2
}
Lg "PowerPoint topildi: $pptExe"

# Ishlab turgan jarayonni tekshirish
$running = Get-Process -Name POWERPNT -ErrorAction SilentlyContinue
if ($running) {
    Lg "POWERPNT ishlab turibdi - o'tkazib yuborildi"
    exit 7
}

# AccessVBOM yoqish
$secKey = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security"
$oldVal = $null
try {
    $oldVal = (Get-ItemProperty $secKey -Name AccessVBOM -ErrorAction Stop).AccessVBOM
} catch { }
Set-ItemProperty $secKey -Name AccessVBOM -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
Lg "AccessVBOM yoqildi"

$addinDir = "$env:APPDATA\Microsoft\AddIns"
if (-not (Test-Path $addinDir)) { New-Item -ItemType Directory -Path $addinDir | Out-Null }

# Trusted Locations
$tlKey = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\Location90"
New-Item $tlKey -Force | Out-Null
Set-ItemProperty $tlKey -Name Path -Value "$addinDir\" -Force
Set-ItemProperty $tlKey -Name AllowSubfolders -Value 1 -Type DWord -Force
Set-ItemProperty $tlKey -Name Description -Value "AlifboOtkazgich" -Force
Lg "Ishonchli joy qoshildi"

$target = "$addinDir\AlifboOtkazgich.ppam"
$code = 0
$ppt = $null

try {
    Lg "PowerPoint COM yaratilmoqda..."
    $ppt = New-Object -ComObject "PowerPoint.Application"
    if ($null -eq $ppt) { throw "COM null qaytdi" }
    Lg "PowerPoint COM ok"

    # Yangi taqdimot (Visible = msoTrue = 1)
    Lg "Yangi taqdimot ochilmoqda..."
    $pres = $ppt.Presentations.Add(1)
    Lg "Yangi taqdimot: ok"

    # Blank slide (ppLayoutBlank = 12)
    $null = $pres.Slides.Add(1, 12)
    Lg "Blank slide qoshildi: ok"

    # Modullar import
    $modules = @("AlifboCore", "AlifboLic", "AlifboKrill", "AlifboQk", "AlifboPpt")
    foreach ($mod in $modules) {
        $modPath = "$dir\$mod.bas"
        if (Test-Path $modPath) {
            $null = $pres.VBProject.VBComponents.Import($modPath)
            Lg "$mod import: ok"
        } else {
            Lg "$mod.bas TOPILMADI"
        }
    }

    $cnt = $pres.VBProject.VBComponents.Count
    Lg "Modullar soni: $cnt"
    if ($cnt -lt 5) { $code = 4 }

    if ($code -eq 0) {
        if (Test-Path $target) { Remove-Item $target -Force -ErrorAction SilentlyContinue }

        Lg "ppam saqlanmoqda: $target"
        try {
            $pres.SaveAs($target, 30)
            Lg "ppam saqlash: ok"
        } catch {
            Lg "ppam saqlash xato: $($_.Exception.Message)"
            # Fallback: pptm
            $tmpPath = "$addinDir\AlifboOtkazgich_tmp.pptm"
            try {
                $pres.SaveAs($tmpPath, 25)
                Lg "pptm saqlash: ok"
                Copy-Item $tmpPath $target -Force
                Remove-Item $tmpPath -Force -ErrorAction SilentlyContinue
                Lg "pptm->ppam: ok"
            } catch {
                Lg "pptm ham muvaffaqiyatsiz: $($_.Exception.Message)"
                $code = 5
            }
        }

        $pres.Close()
    }

    $ppt.Quit()

} catch {
    Lg "COM xato: $($_.Exception.Message)"
    $code = 3
    if ($null -ne $ppt) { try { $ppt.Quit() } catch { } }
}

# COM obektni ozod qilish
if ($null -ne $ppt) {
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch { }
    $ppt = $null
}
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

# AccessVBOM qayta tiklash
try {
    if ($null -ne $oldVal) {
        Set-ItemProperty $secKey -Name AccessVBOM -Value $oldVal -Type DWord -Force
    } else {
        Remove-ItemProperty $secKey -Name AccessVBOM -Force -ErrorAction SilentlyContinue
    }
} catch { }

# ppam yaratildi va hajmi >1KB bo'lsa ribbon yamash va registry
if ($code -eq 0) {
    $sz = 0
    if (Test-Path $target) { $sz = (Get-Item $target).Length }
    if ($sz -gt 1000) {
        Lg "Ribbon yamash boshlandi..."
        $rc = Start-Process powershell.exe -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$dir\patch-ribbon.ps1`"", "-Path", "`"$target`"", "-XmlPath", "`"$dir\ribbon-ppt.xml`"" -Wait -PassThru -WindowStyle Hidden
        Lg "Ribbon yamash kod: $($rc.ExitCode)"
        if ($rc.ExitCode -ne 0) { $code = 6 }

        $rk = "HKCU:\Software\Microsoft\Office\16.0\PowerPoint\AddIns\AlifboOtkazgich"
        New-Item $rk -Force | Out-Null
        Set-ItemProperty $rk -Name Path -Value $target -Force
        Set-ItemProperty $rk -Name AutoLoad -Value 1 -Type DWord -Force
        Lg "Registry AutoLoad: ok"
    } else {
        Lg "ppam hajmi noto'g'ri: $sz bytes"
        $code = 5
    }
}

$finalSz = if (Test-Path $target) { (Get-Item $target).Length } else { 0 }
Lg "Yakuniy kod: $code | ppam hajmi: $finalSz bytes"
exit $code
