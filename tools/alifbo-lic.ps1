# AlifboOtkazgich: umumiy qismlar (litsenziya, yadroni yuklash)
# Muallif: Fozliddinov Safarmurod, Telegram @Safar_master, +998939530473
$script:FREE = $false
$script:SALT = "SAFAR_2026_"
$script:REGP = 'HKCU:\Software\Microsoft\AlifboSetup'
$script:DATD = Join-Path $env:APPDATA 'Alifbo'

function Initialize-AlifboCore([string]$dir) {
    if ('AlifboCore' -as [type]) { return }
    $src = Get-Content (Join-Path $dir 'alifbo-core.cs') -Raw -Encoding UTF8
    if ($PSVersionTable.PSVersion.Major -ge 6) { Add-Type -TypeDefinition $src }
    else { Add-Type -AssemblyName System.IO.Compression; Add-Type -TypeDefinition $src -ReferencedAssemblies System.IO.Compression }
}

function Get-ExcSet {
    return [AlifboCore]::LoadExc((Join-Path $script:DATD 'istisnolar.txt'))
}

function Hash16([string]$s) {
    [int]$h = 17
    for ($i = 0; $i -lt $s.Length; $i++) {
        $c = [int][char]$s[$i]
        $h = (($h * 31 + $c) -band 32767)
    }
    return $h
}

function Get-MachineCode {
    try {
        $fso = New-Object -ComObject Scripting.FileSystemObject
        $drv = $fso.GetDrive("C:")
        $sn = [string][Math]::Abs([long]$drv.SerialNumber)
    } catch {
        $sn = "12345678"
    }
    $h1 = "{0:X4}" -f [int](Hash16 ("M1" + $sn))
    $h2 = "{0:X4}" -f [int](Hash16 ("M2" + $sn))
    return "AF-$h1-$h2"
}

function Calculate-Key([string]$machCode) {
    $clean = $machCode.Replace("-", "").Replace(" ", "").ToUpper()
    $s = $script:SALT + $clean
    $k1 = "{0:X4}" -f [int](Hash16 ("K1" + $s))
    $k2 = "{0:X4}" -f [int](Hash16 ("K2" + $s))
    return "AK-$k1-$k2"
}

function Get-SavedKey {
    $saved = ""
    try {
        $saved = (Get-ItemProperty -Path $script:REGP -Name 'lic_key' -ErrorAction Stop).lic_key
    } catch { }
    if ([string]::IsNullOrWhiteSpace($saved)) {
        try {
            $f = Join-Path $script:DATD 'lic.dat'
            if (Test-Path $f -ErrorAction Stop) {
                $saved = (Get-Content $f -TotalCount 1 -ErrorAction Stop).Trim()
            }
        } catch { }
    }
    return [string]$saved
}

function Set-SavedKey([string]$key) {
    try {
        if (-not (Test-Path $script:REGP -ErrorAction Stop)) {
            New-Item -Path $script:REGP -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $script:REGP -Name 'lic_key' -Value $key -ErrorAction Stop
    } catch { }
    try {
        if (-not (Test-Path $script:DATD)) {
            New-Item -ItemType Directory -Path $script:DATD -Force -ErrorAction Stop | Out-Null
        }
        Set-Content -Path (Join-Path $script:DATD 'lic.dat') -Value $key -Encoding ASCII -ErrorAction Stop
    } catch { }
}

function Test-Licensed {
    if ($script:FREE) { return $true }
    $mc = Get-MachineCode
    $validKey = Calculate-Key $mc
    $saved = Get-SavedKey
    return ($saved.Trim().ToUpper() -eq $validKey.ToUpper())
}

function Test-Allowed {
    return (Test-Licensed)
}

function Get-StatusText {
    if (Test-Licensed) { return "Litsenziya faol" }
    return "Faollashtirilmagan"
}

function Invoke-Gate {
    if (Test-Licensed) { return $true }
    Add-Type -AssemblyName Microsoft.VisualBasic
    Add-Type -AssemblyName System.Windows.Forms
    
    $mc = Get-MachineCode
    $validKey = Calculate-Key $mc
    
    $msg = "Dastur faollashtirilmagan!`r`n`r`nSizning kompyuter kodingiz: $mc`r`n`r`nLitsenziya kalitini olish uchun muallifga murojaat qiling:`r`nMuallif: Fozliddinov Safarmurod`r`nTelegram: @Safar_master`r`nTelefon: +998939530473`r`n`r`nLitsenziya kalitini kiriting:"
    $code = [Microsoft.VisualBasic.Interaction]::InputBox($msg, "AlifboO'tkazgich: Litsenziya", "")
    if ([string]::IsNullOrWhiteSpace($code)) { return $false }
    
    if ($code.Trim().ToUpper() -eq $validKey.ToUpper()) {
        Set-SavedKey $validKey
        [void][System.Windows.Forms.MessageBox]::Show('Litsenziya muvaffaqiyatli faollashtirildi! Rahmat.', "AlifboO'tkazgich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return $true
    }
    
    [void][System.Windows.Forms.MessageBox]::Show("Litsenziya kaliti noto'g'ri!`r`nMuallifga murojaat qiling: @Safar_master", "AlifboO'tkazgich", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    return $false
}
