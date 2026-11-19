# ==============================================================
# alifbo-sssr.ps1 - SSSR kitoblarini Wordga o'tkazish boshqaruvchisi
# Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
# ==============================================================
param([string]$FilePath = "")

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pyScript = Join-Path $scriptDir "alifbo-sssr-ocr.py"

# Python mavjudligini tekshirish
$pyCmd = Get-Command python -ErrorAction SilentlyContinue
if ($pyCmd) {
    if ([string]::IsNullOrWhiteSpace($FilePath)) {
        Start-Process python -ArgumentList "`"$pyScript`""
    } else {
        Start-Process python -ArgumentList "`"$pyScript`" `"$FilePath`""
    }
} else {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "SSSR kitoblarini tahlil qilish uchun Python o'rnatilgan bo'lishi lozim.`nPython orqali sarg'aygan fon tozalash va formulalarni saqlash yuksak aniqlikda ishlaydi.",
        "AlifboOtkazgich - SSSR Kitob OCR",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}
