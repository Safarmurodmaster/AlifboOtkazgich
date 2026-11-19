# ==============================================================
# alifbo-pdf2docx.ps1 - PDF fayllarni tahrirlanuvchi Word (.docx) ga o'tkazish
# Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
# ==============================================================
param(
    [string]$PdfPath = "",
    [string]$DocxPath = "",
    [switch]$OpenAfter = $true
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-PdfPicker {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Word-ga o'tkaziladigan PDF faylni tanlang"
    $ofd.Filter = "PDF Fayllar (*.pdf)|*.pdf|Barcha fayllar (*.*)|*.*"
    $ofd.Multiselect = $true
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $ofd.FileNames
    }
    return $null
}

$filesToConvert = @()
if ([string]::IsNullOrWhiteSpace($PdfPath)) {
    $selected = Show-PdfPicker
    if (-not $selected) { exit 0 }
    $filesToConvert = $selected
} else {
    $filesToConvert = @($PdfPath)
}

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $count = 0
    foreach ($file in $filesToConvert) {
        if (-not (Test-Path $file)) { continue }
        $outPath = ""
        if (-not [string]::IsNullOrWhiteSpace($DocxPath) -and $filesToConvert.Count -eq 1) {
            $outPath = $DocxPath
        } else {
            $outPath = [System.IO.Path]::ChangeExtension($file, ".docx")
        }

        # PDF ni Word bilan ochish (Word 2013+ o'zi to'liq tahrirlanuvchi DOCX ga reflow qiladi)
        $doc = $word.Documents.Open($file, $false, $true, $false)
        if ($doc) {
            # wdFormatXMLDocument = 16 (DOCX)
            $doc.SaveAs2($outPath, 16)
            $doc.Close($false)
            $count++
            Write-Host "O'tkazildi: $outPath" -ForegroundColor Green
        }
    }

    if ($count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Muvaffaqiyatli yakunlandi!`n$count ta PDF fayl Word (.docx) formatiga o'tkazildi.",
            "AlifboOtkazgich - PDF to Word",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        if ($OpenAfter -and $filesToConvert.Count -eq 1) {
            $firstDocx = [System.IO.Path]::ChangeExtension($filesToConvert[0], ".docx")
            if (Test-Path $firstDocx) {
                Start-Process $firstDocx
            }
        }
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Xatolik yuz berdi: $($_.Exception.Message)",
        "Xatolik",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}
finally {
    if ($word) {
        $word.Quit(0)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}
