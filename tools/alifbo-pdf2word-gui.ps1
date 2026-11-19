# ==============================================================
# alifbo-pdf2word-gui.ps1 - PDF fayllarni Wordga o'tkazish (GUI)
# Professional progress bar bilan, qora oyna YOQILMAYDI
# Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
# ==============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "AlifboOtkazgich - PDF fayllarni Word (.docx) ga o'tkazish"
$form.Size = New-Object System.Drawing.Size(680, 440)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

# Sarlavha
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "PDF Fayllarni Tahrirlanuvchi Word (.docx) ga O'tkazish"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(15, 60, 110)
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.Size = New-Object System.Drawing.Size(620, 30)
$form.Controls.Add($lblTitle)

$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text = "Microsoft Word-ning o'rnatilgan PDF-reflow texnologiyasi orqali jadvallar, shriftlar,`nformatlar va rasmlarni to'liq saqlab qolgan holda, to'liq tahrirlanuvchi Word hujjatini yaratadi."
$lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$lblDesc.Location = New-Object System.Drawing.Point(20, 50)
$lblDesc.Size = New-Object System.Drawing.Size(620, 40)
$form.Controls.Add($lblDesc)

# Fayl tanlash guruhi
$grpFile = New-Object System.Windows.Forms.GroupBox
$grpFile.Text = "PDF fayl"
$grpFile.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpFile.Location = New-Object System.Drawing.Point(20, 100)
$grpFile.Size = New-Object System.Drawing.Size(625, 70)
$form.Controls.Add($grpFile)

$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtFile.Location = New-Object System.Drawing.Point(15, 28)
$txtFile.Size = New-Object System.Drawing.Size(465, 25)
$grpFile.Controls.Add($txtFile)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "PDF tanlash..."
$btnBrowse.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(13, 110, 253)
$btnBrowse.ForeColor = [System.Drawing.Color]::White
$btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBrowse.Location = New-Object System.Drawing.Point(490, 26)
$btnBrowse.Size = New-Object System.Drawing.Size(120, 28)
$grpFile.Controls.Add($btnBrowse)

# Ma'lumot
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "✔  Word 2013 va undan yuqori versiyalar talab etiladi`n✔  PDF faylining matni kopirlana oladigan bo'lishi tavsiya etiladi`n✔  Skanerdan olingan fayllar uchun SSSR Kitob (OCR) vositasidan foydalaning"
$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblInfo.ForeColor = [System.Drawing.Color]::FromArgb(40, 120, 40)
$lblInfo.Location = New-Object System.Drawing.Point(20, 185)
$lblInfo.Size = New-Object System.Drawing.Size(625, 60)
$form.Controls.Add($lblInfo)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Holat: PDF fayl tanlanishi kutilmoqda..."
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$lblStatus.Location = New-Object System.Drawing.Point(20, 260)
$lblStatus.Size = New-Object System.Drawing.Size(620, 20)
$form.Controls.Add($lblStatus)

$pbar = New-Object System.Windows.Forms.ProgressBar
$pbar.Location = New-Object System.Drawing.Point(20, 285)
$pbar.Size = New-Object System.Drawing.Size(625, 26)
$pbar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($pbar)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "PDF ni Word (.docx) ga O'tkazish"
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRun.Location = New-Object System.Drawing.Point(20, 325)
$btnRun.Size = New-Object System.Drawing.Size(625, 42)
$btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnRun)

$lblFt = New-Object System.Windows.Forms.Label
$lblFt.Text = "Muallif: Fozliddinov Safarmurod  |  Telegram: @Safar_master  |  +998939530473"
$lblFt.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblFt.ForeColor = [System.Drawing.Color]::FromArgb(150, 150, 150)
$lblFt.Location = New-Object System.Drawing.Point(20, 378)
$lblFt.Size = New-Object System.Drawing.Size(620, 20)
$form.Controls.Add($lblFt)

# Fayl tanlash
$btnBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Word (.docx) ga o'tkaziladigan PDF faylni tanlang"
    $ofd.Filter = "PDF Hujjatlar (*.pdf)|*.pdf|Barcha fayllar (*.*)|*.*"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFile.Text = $ofd.FileName
        $lblStatus.Text = "Tanlandi: $([System.IO.Path]::GetFileName($ofd.FileName))"
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(15, 100, 15)
    }
})

# O'tkazish
$btnRun.Add_Click({
    $src = $txtFile.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($src) -or -not (Test-Path $src)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Iltimos, avval PDF faylini tanlang!",
            "Fayl tanlanmadi",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $btnRun.Enabled = $false
    $btnBrowse.Enabled = $false

    $pbar.Value = 0
    $lblStatus.Text = "1/3: Microsoft Word ishga tushirilmoqda..."
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $form.Update()

    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0

        $pbar.Value = 30
        $lblStatus.Text = "2/3: PDF tahlil qilinmoqda va Word hujjatiga aylantirilmoqda..."
        $form.Update()

        # PDF ni to'g'ridan-to'g'ri Word'ga ochish (intellektual reflow)
        $doc = $word.Documents.Open($src, $false, $false, $false)

        $pbar.Value = 70
        $lblStatus.Text = "3/3: Word (.docx) saqlanmoqda..."
        $form.Update()

        $dotPos = $src.LastIndexOf(".")
        $baseName = if ($dotPos -gt 0) { $src.Substring(0, $dotPos) } else { $src }
        $outDocx = $baseName + ".docx"

        # DOCX sifatida saqlash
        $doc.SaveAs2($outDocx, 16) # wdFormatXMLDocument = 16

        $doc.Close($false)
        $word.Quit(0)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        $word = $null

        $pbar.Value = 100
        $lblStatus.Text = "Tayyor! 100% muvaffaqiyatli bajarildi."
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(15, 130, 15)
        $form.Update()

        $res = [System.Windows.Forms.MessageBox]::Show(
            "PDF muvaffaqiyatli Word (.docx) formatiga o'tkazildi!`n`nFayl: $outDocx`n`nWord da ochish uchun 'Ha' tugmasini bosing.",
            "Muvaffaqiyatli!",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process $outDocx
        }
    }
    catch {
        $lblStatus.Text = "Xatolik yuz berdi!"
        $lblStatus.ForeColor = [System.Drawing.Color]::Red
        if ($doc -ne $null) { try { $doc.Close($false) } catch {} }
        if ($word -ne $null) { try { $word.Quit(0) } catch {} }
        if ($word -ne $null) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
        [System.Windows.Forms.MessageBox]::Show(
            "Xatolik: $($_.Exception.Message)",
            "Xato",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        $btnRun.Enabled = $true
        $btnBrowse.Enabled = $true
        if ($pbar.Value -lt 100) { $pbar.Value = 0 }
    }
})

$form.ShowDialog() | Out-Null
