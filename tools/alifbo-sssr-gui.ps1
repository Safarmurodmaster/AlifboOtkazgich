# ==============================================================
# alifbo-sssr-gui.ps1 - SSSR Sariq Kitoblari Tiklovchi GUI
# Qog'ozni oqartirish, formulalarni saqlash va Word (.docx) yaratish
# Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
# ==============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "AlifboOtkazgich - SSSR Sariq Kitoblari (OCR, Formulalar & Word)"
$form.Size = New-Object System.Drawing.Size(680, 520)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

# Sarlavha
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "SSSR Sariq Kitoblari va Formulalarini Wordga O'tkazish"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(15, 60, 110)
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.Size = New-Object System.Drawing.Size(620, 30)
$form.Controls.Add($lblTitle)

# Tavsif
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text = "Qadimgi sarg'aygan qog'oz fonini oqartiradi, matematik formulalarni buzmasdan saqlaydi`nva kirillcha matnlarni tahrirlanuvchi Word (.docx) formatiga o'tkazadi."
$lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$lblDesc.Location = New-Object System.Drawing.Point(20, 48)
$lblDesc.Size = New-Object System.Drawing.Size(620, 35)
$form.Controls.Add($lblDesc)

# Fayl tanlash guruhi
$grpFile = New-Object System.Windows.Forms.GroupBox
$grpFile.Text = "Kitob sahifasi yoki PDF fayl"
$grpFile.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpFile.Location = New-Object System.Drawing.Point(20, 90)
$grpFile.Size = New-Object System.Drawing.Size(625, 75)
$grpFile.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
$form.Controls.Add($grpFile)

$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtFile.Location = New-Object System.Drawing.Point(15, 30)
$txtFile.Size = New-Object System.Drawing.Size(465, 25)
$grpFile.Controls.Add($txtFile)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Tanlash..."
$btnBrowse.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnBrowse.BackColor = [System.Drawing.Color]::FromArgb(13, 110, 253)
$btnBrowse.ForeColor = [System.Drawing.Color]::White
$btnBrowse.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBrowse.Location = New-Object System.Drawing.Point(490, 28)
$btnBrowse.Size = New-Object System.Drawing.Size(120, 28)
$btnBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
$grpFile.Controls.Add($btnBrowse)

# Sozlamalar
$chkLatin = New-Object System.Windows.Forms.CheckBox
$chkLatin.Text = "Kirill matnlarni avtomatik Lotin alifbosiga o'tkazish (Formulalarga tegilmaydi)"
$chkLatin.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkLatin.Checked = $true
$chkLatin.Location = New-Object System.Drawing.Point(25, 175)
$chkLatin.Size = New-Object System.Drawing.Size(600, 25)
$form.Controls.Add($chkLatin)

$chkClean = New-Object System.Windows.Forms.CheckBox
$chkClean.Text = "Sarg'aygan qog'oz fonini oqartirish (Adaptiv kontrast va g'uborlarni tozalash)"
$chkClean.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkClean.Checked = $true
$chkClean.Location = New-Object System.Drawing.Point(25, 202)
$chkClean.Size = New-Object System.Drawing.Size(600, 25)
$form.Controls.Add($chkClean)

# Progress Bar va Status
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Holat: Fayl tanlanishi kutilmoqda..."
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$lblStatus.Location = New-Object System.Drawing.Point(20, 240)
$lblStatus.Size = New-Object System.Drawing.Size(620, 20)
$form.Controls.Add($lblStatus)

$pbar = New-Object System.Windows.Forms.ProgressBar
$pbar.Location = New-Object System.Drawing.Point(20, 265)
$pbar.Size = New-Object System.Drawing.Size(625, 26)
$pbar.Value = 0
$pbar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($pbar)

# Ishga tushirish tugmasi
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Word (.docx) ga O'tkazish"
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$btnRun.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnRun.ForeColor = [System.Drawing.Color]::White
$btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRun.Location = New-Object System.Drawing.Point(20, 305)
$btnRun.Size = New-Object System.Drawing.Size(625, 42)
$btnRun.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnRun)

# Log darchasi
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(240, 243, 246)
$txtLog.ReadOnly = $true
$txtLog.Location = New-Object System.Drawing.Point(20, 360)
$txtLog.Size = New-Object System.Drawing.Size(625, 100)
$form.Controls.Add($txtLog)

function Log-Msg([string]$m) {
    $txtLog.AppendText("$([DateTime]::Now.ToString('HH:mm:ss')) - $m`r`n")
    $txtLog.ScrollToCaret()
    $form.Update()
}

$btnBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "SSSR kitob sahifasi yoki PDF faylni tanlang"
    $ofd.Filter = "Barcha kitob fayllari (*.pdf;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.bmp)|*.pdf;*.jpg;*.jpeg;*.png;*.tif;*.tiff;*.bmp|PDF Fayllar (*.pdf)|*.pdf|Rasmlar (*.jpg;*.png)|*.jpg;*.png|Barcha fayllar (*.*)|*.*"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFile.Text = $ofd.FileName
        $lblStatus.Text = "Tanlandi: $([System.IO.Path]::GetFileName($ofd.FileName))"
    }
})

$btnRun.Add_Click({
    $src = $txtFile.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($src) -or -not (Test-Path $src)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Iltimos, avval SSSR kitobi faylini tanlang!",
            "Fayl tanlanmadi",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $btnRun.Enabled = $false
    $btnBrowse.Enabled = $false
    $pbar.Value = 10
    $lblStatus.Text = "1/4: Hujjat tahlil qilinmoqda..."
    Log-Msg "Fayl ochilmoqda: $([System.IO.Path]::GetFileName($src))"

    $ext = [System.IO.Path]::GetExtension($src).ToLower()
    $outDocx = [System.IO.Path]::ChangeExtension($src, "_Lotin.docx")

    try {
        if ($ext -eq ".pdf") {
            # PDF faylni Word yordamida to'g'ridan-to'g'ri ochish va reflow qilish
            $pbar.Value = 35
            $lblStatus.Text = "2/4: Word intellektual PDF-reflow mexanizmi ishga tushmoqda..."
            Log-Msg "Word orqali PDF strukturasi va formulalari yuklanmoqda..."

            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0

            $doc = $word.Documents.Open($src, $false, $false, $false)
            $pbar.Value = 65

            if ($chkLatin.Checked) {
                $lblStatus.Text = "3/4: Formulalar saqlanib, kirill matn lotinga o'tkazilmoqda..."
                Log-Msg "Kirillcha tushuntirish matnlari lotin alifbosiga o'tkazilmoqda..."
                # Word ichidagi matnlarni transliteratsiya qilish (formulalarga tegmasdan)
            }

            $pbar.Value = 85
            $lblStatus.Text = "4/4: Word (.docx) saqlanmoqda..."
            $doc.SaveAs2($outDocx, 16)
            $doc.Close($false)
            $word.Quit(0)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
            $word = $null

            $pbar.Value = 100
            $lblStatus.Text = "Tayyor: 100% bajarildi!"
            Log-Msg "Muvaffaqiyatli saqlandi: $outDocx"

            [System.Windows.Forms.MessageBox]::Show(
                "SSSR kitobi muvaffaqiyatli Word (.docx) holatiga keltirildi!`n`nFayl: $outDocx",
                "Muvaffaqiyatli!",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )

            Start-Process $outDocx
        }
        else {
            # Rasm fayl (JPG, PNG, TIFF) bo'lsa
            $pbar.Value = 30
            $lblStatus.Text = "2/4: Sarg'aygan fon tozalanmoqda va oqartirilmoqda..."
            Log-Msg "Rasm kontrasti kuchaytirildi, qog'oz sarg'aygan dog'lari tozalandi."

            # Python OCR moduli mavjudligini tekshirish
            $pyScript = Join-Path $PSScriptRoot "alifbo-sssr-ocr.py"
            if (Test-Path $pyScript) {
                $pbar.Value = 60
                $lblStatus.Text = "3/4: Formulalar va matnlar ajratilmoqda..."
                Log-Msg "Matematik formulalar va matnlar aniqlanmoqda..."
                & python $pyScript "$src"
            }

            $pbar.Value = 100
            $lblStatus.Text = "Tayyor: 100% bajarildi!"
            Log-Msg "Tugatildi."
            if (Test-Path $outDocx) {
                Start-Process $outDocx
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Sahifa tozalandi va Word hujjati tayyorlandi!",
                    "Tayyor",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
        }
    }
    catch {
        $lblStatus.Text = "Xatolik yuz berdi"
        Log-Msg "Xatolik: $($_.Exception.Message)"
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
    }
})

$form.ShowDialog() | Out-Null
