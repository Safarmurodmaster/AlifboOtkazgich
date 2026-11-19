# ==============================================================
# alifbo-sssr-gui.ps1 - Sariq va Qadimgi Kitoblarni Wordga O'tkazish (GUI)
# Qog'ozni oqartirish, formulalarni saqlash va Word (.docx) yaratish
# Muallif: Fozliddinov Safarmurod | @Safar_master | +998939530473
# ==============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = "AlifboOtkazgich - Sariq va Eskirgan Kitoblar (Formulalar & Word)"
$form.Size = New-Object System.Drawing.Size(680, 560)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)

# Sarlavha
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Sariq Kitoblar va Formulalarni Wordga O'tkazish"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::FromArgb(15, 60, 110)
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.Size = New-Object System.Drawing.Size(620, 30)
$form.Controls.Add($lblTitle)

# Tavsif
$lblDesc = New-Object System.Windows.Forms.Label
$lblDesc.Text = "Sarg'aygan qog'oz fonini oqartiradi, matematik formulalarni buzmasdan saqlaydi`nva kirillcha matnlarni tahrirlanuvchi Word (.docx) formatiga o'tkazadi."
$lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblDesc.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$lblDesc.Location = New-Object System.Drawing.Point(20, 48)
$lblDesc.Size = New-Object System.Drawing.Size(620, 35)
$form.Controls.Add($lblDesc)

# Fayl tanlash guruhi
$grpFile = New-Object System.Windows.Forms.GroupBox
$grpFile.Text = "Kitob fayli (PDF yoki Rasm)"
$grpFile.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpFile.Location = New-Object System.Drawing.Point(20, 88)
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

# Sahifalarni tanlash guruhi
$grpPages = New-Object System.Windows.Forms.GroupBox
$grpPages.Text = "Kerakli betlar oralig'i (Sahifalar)"
$grpPages.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpPages.Location = New-Object System.Drawing.Point(20, 170)
$grpPages.Size = New-Object System.Drawing.Size(625, 68)
$grpPages.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
$form.Controls.Add($grpPages)

$lblPagesHint = New-Object System.Windows.Forms.Label
$lblPagesHint.Text = "Sahifalar (masalan: 1-50 yoki 1, 3, 5-10):"
$lblPagesHint.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblPagesHint.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$lblPagesHint.Location = New-Object System.Drawing.Point(15, 28)
$lblPagesHint.Size = New-Object System.Drawing.Size(250, 23)
$grpPages.Controls.Add($lblPagesHint)

$txtPages = New-Object System.Windows.Forms.TextBox
$txtPages.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$txtPages.Location = New-Object System.Drawing.Point(270, 25)
$txtPages.Size = New-Object System.Drawing.Size(180, 25)
$grpPages.Controls.Add($txtPages)

$lblPagesInfo = New-Object System.Windows.Forms.Label
$lblPagesInfo.Text = "Bo'sh = barcha betlar"
$lblPagesInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Italic)
$lblPagesInfo.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
$lblPagesInfo.Location = New-Object System.Drawing.Point(460, 28)
$lblPagesInfo.Size = New-Object System.Drawing.Size(155, 23)
$grpPages.Controls.Add($lblPagesInfo)

# Sozlamalar
$chkLatin = New-Object System.Windows.Forms.CheckBox
$chkLatin.Text = "Kirill matnlarni avtomatik Lotin alifbosiga o'tkazish (Formulalarga tegilmaydi)"
$chkLatin.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkLatin.Checked = $true
$chkLatin.Location = New-Object System.Drawing.Point(25, 245)
$chkLatin.Size = New-Object System.Drawing.Size(600, 23)
$form.Controls.Add($chkLatin)

# Progress Bar va Status
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Holat: Fayl tanlanishi kutilmoqda..."
$lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$lblStatus.Location = New-Object System.Drawing.Point(20, 275)
$lblStatus.Size = New-Object System.Drawing.Size(620, 20)
$form.Controls.Add($lblStatus)

$pbar = New-Object System.Windows.Forms.ProgressBar
$pbar.Location = New-Object System.Drawing.Point(20, 298)
$pbar.Size = New-Object System.Drawing.Size(625, 24)
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
$btnRun.Location = New-Object System.Drawing.Point(20, 332)
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
$txtLog.Location = New-Object System.Drawing.Point(20, 385)
$txtLog.Size = New-Object System.Drawing.Size(625, 120)
$form.Controls.Add($txtLog)

function Log-Msg([string]$m) {
    $txtLog.AppendText("$([DateTime]::Now.ToString('HH:mm:ss')) - $m`r`n")
    $txtLog.ScrollToCaret()
    $form.Update()
}

$btnBrowse.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Kitob sahifasi yoki PDF faylni tanlang"
    $ofd.Filter = "Kitob fayllari (*.pdf;*.jpg;*.jpeg;*.png)|*.pdf;*.jpg;*.jpeg;*.png|PDF Hujjatlar (*.pdf)|*.pdf|Rasmlar (*.jpg;*.png)|*.jpg;*.png|Barcha fayllar (*.*)|*.*"
    if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFile.Text = $ofd.FileName
        $fname = [System.IO.Path]::GetFileName($ofd.FileName)
        $lblStatus.Text = "Tanlandi: $fname"
        
        # Agar PDF bo'lsa, sahifalar sonini aniqlashga urinish
        if ($ofd.FileName.ToLower().EndsWith(".pdf")) {
            try {
                $pyCheck = "import fitz; doc=fitz.open(r'''$($ofd.FileName)'''); print(len(doc))"
                $totalP = & python -c $pyCheck 2>$null
                if ($totalP -and $totalP.Trim() -match '^\d+$') {
                    $lblPagesInfo.Text = "Jami betlar: $($totalP.Trim()) ta"
                    $lblPagesInfo.ForeColor = [System.Drawing.Color]::FromArgb(15, 100, 15)
                }
            } catch {}
        }
    }
})

$btnRun.Add_Click({
    $src = $txtFile.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($src) -or -not (Test-Path $src)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Iltimos, avval kitob faylini tanlang!",
            "Fayl tanlanmadi",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $pages = $txtPages.Text.Trim()
    $latinFlag = if ($chkLatin.Checked) { "1" } else { "0" }

    $btnRun.Enabled = $false
    $btnBrowse.Enabled = $false
    $pbar.Value = 10
    $lblStatus.Text = "1/4: Hujjat va sahifalar tekshirilmoqda..."
    $form.Update()

    Log-Msg "Fayl: $([System.IO.Path]::GetFileName($src))"
    if ($pages) {
        Log-Msg "Tanlangan sahifalar: $pages"
    } else {
        Log-Msg "Sahifalar: Barchasi"
    }

    $dotPos = $src.LastIndexOf(".")
    $baseName = if ($dotPos -gt 0) { $src.Substring(0, $dotPos) } else { $src }
    $outDocx = $baseName + "_Word.docx"

    try {
        $pbar.Value = 30
        $lblStatus.Text = "2/4: Kerakli sahifalar ajratilmoqda va Wordga o'tkazilmoqda..."
        Log-Msg "Sahifalar tahlil qilinmoqda, formulalar aniqlanmoqda..."
        $form.Update()

        $scriptDir = $PSScriptRoot
        if (-not $scriptDir) { $scriptDir = "E:\ekran\new alifbo\AlifboOtkazgich-kodlar\tools" }
        $pyScript = Join-Path $scriptDir "alifbo-sssr-ocr.py"

        if (-not (Test-Path $pyScript)) {
            $pyScript = Join-Path $env:LOCALAPPDATA "AlifboOtkazgich\tools\alifbo-sssr-ocr.py"
        }

        if (-not (Test-Path $pyScript)) {
            throw "alifbo-sssr-ocr.py moduli topilmadi: $pyScript"
        }

        $pbar.Value = 50
        $lblStatus.Text = "3/4: Formulalar saqlanib, Word hujjati yaratilmoqda..."
        $form.Update()

        # Python jarayonini ishga tushirish (xavfsiz va tez)
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "python.exe"
        $psi.Arguments = "`"$pyScript`" `"$src`" `"$pages`" `"$latinFlag`""
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $proc = [System.Diagnostics.Process]::Start($psi)
        while (-not $proc.HasExited) {
            $form.Update()
            Start-Sleep -Milliseconds 200
        }

        $stdout = $proc.StandardOutput.ReadToEnd()
        $stderr = $proc.StandardError.ReadToEnd()

        if ($stdout) {
            foreach ($line in $stdout.Split("`n")) {
                if ($line.Trim()) { Log-Msg $line.Trim() }
            }
        }

        if ($proc.ExitCode -ne 0) {
            if ($stderr) {
                Log-Msg "Xatolik tafsiloti: $stderr"
            }
            throw "Konvertatsiya jarayonida xatolik yuz berdi (Kod: $($proc.ExitCode))"
        }

        $pbar.Value = 100
        $lblStatus.Text = "Tayyor: 100% bajarildi!"
        Log-Msg "Muvaffaqiyatli saqlandi: $outDocx"
        $form.Update()

        $res = [System.Windows.Forms.MessageBox]::Show(
            "Kitob muvaffaqiyatli Word (.docx) holatiga keltirildi!`n`nFayl: $outDocx`n`nHujjatni hozir Word da ochishni xohlaysizmi?",
            "Muvaffaqiyatli!",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )

        if ($res -eq [System.Windows.Forms.DialogResult]::Yes -and (Test-Path $outDocx)) {
            Start-Process $outDocx
        }
    }
    catch {
        $lblStatus.Text = "Xatolik yuz berdi"
        Log-Msg "Xatolik: $($_.Exception.Message)"
        $form.Update()
        [System.Windows.Forms.MessageBox]::Show(
            "Xatolik: $($_.Exception.Message)",
            "Xatolik",
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
