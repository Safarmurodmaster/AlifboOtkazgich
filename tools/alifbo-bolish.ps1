# AlifboÖtkazğich: Word hujjatini varaqlar bo'yicha alohida fayllarga bo'lish
# Muallif: Fozliddinov Safarmurod, Telegram @Safar_master, +998939530473
param([string]$DocPath = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'alifbo-lic.ps1')
if (-not (Invoke-Gate)) { exit }

$script:word = $null; $script:doc = $null; $script:cancel = $false; $script:pieces = $null

function Close-Word {
    try { if ($script:doc) { $script:doc.Close($false) } } catch { }
    try { if ($script:word) { $script:word.Quit() } } catch { }
    $script:doc = $null; $script:word = $null
}
function Open-Doc([string]$path) {
    Close-Word
    $script:word = New-Object -ComObject Word.Application
    $script:word.Visible = $false
    $script:word.DisplayAlerts = 0
    $script:doc = $script:word.Documents.Open($path, $false, $true)
    $script:doc.Repaginate()
    return [int]$script:doc.ComputeStatistics(2)
}
function Get-SafeName([string]$s) {
    if (-not $s) { return '' }
    $s = ($s -replace '[\r\n\a\v]', ' ') -replace '[\\/:*?"<>|]', ' '
    $s = ($s -replace '\s{2,}', ' ').Trim()
    if ($s.Length -gt 50) { $s = $s.Substring(0, 50).Trim() }
    return $s
}
function Add-PagePiece([int]$a, [int]$b, [int]$total) {
    if ($a -lt 1) { $a = 1 }
    if ($b -gt $total) { $b = $total }
    if ($b -lt $a) { return }
    $d = $script:doc
    $s = $d.GoTo(1, 1, $a).Start
    if ($b -lt $total) { $e = $d.GoTo(1, 1, $b + 1).Start } else { $e = $d.Content.End }
    [void]$script:pieces.Add([pscustomobject]@{ S = $s; E = $e; PA = $a; PB = $b; Title = '' })
}
function Get-Pieces([string]$mode, [int]$n, [string]$ranges, [int]$total) {
    $script:pieces = New-Object System.Collections.ArrayList
    $d = $script:doc
    switch ($mode) {
        'each' { for ($i = 1; $i -le $total; $i++) { Add-PagePiece $i $i $total } }
        'n' { if ($n -lt 1) { $n = 1 }; for ($i = 1; $i -le $total; $i += $n) { Add-PagePiece $i ($i + $n - 1) $total } }
        'range' {
            foreach ($part in ($ranges -split ',')) {
                $p = $part.Trim(); if (-not $p) { continue }
                if ($p -match '^(\d+)\s*-\s*(\d+)$') { Add-PagePiece ([int]$Matches[1]) ([int]$Matches[2]) $total }
                elseif ($p -match '^\d+$') { Add-PagePiece ([int]$p) ([int]$p) $total }
            }
        }
        'heading' {
            $starts = New-Object System.Collections.ArrayList
            foreach ($par in $d.Paragraphs) {
                if ([int]$par.OutlineLevel -eq 1) { [void]$starts.Add([pscustomobject]@{ S = [int]$par.Range.Start; T = (Get-SafeName $par.Range.Text) }) }
            }
            if ($starts.Count -eq 0) { return $script:pieces }
            if ($starts[0].S -gt 0) { [void]$script:pieces.Add([pscustomobject]@{ S = 0; E = $starts[0].S; PA = 0; PB = 0; Title = 'Kirish' }) }
            for ($i = 0; $i -lt $starts.Count; $i++) {
                $e = if ($i + 1 -lt $starts.Count) { $starts[$i + 1].S } else { [int]$d.Content.End }
                [void]$script:pieces.Add([pscustomobject]@{ S = $starts[$i].S; E = $e; PA = 0; PB = 0; Title = $starts[$i].T })
            }
            foreach ($pc in $script:pieces) {
                $pc.PA = [int]$d.Range($pc.S, $pc.S).Information(3)
                $pc.PB = [int]$d.Range([math]::Max($pc.S, $pc.E - 1), [math]::Max($pc.S, $pc.E - 1)).Information(3)
            }
        }
    }
    return $script:pieces
}

# ---------- oyna ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'AlifboÖtkazğich: varaqlarni bo''lish'
$form.Size = New-Object System.Drawing.Size(740, 700)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$l1 = New-Object System.Windows.Forms.Label; $l1.Text = 'Word hujjati:'; $l1.Location = '15,15'; $l1.Size = '120,24'; $form.Controls.Add($l1)
$tbFile = New-Object System.Windows.Forms.TextBox; $tbFile.Location = '15,40'; $tbFile.Size = '580,26'; $form.Controls.Add($tbFile)
$bFile = New-Object System.Windows.Forms.Button; $bFile.Text = 'Tanlash...'; $bFile.Location = '605,38'; $bFile.Size = '100,30'; $form.Controls.Add($bFile)
$lInfo = New-Object System.Windows.Forms.Label; $lInfo.Text = 'Hujjat tanlanmagan.'; $lInfo.Location = '15,72'; $lInfo.Size = '690,24'; $lInfo.ForeColor = [System.Drawing.Color]::DimGray; $form.Controls.Add($lInfo)

$g = New-Object System.Windows.Forms.GroupBox; $g.Text = 'Qanday bo''lamiz?'; $g.Location = '15,102'; $g.Size = '690,170'; $form.Controls.Add($g)
$r1 = New-Object System.Windows.Forms.RadioButton; $r1.Text = 'Har bir varaq alohida fayl'; $r1.Location = '15,26'; $r1.Size = '300,24'; $r1.Checked = $true
$r2 = New-Object System.Windows.Forms.RadioButton; $r2.Text = 'Har N varaq bitta faylga:'; $r2.Location = '15,56'; $r2.Size = '230,24'
$nud = New-Object System.Windows.Forms.NumericUpDown; $nud.Location = '250,54'; $nud.Size = '70,26'; $nud.Minimum = 1; $nud.Maximum = 5000; $nud.Value = 5
$r3 = New-Object System.Windows.Forms.RadioButton; $r3.Text = 'Sarlavhalar bo''yicha (1-daraja sarlavha, Heading 1)'; $r3.Location = '15,86'; $r3.Size = '450,24'
$r4 = New-Object System.Windows.Forms.RadioButton; $r4.Text = 'Oraliqlar bo''yicha:'; $r4.Location = '15,116'; $r4.Size = '170,24'
$tbR = New-Object System.Windows.Forms.TextBox; $tbR.Location = '190,114'; $tbR.Size = '480,26'; $tbR.Text = '1-5, 6-12, 13-20'
$g.Controls.AddRange(@($r1, $r2, $nud, $r3, $r4, $tbR))

$g2 = New-Object System.Windows.Forms.GroupBox; $g2.Text = 'Format va nom'; $g2.Location = '15,282'; $g2.Size = '690,110'; $form.Controls.Add($g2)
$cPdf = New-Object System.Windows.Forms.CheckBox; $cPdf.Text = 'PDF'; $cPdf.Location = '15,26'; $cPdf.Size = '80,24'; $cPdf.Checked = $true
$cDoc = New-Object System.Windows.Forms.CheckBox; $cDoc.Text = 'Word (.docx)'; $cDoc.Location = '100,26'; $cDoc.Size = '140,24'; $cDoc.Checked = $true
$cName = New-Object System.Windows.Forms.CheckBox; $cName.Text = 'Fayl nomiga sarlavha (birinchi qator) qo''shilsin'; $cName.Location = '15,60'; $cName.Size = '420,24'
$g2.Controls.AddRange(@($cPdf, $cDoc, $cName))

$l2 = New-Object System.Windows.Forms.Label; $l2.Text = 'Saqlash papkasi:'; $l2.Location = '15,402'; $l2.Size = '200,24'; $form.Controls.Add($l2)
$tbOut = New-Object System.Windows.Forms.TextBox; $tbOut.Location = '15,428'; $tbOut.Size = '580,26'; $form.Controls.Add($tbOut)
$bOut = New-Object System.Windows.Forms.Button; $bOut.Text = 'Tanlash...'; $bOut.Location = '605,426'; $bOut.Size = '100,30'; $form.Controls.Add($bOut)

$go = New-Object System.Windows.Forms.Button; $go.Text = 'Boshlash'; $go.Location = '15,466'; $go.Size = '180,40'
$go.BackColor = [System.Drawing.Color]::FromArgb(15, 108, 189); $go.ForeColor = [System.Drawing.Color]::White; $go.FlatStyle = 'Flat'
$bStop = New-Object System.Windows.Forms.Button; $bStop.Text = 'To''xtatish'; $bStop.Location = '205,466'; $bStop.Size = '120,40'; $bStop.Enabled = $false
$form.Controls.AddRange(@($go, $bStop))
$bar = New-Object System.Windows.Forms.ProgressBar; $bar.Location = '15,516'; $bar.Size = '690,20'; $form.Controls.Add($bar)
$log = New-Object System.Windows.Forms.TextBox; $log.Multiline = $true; $log.ReadOnly = $true; $log.ScrollBars = 'Vertical'; $log.Location = '15,544'; $log.Size = '690,70'; $form.Controls.Add($log)
$ft = New-Object System.Windows.Forms.Label
$ft.Text = ('Muallif: Fozliddinov Safarmurod  |  @Safar_master  |  +998939530473   (' + (Get-StatusText) + ')')
$ft.Location = '15,620'; $ft.Size = '690,24'; $ft.ForeColor = [System.Drawing.Color]::Gray; $form.Controls.Add($ft)

$script:total = 0
function Load-Selected([string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return }
    $tbFile.Text = $path
    $lInfo.Text = 'Hujjat ochilmoqda...'
    [System.Windows.Forms.Application]::DoEvents()
    try {
        $script:total = Open-Doc $path
        $lInfo.Text = ('Hujjatda ' + $script:total + ' ta varaq bor.')
        $base = [System.IO.Path]::GetFileNameWithoutExtension($path)
        $tbOut.Text = Join-Path (Split-Path -Parent $path) ($base + '_bolingan')
    } catch {
        $lInfo.Text = 'Hujjatni ochib bo''lmadi. Word o''rnatilganini tekshiring.'
        [void][System.Windows.Forms.MessageBox]::Show(('Xatolik: ' + $_.Exception.Message), 'AlifboÖtkazğich')
    }
}
$bFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Word hujjatlari|*.docx;*.doc;*.docm|Hammasi|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { Load-Selected $dlg.FileName }
})
$bOut.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($tbOut.Text) { $dlg.SelectedPath = $tbOut.Text }
    if ($dlg.ShowDialog() -eq 'OK') { $tbOut.Text = $dlg.SelectedPath }
})
$bStop.Add_Click({ $script:cancel = $true })

$go.Add_Click({
    if (-not $script:doc) { [void][System.Windows.Forms.MessageBox]::Show('Avval hujjatni tanlang.', 'AlifboÖtkazğich'); return }
    if (-not ($cPdf.Checked -or $cDoc.Checked)) { [void][System.Windows.Forms.MessageBox]::Show('PDF yoki Word formatini belgilang.', 'AlifboÖtkazğich'); return }
    $mode = if ($r1.Checked) { 'each' } elseif ($r2.Checked) { 'n' } elseif ($r3.Checked) { 'heading' } else { 'range' }
    $go.Enabled = $false; $bStop.Enabled = $true; $script:cancel = $false; $log.Clear()
    try {
        $pcs = Get-Pieces $mode ([int]$nud.Value) $tbR.Text $script:total
        if ($pcs.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('Bo''linadigan qism topilmadi (sarlavha uslubi yoki oraliqlarni tekshiring).', 'AlifboÖtkazğich')
            return
        }
        $out = $tbOut.Text
        if (-not (Test-Path -LiteralPath $out)) { New-Item -ItemType Directory -Path $out -Force | Out-Null }
        $base = [System.IO.Path]::GetFileNameWithoutExtension($tbFile.Text)
        $bar.Minimum = 0; $bar.Maximum = $pcs.Count; $bar.Value = 0; $ok = 0; $bad = 0
        $d = $script:doc; $w = $script:word; $i = 0
        foreach ($pc in $pcs) {
            if ($script:cancel) { $log.AppendText("To'xtatildi.`r`n"); break }
            $i++
            $fn = Join-Path $out ($base + '_' + $i.ToString('000'))
            if ($cName.Checked) {
                $t = $pc.Title
                if (-not $t) { try { $t = Get-SafeName (($d.Range($pc.S, [math]::Min($pc.E, $pc.S + 400)).Text -split "[\r\n]" | Where-Object { $_.Trim() } | Select-Object -First 1)) } catch { $t = '' } }
                if ($t) { $fn = $fn + '_' + $t }
            }
            $good = $true
            if ($cPdf.Checked) {
                try {
                    $pa = $pc.PA; $pb = $pc.PB
                    if ($pa -lt 1) { $pa = 1 }; if ($pb -lt $pa) { $pb = $pa }
                    $d.ExportAsFixedFormat(($fn + '.pdf'), 17, $false, 0, 3, $pa, $pb)
                } catch { $good = $false; $log.AppendText(('XATO ' + $i + ' (PDF): ' + $_.Exception.Message + "`r`n")) }
            }
            if ($cDoc.Checked) {
                $nd = $null
                try {
                    $nd = $w.Documents.Add()
                    try {
                        $nd.PageSetup.Orientation = $d.PageSetup.Orientation
                        $nd.PageSetup.PageWidth = $d.PageSetup.PageWidth; $nd.PageSetup.PageHeight = $d.PageSetup.PageHeight
                        $nd.PageSetup.TopMargin = $d.PageSetup.TopMargin; $nd.PageSetup.BottomMargin = $d.PageSetup.BottomMargin
                        $nd.PageSetup.LeftMargin = $d.PageSetup.LeftMargin; $nd.PageSetup.RightMargin = $d.PageSetup.RightMargin
                    } catch { }
                    $rng = $d.Range($pc.S, $pc.E)
                    try {
                        $nd.Content.FormattedText = $rng.FormattedText
                    } catch {
                        $rng.Copy()
                        $nd.Content.Paste()
                    }
                    $path = $fn + '.docx'
                    try {
                        $nd.SaveAs2($path, 16)
                    } catch {
                        $nd.SaveAs($path, 16)
                    }
                } catch { $good = $false; $log.AppendText(('XATO ' + $i + ' (Word): ' + $_.Exception.Message + "`r`n")) }
                finally { try { if ($nd) { $nd.Close(0) } } catch { } }
            }
            if ($good) { $ok++ } else { $bad++ }
            $log.AppendText(('Saqlandi: ' + (Split-Path -Leaf $fn) + "`r`n"))
            $bar.Value = [math]::Min($bar.Maximum, $i)
            [System.Windows.Forms.Application]::DoEvents()
        }
        $log.AppendText("`r`nTugadi: " + $ok + ' ta muvaffaqiyatli, ' + $bad + ' ta xato.')
        if ($ok -gt 0) { Start-Process explorer.exe -ArgumentList ('"' + $out + '"') }
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(('Xatolik: ' + $_.Exception.Message), 'AlifboÖtkazğich')
    } finally {
        $go.Enabled = $true; $bStop.Enabled = $false
    }
})
$form.Add_Shown({ if ($DocPath) { Load-Selected $DocPath } })
$form.Add_FormClosed({ Close-Word })
[void]$form.ShowDialog()
Close-Word
