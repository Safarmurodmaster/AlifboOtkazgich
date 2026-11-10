# AlifboÖtkazğich: fayllarni yangi alifboga o'tkazish (Word, Excel, PowerPoint, matn fayllari)
# Muallif: Fozliddinov Safarmurod, Telegram @Safar_master, +998939530473
param([string[]]$Files, [string]$Mode = '')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'alifbo-lic.ps1')
Initialize-AlifboCore $here
$supported = @('.docx', '.xlsx', '.pptx', '.txt', '.md', '.csv')
$oldFormats = @('.doc', '.xls', '.ppt')

function Get-OutPath([string]$src, [bool]$toNew) {
    $dir = Split-Path -Parent $src
    $base = [System.IO.Path]::GetFileNameWithoutExtension($src)
    $ext = [System.IO.Path]::GetExtension($src)
    $suffix = if ($toNew) { '_yangi' } else { '_eski' }
    $p = Join-Path $dir ($base + $suffix + $ext); $i = 2
    while (Test-Path $p) { $p = Join-Path $dir ($base + $suffix + '(' + $i + ')' + $ext); $i++ }
    return $p
}
function Convert-One([string]$src, [bool]$toNew, [bool]$replace, [bool]$backup, $exc) {
    $ext = [System.IO.Path]::GetExtension($src).ToLowerInvariant()
    if ($oldFormats -contains $ext) { return @{ ok = $false; msg = 'Eski format ' + $ext + ': avval faylni yangi shaklda (.docx / .xlsx / .pptx) saqlang.' } }
    if ($supported -notcontains $ext) { return @{ ok = $false; msg = 'Qo''llab-quvvatlanmaydi: ' + $ext } }
    try {
        if ($replace) {
            $tmp = $src + '.alifbo.tmp'
            $n = [AlifboCore]::ConvertFile($src, $tmp, $toNew, $exc)
            if ($backup) { Copy-Item -LiteralPath $src -Destination ($src + '.bak') -Force }
            Move-Item -LiteralPath $tmp -Destination $src -Force
            return @{ ok = $true; n = $n; dst = $src }
        } else {
            $dst = Get-OutPath $src $toNew
            $n = [AlifboCore]::ConvertFile($src, $dst, $toNew, $exc)
            return @{ ok = $true; n = $n; dst = $dst }
        }
    } catch {
        return @{ ok = $false; msg = ('Xatolik (fayl ochiq bo''lishi mumkin, Office''ni yoping): ' + $_.Exception.Message) }
    }
}

# ---------- kontekst menyudan: bitta/bir nechta fayl ----------
if ($Files -and $Files.Count -gt 0 -and ($Mode -eq 'new' -or $Mode -eq 'old')) {
    if (-not (Invoke-Gate)) { exit }
    $toNew = ($Mode -eq 'new'); $exc = Get-ExcSet; $lines = @(); $okc = 0
    foreach ($f in $Files) {
        $r = Convert-One $f $toNew $false $false $exc
        if ($r.ok) { $okc++; $lines += ('OK: ' + (Split-Path -Leaf $r.dst) + ' (' + $r.n + ' ta bo''lak)') }
        else { $lines += ('XATO: ' + (Split-Path -Leaf $f) + ' - ' + $r.msg) }
    }
    $head = if ($okc -gt 0) { 'Tayyor. Yangi fayl asl fayl yonida saqlandi.' } else { 'O''tkazib bo''lmadi.' }
    [void][System.Windows.Forms.MessageBox]::Show(($head + "`r`n`r`n" + ($lines -join "`r`n") + "`r`n`r`n" + (Get-StatusText)), 'AlifboÖtkazğich')
    exit
}

# ---------- asosiy oyna ----------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'AlifboÖtkazğich: fayllarni o''tkazish'
$form.Size = New-Object System.Drawing.Size(720, 640)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = 'Word (.docx), Excel (.xlsx), PowerPoint (.pptx) va matn (.txt, .md, .csv) fayllarini o''tkazadi. Office kerak emas.'
$lbl.Location = '15,12'; $lbl.Size = '670,40'; $form.Controls.Add($lbl)

$list = New-Object System.Windows.Forms.ListBox
$list.Location = '15,58'; $list.Size = '670,170'; $list.HorizontalScrollbar = $true; $list.SelectionMode = 'MultiExtended'
$form.Controls.Add($list)

function Add-FileToList([string]$p) {
    $e = [System.IO.Path]::GetExtension($p).ToLowerInvariant()
    if (($supported + $oldFormats) -contains $e) { if (-not $list.Items.Contains($p)) { [void]$list.Items.Add($p) } }
}
$b1 = New-Object System.Windows.Forms.Button; $b1.Text = 'Fayl qo''shish...'; $b1.Location = '15,236'; $b1.Size = '150,34'
$b1.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    $dlg.Filter = 'Hujjatlar|*.docx;*.xlsx;*.pptx;*.txt;*.md;*.csv|Hammasi|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { foreach ($f in $dlg.FileNames) { Add-FileToList $f } }
})
$b2 = New-Object System.Windows.Forms.Button; $b2.Text = 'Papka qo''shish...'; $b2.Location = '175,236'; $b2.Size = '150,34'
$b2.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq 'OK') {
        Get-ChildItem -LiteralPath $dlg.SelectedPath -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notlike '~$*' -and $_.Name -notlike '*_yangi.*' -and $_.Name -notlike '*_eski.*' } |
            Select-Object -First 3000 | ForEach-Object { Add-FileToList $_.FullName }
    }
})
$b3 = New-Object System.Windows.Forms.Button; $b3.Text = 'Tozalash'; $b3.Location = '335,236'; $b3.Size = '110,34'
$b3.Add_Click({ $list.Items.Clear() })
$form.Controls.AddRange(@($b1, $b2, $b3))

$g1 = New-Object System.Windows.Forms.GroupBox; $g1.Text = 'Yo''nalish'; $g1.Location = '15,280'; $g1.Size = '330,80'
$rNew = New-Object System.Windows.Forms.RadioButton; $rNew.Text = 'Eski alifbo → yangi alifbo'; $rNew.Location = '12,24'; $rNew.Size = '300,24'; $rNew.Checked = $true
$rOld = New-Object System.Windows.Forms.RadioButton; $rOld.Text = 'Yangi alifbo → eski alifbo'; $rOld.Location = '12,48'; $rOld.Size = '300,24'
$g1.Controls.AddRange(@($rNew, $rOld)); $form.Controls.Add($g1)

$g2 = New-Object System.Windows.Forms.GroupBox; $g2.Text = 'Natija'; $g2.Location = '355,280'; $g2.Size = '330,80'
$rCopy = New-Object System.Windows.Forms.RadioButton; $rCopy.Text = 'Yangi fayl (nom_yangi.docx)'; $rCopy.Location = '12,24'; $rCopy.Size = '300,24'; $rCopy.Checked = $true
$rRep = New-Object System.Windows.Forms.RadioButton; $rRep.Text = 'Asl faylni almashtirish'; $rRep.Location = '12,48'; $rRep.Size = '300,24'
$g2.Controls.AddRange(@($rCopy, $rRep)); $form.Controls.Add($g2)

$chkBak = New-Object System.Windows.Forms.CheckBox; $chkBak.Text = 'Almashtirishda zaxira nusxa (.bak) saqlansin'; $chkBak.Location = '355,366'; $chkBak.Size = '330,24'; $chkBak.Checked = $true
$form.Controls.Add($chkBak)

$go = New-Object System.Windows.Forms.Button; $go.Text = 'Boshlash'; $go.Location = '15,372'; $go.Size = '200,40'
$go.BackColor = [System.Drawing.Color]::FromArgb(15, 108, 189); $go.ForeColor = [System.Drawing.Color]::White; $go.FlatStyle = 'Flat'
$form.Controls.Add($go)

$bar = New-Object System.Windows.Forms.ProgressBar; $bar.Location = '15,422'; $bar.Size = '670,20'; $form.Controls.Add($bar)
$log = New-Object System.Windows.Forms.TextBox; $log.Multiline = $true; $log.ReadOnly = $true; $log.ScrollBars = 'Vertical'
$log.Location = '15,450'; $log.Size = '670,100'; $form.Controls.Add($log)
$ft = New-Object System.Windows.Forms.Label
$ft.Text = ('Muallif: Fozliddinov Safarmurod  |  Telegram: @Safar_master  |  +998939530473   (' + (Get-StatusText) + ')')
$ft.Location = '15,562'; $ft.Size = '670,36'; $ft.ForeColor = [System.Drawing.Color]::Gray; $form.Controls.Add($ft)

$go.Add_Click({
    if ($list.Items.Count -eq 0) { [void][System.Windows.Forms.MessageBox]::Show('Avval fayl yoki papka qo''shing.', 'AlifboÖtkazğich'); return }
    if (-not (Invoke-Gate)) { return }
    if ($rRep.Checked) {
        $q = [System.Windows.Forms.MessageBox]::Show('Asl fayllar o''zgartiriladi. Davom etasizmi?', 'AlifboÖtkazğich', 'YesNo', 'Question')
        if ($q -ne 'Yes') { return }
    }
    $exc = Get-ExcSet; $toNew = $rNew.Checked; $go.Enabled = $false
    $log.Clear(); $bar.Minimum = 0; $bar.Maximum = $list.Items.Count; $bar.Value = 0; $ok = 0; $bad = 0
    foreach ($f in @($list.Items)) {
        $r = Convert-One $f $toNew $rRep.Checked $chkBak.Checked $exc
        if ($r.ok) { $ok++; $log.AppendText('OK: ' + (Split-Path -Leaf $r.dst) + ' (' + $r.n + ' ta bo''lak)' + "`r`n") }
        else { $bad++; $log.AppendText('XATO: ' + (Split-Path -Leaf $f) + ' - ' + $r.msg + "`r`n") }
        $bar.Value = [math]::Min($bar.Maximum, $bar.Value + 1)
        [System.Windows.Forms.Application]::DoEvents()
    }
    $log.AppendText("`r`nTugadi: " + $ok + ' ta muvaffaqiyatli, ' + $bad + ' ta xato.')
    $go.Enabled = $true
})
[void]$form.ShowDialog()
