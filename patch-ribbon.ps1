# Ofis qo'shimcha faylga (.xlam / .ppam) "Alifbo" lentasini (customUI) qo'shadi
param([string]$Path, [string]$XmlPath)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$bytes = [System.IO.File]::ReadAllBytes($XmlPath)
$fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $old = $zip.GetEntry('customUI/customUI14.xml')
    if ($old) { $old.Delete() }
    $e = $zip.CreateEntry('customUI/customUI14.xml')
    $st = $e.Open(); $st.Write($bytes, 0, $bytes.Length); $st.Dispose()
    $re = $zip.GetEntry('_rels/.rels')
    if (-not $re) { throw '_rels/.rels topilmadi' }
    $sr = New-Object System.IO.StreamReader($re.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $rels = $sr.ReadToEnd(); $sr.Dispose()
    if ($rels -notmatch 'customUI14\.xml') {
        $rel = '<Relationship Id="rIdAlifboUI" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
        $rels = $rels -replace '</Relationships>', ($rel + '</Relationships>')
        $nb = (New-Object System.Text.UTF8Encoding($false)).GetBytes($rels)
        $re.Delete()
        $newRe = $zip.CreateEntry('_rels/.rels')
        $st = $newRe.Open(); $st.Write($nb, 0, $nb.Length); $st.Dispose()
    }
    # PPam uchun: presentation.xml 0-bayt bo'lsa, minimal valid XML kiritish
    if ($Path -match '\.ppam$') {
        $presEntry = $zip.GetEntry('ppt/presentation.xml')
        $presEmpty = ($presEntry -eq $null -or $presEntry.Length -eq 0)
        if ($presEmpty) {
            if ($presEntry) { $presEntry.Delete() }
            $presXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1"><p:sldMasterIdLst/><p:sldSz cx="9144000" cy="6858000" type="screen4x3"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>'
            $pb = (New-Object System.Text.UTF8Encoding($false)).GetBytes($presXml)
            $pe = $zip.CreateEntry('ppt/presentation.xml')
            $ps = $pe.Open(); $ps.Write($pb, 0, $pb.Length); $ps.Dispose()
        }
    }
} finally { $zip.Dispose(); $fs.Dispose() }
