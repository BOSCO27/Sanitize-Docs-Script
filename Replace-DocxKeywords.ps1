<#
.SYNOPSIS
    Bulk find-and-replace across .docx files in a folder.

.DESCRIPTION
    Scans a folder for .docx files and applies a list of find/replace pairs to
    document body text, tables, text boxes, headers, footers, footnotes,
    endnotes, and comments.

    Word frequently splits a visible phrase across multiple XML runs (spell
    check, revision tracking, formatting). A naive replace misses those. This
    script handles both cases: it replaces within a single run where possible
    (preserving formatting), and merges the minimum span of runs where a match
    crosses a run boundary.

    Requires no Microsoft Word installation. Works on Windows PowerShell 5.1
    and PowerShell 7+.

.PARAMETER Path
    Folder containing .docx files. If omitted, the script prompts; pressing
    Enter at that prompt uses the current folder.

.PARAMETER MappingsFile
    Keyword list with columns Find,Replace. One row per keyword pair.
    Accepts either a .csv or an .xlsx / .xlsm workbook (first worksheet,
    headers in the first populated row). Excel is read natively - no add-in
    module and no Excel installation required.
    If omitted, the script prompts; pressing Enter at that prompt uses the
    single .xlsx or .csv found in the scan folder, when there is exactly one.

.PARAMETER OutputFolder
    Where modified copies are written. Originals are untouched.
    Omit this and use -InPlace to modify originals instead.

.PARAMETER InPlace
    Modify originals rather than writing copies. Creates a .bak unless
    -NoBackup is specified.

.PARAMETER NoBackup
    With -InPlace, skip creating .bak files.

.PARAMETER Recurse
    Include subfolders.

.PARAMETER CaseSensitive
    Match case exactly. Default is case-insensitive.

.PARAMETER DryRun
    Report what would change without writing anything.

.PARAMETER LogFile
    Optional CSV path for a per-file, per-keyword replacement log.

.EXAMPLE
    .\Replace-DocxKeywords.ps1 -Path .\Docs -MappingsFile .\map.csv -OutputFolder .\Sanitized -Recurse

.EXAMPLE
    .\Replace-DocxKeywords.ps1 -Path .\Docs -MappingsFile .\map.csv -DryRun -Recurse

.EXAMPLE
    .\Replace-DocxKeywords.ps1 -Path .\Docs -MappingsFile .\map.csv -InPlace
#>

[CmdletBinding()]
param(
    [string]$Path,

    [string]$MappingsFile,

    [string]$OutputFolder,

    [switch]$InPlace,

    [switch]$NoBackup,

    [switch]$Recurse,

    [switch]$CaseSensitive,

    [switch]$DryRun,

    [string]$LogFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$WNS   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$XMLNS = 'http://www.w3.org/XML/1998/namespace'

# ---------------------------------------------------------------- validation

$here = (Get-Location).Path

if (-not $Path) {
    Write-Host ""
    Write-Host "Which folder holds the .docx files?"
    Write-Host "  Enter a folder path, or press Enter to use the current folder:"
    Write-Host "  $here"
    Write-Host ""
    $answer = (Read-Host "Path").Trim().Trim('"').Trim("'")
    $Path = if ([string]::IsNullOrWhiteSpace($answer)) { $here } else { $answer }
}

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    throw "Path not found or not a folder: $Path"
}

if (-not $MappingsFile) {
    $found = @(
        Get-ChildItem -LiteralPath $Path -File |
        Where-Object { $_.Extension -in '.xlsx', '.xlsm', '.csv' -and $_.Name -notlike '~$*' }
    )

    Write-Host ""
    Write-Host "Which file holds the Find/Replace list?"
    if ($found.Length -eq 1) {
        Write-Host "  Enter a file path, or press Enter to use the one found here:"
        Write-Host "  $($found[0].Name)"
    }
    elseif ($found.Length -gt 1) {
        Write-Host "  Enter a file path. Candidates in this folder:"
        foreach ($f in $found) { Write-Host "    $($f.Name)" }
    }
    else {
        Write-Host "  Enter a file path. No .xlsx or .csv found in the scan folder."
    }
    Write-Host ""

    $answer = (Read-Host "MappingsFile").Trim().Trim('"').Trim("'")

    if ([string]::IsNullOrWhiteSpace($answer)) {
        if ($found.Length -eq 1) {
            $MappingsFile = $found[0].FullName
            Write-Host "Using $($found[0].Name)" -ForegroundColor DarkGray
        }
        else {
            throw "No mappings file given, and none could be chosen automatically."
        }
    }
    else {
        $MappingsFile = $answer
    }
}

if (-not (Test-Path -LiteralPath $MappingsFile -PathType Leaf)) {
    throw "Mappings file not found: $MappingsFile"
}
if ($InPlace -and $OutputFolder) {
    throw "Use either -OutputFolder or -InPlace, not both."
}

if (-not $InPlace -and -not $OutputFolder) {
    $default = Join-Path $Path 'Sanitized'
    Write-Host ""
    Write-Host "Where should the modified files be written?"
    Write-Host "  Enter a folder path, or"
    Write-Host "  press Enter for the default: $default, or"
    Write-Host "  type INPLACE to modify the originals (a .bak copy is kept)."
    Write-Host ""
    $answer = Read-Host "Output"
    $answer = $answer.Trim().Trim('"').Trim("'")

    if ([string]::IsNullOrWhiteSpace($answer)) {
        $OutputFolder = $default
    }
    elseif ($answer -match '^(?i)inplace$') {
        $InPlace = $true
        Write-Host "Originals will be modified. A .bak copy is kept for each file." -ForegroundColor Yellow
    }
    else {
        $OutputFolder = $answer
    }
}

function Get-MappingsFromXlsx {
    param([string]$XlsxPath)

    $zip = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path -LiteralPath $XlsxPath).Path)
    try {
        function Read-ZipXml {
            param($Archive, [string]$EntryName)
            $entry = $Archive.Entries | Where-Object { $_.FullName -eq $EntryName }
            if (-not $entry) { return $null }
            $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
            $text = $reader.ReadToEnd()
            $reader.Dispose()
            $doc = New-Object System.Xml.XmlDocument
            $doc.LoadXml($text)
            return $doc
        }

        $SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        $RL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $PR = 'http://schemas.openxmlformats.org/package/2006/relationships'

        # Resolve the first worksheet through the workbook relationships.
        $wb = Read-ZipXml -Archive $zip -EntryName 'xl/workbook.xml'
        if (-not $wb) { throw "Not a readable .xlsx file: $XlsxPath" }

        $nsWb = New-Object System.Xml.XmlNamespaceManager($wb.NameTable)
        $nsWb.AddNamespace('s', $SS)
        $nsWb.AddNamespace('r', $RL)
        $firstSheet = $wb.SelectSingleNode('//s:sheets/s:sheet', $nsWb)
        if (-not $firstSheet) { throw "No worksheets found in $XlsxPath" }
        $rid = $firstSheet.GetAttribute('id', $RL)

        $sheetPath = 'xl/worksheets/sheet1.xml'
        $rels = Read-ZipXml -Archive $zip -EntryName 'xl/_rels/workbook.xml.rels'
        if ($rels -and $rid) {
            $nsR = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
            $nsR.AddNamespace('p', $PR)
            $rel = $rels.SelectSingleNode("//p:Relationship[@Id='$rid']", $nsR)
            if ($rel) {
                $t = $rel.GetAttribute('Target')
                $sheetPath = if ($t.StartsWith('/')) { $t.TrimStart('/') } else { "xl/$t" }
                $sheetPath = $sheetPath -replace '/\./', '/'
            }
        }

        # Shared string table, if the file uses one (Excel normally does).
        $shared = New-Object 'System.Collections.Generic.List[string]'
        $sst = Read-ZipXml -Archive $zip -EntryName 'xl/sharedStrings.xml'
        if ($sst) {
            $nsS = New-Object System.Xml.XmlNamespaceManager($sst.NameTable)
            $nsS.AddNamespace('s', $SS)
            foreach ($si in $sst.SelectNodes('//s:si', $nsS)) {
                $val = ''
                foreach ($t in $si.SelectNodes('.//s:t', $nsS)) { $val += $t.InnerText }
                $shared.Add($val)
            }
        }

        # Parse the worksheet.
        $sheet = Read-ZipXml -Archive $zip -EntryName $sheetPath
        if (-not $sheet) { throw "Worksheet not found: $sheetPath" }

        $nsSheet = New-Object System.Xml.XmlNamespaceManager($sheet.NameTable)
        $nsSheet.AddNamespace('s', $SS)

        $rows = @($sheet.SelectNodes('//s:sheetData/s:row', $nsSheet))
        if ($rows.Length -eq 0) { throw "Worksheet in $XlsxPath is empty." }

        # Find the header row and column indices.
        $headerRow = $null
        $findIdx = -1
        $replaceIdx = -1
        foreach ($r in $rows) {
            $cells = @($r.SelectNodes('.//s:c', $nsSheet))
            foreach ($i in 0..($cells.Length - 1)) {
                $cell = $cells[$i]
                $v = $cell.SelectSingleNode('.//s:v', $nsSheet)
                if (-not $v) {
                    $t = $cell.SelectSingleNode('.//s:t', $nsSheet)
                    if ($t) { $text = $t.InnerText } else { $text = '' }
                } else {
                    $text = $v.InnerText
                    $idx = [int][regex]::Match($text, '\d+').Value
                    if ($idx -ge 0 -and $idx -lt $shared.Count) { $text = $shared[$idx] }
                }

                if ($text -eq 'Find') { $findIdx = $i }
                if ($text -eq 'Replace') { $replaceIdx = $i }
            }

            if ($findIdx -ge 0 -and $replaceIdx -ge 0) {
                $headerRow = $r
                break
            }
        }

        if (-not $headerRow) { throw "No 'Find' and 'Replace' header row found in $XlsxPath" }

        # Extract data rows.
        $grid = @{}
        $inData = $false
        foreach ($r in $rows) {
            if ($r -eq $headerRow) { $inData = $true; continue }
            if (-not $inData) { continue }

            $cells = @($r.SelectNodes('.//s:c', $nsSheet))
            $findVal = $null
            $replaceVal = $null

            foreach ($i in 0..($cells.Length - 1)) {
                $cell = $cells[$i]
                $v = $cell.SelectSingleNode('.//s:v', $nsSheet)
                $text = if ($v) { $v.InnerText } else { '' }

                if ($text -match '^\d+$') {
                    $idx = [int]$text
                    if ($idx -ge 0 -and $idx -lt $shared.Count) { $text = $shared[$idx] }
                }

                $t = $cell.SelectSingleNode('.//s:t', $nsSheet)
                if ($t) { $text = $t.InnerText }

                if ($i -eq $findIdx) { $findVal = $text }
                if ($i -eq $replaceIdx) { $replaceVal = $text }
            }

            if ($findVal -and $replaceVal) {
                $grid[$findVal] = $replaceVal
            }
        }

        if ($grid.Keys.Count -eq 0) { throw "Worksheet in $XlsxPath has no data rows." }

        return $grid
    } finally {
        $zip.Dispose()
    }
}

function Get-MappingsFromCsv {
    param([string]$CsvPath)
    $grid = @{}
    $rows = @(Import-Csv -LiteralPath $CsvPath -Encoding UTF8)
    foreach ($row in $rows) {
        if ($row.Find -and $row.Replace) {
            $grid[$row.Find] = $row.Replace
        }
    }
    return $grid
}

# Load mappings.
Write-Host ""
Write-Host "Reading mappings from Excel workbook: $MappingsFile"

$ext = [System.IO.Path]::GetExtension($MappingsFile).ToLower()
$mappingsGrid = if ($ext -in '.xlsx', '.xlsm') {
    Get-MappingsFromXlsx -XlsxPath $MappingsFile
} else {
    Get-MappingsFromCsv -CsvPath $MappingsFile
}

$mappings = @(
    $mappingsGrid.GetEnumerator() |
    ForEach-Object { [PSCustomObject]@{ Find = $_.Key; Replace = $_.Value } }
)

if ($mappings.Length -eq 0) { throw "No usable rows in the mappings file." }

$cmp = if ($CaseSensitive) {
    [System.StringComparison]::Ordinal
} else {
    [System.StringComparison]::OrdinalIgnoreCase
}

# ----------------------------------------------------------------- functions

function Invoke-LiteralReplace {
    param([string]$Text, [string]$Find, [string]$Replace,
          [System.StringComparison]$Comparison)

    $sb = New-Object System.Text.StringBuilder
    $i = 0
    $n = 0
    while ($true) {
        $j = $Text.IndexOf($Find, $i, $Comparison)
        if ($j -lt 0) { break }
        [void]$sb.Append($Text.Substring($i, $j - $i))
        [void]$sb.Append($Replace)
        $i = $j + $Find.Length
        $n++
    }
    [void]$sb.Append($Text.Substring($i))
    [PSCustomObject]@{ Text = $sb.ToString(); Hits = $n }
}

function Invoke-PartReplace {
    param([byte[]]$Bytes, $Mappings, [System.StringComparison]$Comparison)

    $xml = New-Object System.Xml.XmlDocument
    $ms = New-Object System.IO.MemoryStream
    $ms.Write($Bytes, 0, $Bytes.Length)
    $ms.Seek(0, [System.IO.SeekOrigin]::Begin) | Out-Null
    $xml.Load($ms)
    $ms.Dispose()

    $nsm = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsm.AddNamespace('w', $WNS)

    $tally = @{}
    foreach ($m in $Mappings) { $tally[$m.Find] = 0 }

    $paras = @($xml.SelectNodes('.//w:p', $nsm))
    foreach ($p in $paras) {
        $tNodes = @($p.SelectNodes('.//w:t', $nsm))
        if ($tNodes.Length -eq 0) { continue }

        # Pass 1 - match contained within a single run. Formatting preserved.
        foreach ($t in $tNodes) {
            $orig = $t.InnerText
            $cur  = $orig
            foreach ($m in $Mappings) {
                $r = Invoke-LiteralReplace -Text $cur -Find $m.Find `
                        -Replace $m.Replace -Comparison $Comparison
                if ($r.Hits -gt 0) {
                    $cur = $r.Text
                    $tally[$m.Find] += $r.Hits
                }
            }
            if ($cur -ne $orig) {
                $t.InnerText = $cur
                [void]$t.SetAttribute('space', $XMLNS, 'preserve')
            }
        }

        # Pass 2 - match spans runs. Merge the minimum span of runs needed.
        foreach ($m in $Mappings) {
            for ($guard = 0; $guard -lt 500; $guard++) {

                $texts = New-Object 'System.Collections.Generic.List[string]'
                foreach ($t in $tNodes) { $texts.Add($t.InnerText) }
                $combined = -join $texts

                $pos = $combined.IndexOf($m.Find, 0, $Comparison)
                if ($pos -lt 0) { break }
                $end = $pos + $m.Find.Length - 1

                $starts = New-Object 'System.Collections.Generic.List[int]'
                $ends   = New-Object 'System.Collections.Generic.List[int]'
                $acc = 0
                foreach ($s in $texts) {
                    $starts.Add($acc)
                    $ends.Add($acc + $s.Length - 1)
                    $acc += $s.Length
                }

                $si = -1; $ei = -1
                for ($i = 0; $i -lt $texts.Count; $i++) {
                    if ($texts[$i].Length -eq 0) { continue }
                    if ($si -lt 0 -and $pos -ge $starts[$i] -and $pos -le $ends[$i]) { $si = $i }
                    if ($end -ge $starts[$i] -and $end -le $ends[$i]) { $ei = $i; break }
                }
                if ($si -lt 0 -or $ei -lt 0) { break }

                $span   = -join $texts[$si..$ei]
                $offset = $pos - $starts[$si]
                $merged = $span.Substring(0, $offset) + $m.Replace +
                          $span.Substring($offset + $m.Find.Length)

                $tNodes[$si].InnerText = $merged
                [void]$tNodes[$si].SetAttribute('space', $XMLNS, 'preserve')
                for ($i = $si + 1; $i -le $ei; $i++) {
                    $tNodes[$i].InnerText = ''
                }
                $tally[$m.Find] += 1
            }
        }
    }

    $total = 0
    foreach ($k in $tally.Keys) { $total += $tally[$k] }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent   = $false
    $out = New-Object System.IO.MemoryStream
    $writer = [System.Xml.XmlWriter]::Create($out, $settings)
    $xml.Save($writer)
    $writer.Flush(); $writer.Close()
    $bytes = $out.ToArray()
    $out.Dispose()

    [PSCustomObject]@{ Bytes = $bytes; Total = $total; Tally = $tally }
}

function Test-TargetPart {
    param([string]$Name)
    if ($Name -eq 'word/document.xml')  { return $true }
    if ($Name -eq 'word/footnotes.xml') { return $true }
    if ($Name -eq 'word/endnotes.xml')  { return $true }
    if ($Name -eq 'word/comments.xml')  { return $true }
    if ($Name -match '^word/(header|footer)\d*\.xml$') { return $true }
    return $false
}

# --------------------------------------------------------------------- main

# If the output folder sits inside the scan folder, skip anything already in it
# so repeat runs don't re-process their own output.
$excludeRoot = $null
if ($OutputFolder -and (Test-Path -LiteralPath $OutputFolder)) {
    $excludeRoot = (Resolve-Path -LiteralPath $OutputFolder).Path.TrimEnd('\', '/') + [IO.Path]::DirectorySeparatorChar
}

$files = @(Get-ChildItem -LiteralPath $Path -Filter *.docx -File -Recurse:$Recurse |
         Where-Object { $_.Name -notlike '~$*' } |
         Where-Object {
             -not $excludeRoot -or
             -not $_.FullName.StartsWith($excludeRoot, [System.StringComparison]::OrdinalIgnoreCase)
         })

if ($files.Length -eq 0) {
    Write-Warning "No .docx files found under $Path"
    return
}

Write-Host ""
Write-Host "Files found : $($files.Length)"
Write-Host "Keywords    : $($mappings.Length)"
Write-Host "Mode        : $(if ($DryRun) { 'DRY RUN - nothing will be written' }
                            elseif ($InPlace) { 'In place' }
                            else { "Copies to $OutputFolder" })"
Write-Host ""

$log     = @()
$summary = @()

foreach ($file in $files) {

    $entries = @{}
    $order   = New-Object 'System.Collections.Generic.List[string]'

    $zip = [System.IO.Compression.ZipFile]::OpenRead($file.FullName)
    try {
        foreach ($e in $zip.Entries) {
            $s  = $e.Open()
            $ms = New-Object System.IO.MemoryStream
            $s.CopyTo($ms)
            $entries[$e.FullName] = $ms.ToArray()
            $order.Add($e.FullName)
            $ms.Dispose(); $s.Dispose()
        }
    } finally {
        $zip.Dispose()
    }

    $fileTotal = 0
    $fileTally = @{}
    foreach ($m in $mappings) { $fileTally[$m.Find] = 0 }

    foreach ($name in $order) {
        if (-not (Test-TargetPart -Name $name)) { continue }
        $res = @(Invoke-PartReplace -Bytes $entries[$name] -Mappings $mappings -Comparison $cmp)[-1]
        if ($res.Total -gt 0) {
            $entries[$name] = $res.Bytes
            $fileTotal += $res.Total
            foreach ($k in $res.Tally.Keys) { $fileTally[$k] += $res.Tally[$k] }
        }
    }

    foreach ($m in $mappings) {
        if ($fileTally[$m.Find] -gt 0) {
            $log += [PSCustomObject]@{
                File         = $file.FullName
                Find         = $m.Find
                Replace      = $m.Replace
                Replacements = $fileTally[$m.Find]
            }
        }
    }

    $summary += [PSCustomObject]@{
        File         = $file.Name
        Replacements = $fileTotal
    }

    if ($fileTotal -eq 0) {
        Write-Host ("  {0,-55} no matches" -f $file.Name) -ForegroundColor DarkGray
        continue
    }

    Write-Host ("  {0,-55} {1} replacement(s)" -f $file.Name, $fileTotal) -ForegroundColor Green

    if ($DryRun) { continue }

    if ($InPlace) {
        if (-not $NoBackup) {
            Copy-Item -LiteralPath $file.FullName -Destination "$($file.FullName).bak" -Force
        }
        $dest = $file.FullName
    } else {
        $rel = $file.FullName.Substring((Resolve-Path -LiteralPath $Path).Path.Length).TrimStart('\', '/')
        $dest = Join-Path $OutputFolder $rel
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
    }

    $tmp = [System.IO.Path]::GetTempFileName()
    Remove-Item -LiteralPath $tmp -Force
    $tmp = "$tmp.docx"

    $out = [System.IO.Compression.ZipFile]::Open($tmp, 'Create')
    try {
        foreach ($name in $order) {
            $entry  = $out.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
            $stream = $entry.Open()
            $stream.Write($entries[$name], 0, $entries[$name].Length)
            $stream.Dispose()
        }
    } finally {
        $out.Dispose()
    }

    Move-Item -LiteralPath $tmp -Destination $dest -Force
}

# ------------------------------------------------------------------ reporting

$grand = ($summary | Measure-Object -Property Replacements -Sum).Sum
Write-Host ""
Write-Host "Total replacements: $grand across $($files.Length) file(s)."

if ($DryRun) {
    Write-Host "Dry run - no files were modified." -ForegroundColor Yellow
}

if ($LogFile -and $log.Length -gt 0) {
    $log | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8
    Write-Host "Log written to $LogFile"
}

Write-Host ""
