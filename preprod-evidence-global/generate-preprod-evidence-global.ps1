param(
    [string]$InputDir,
    [string]$OutputDir,
    [string]$OutputFileName,
    [string]$OutputBaseName = 'UFO - FMC Pre-Deployment Evidence',
    [string]$ReportTitle = 'UFO - FMC Pre-Deployment Evidence',
    [string]$SofficePath,
    [switch]$ForceOverwrite
)

$ErrorActionPreference = 'Stop'

if (-not $InputDir) {
    $InputDir = Join-Path $PSScriptRoot 'input'
}
if (-not $OutputDir) {
    $OutputDir = Join-Path $PSScriptRoot 'out'
}

function Resolve-SofficePath {
    param([string]$CustomPath)

    if ($CustomPath -and (Test-Path $CustomPath)) {
        return $CustomPath
    }

    $fromPath = Get-Command soffice -ErrorAction SilentlyContinue
    if ($fromPath -and $fromPath.Source -and (Test-Path $fromPath.Source)) {
        return $fromPath.Source
    }

    $candidates = @(
        'C:\Program Files\LibreOffice\program\soffice.exe',
        'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    throw "LibreOffice (soffice) not found. Install LibreOffice or pass -SofficePath manually."
}

$soffice = Resolve-SofficePath -CustomPath $SofficePath

if (-not (Test-Path $InputDir)) {
    throw "Input folder not found: $InputDir"
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

if ($OutputFileName) {
    if ($OutputFileName -notmatch '\.docx$') {
        $OutputFileName = "$OutputFileName.docx"
    }
    $OutputDocx = Join-Path $OutputDir $OutputFileName
} else {
    $OutputDocx = Join-Path $OutputDir ($OutputBaseName + '.docx')
}

$workDir = Join-Path $OutputDir '_work'
if (Test-Path $workDir) { Remove-Item $workDir -Recurse -Force }
New-Item -ItemType Directory -Path $workDir | Out-Null

# PDF source: read directly from $InputDir/*.pdf
$pdfs = Get-ChildItem -Path $InputDir -Filter '*.pdf' |
    Where-Object { $_.Name -notlike 'Generated - *' }

$pdfs = $pdfs | Sort-Object FullName
if (-not $pdfs -or $pdfs.Count -eq 0) {
    throw "No PDF files found. Check input folder: $InputDir"
}

Write-Host "LibreOffice: $soffice"
Write-Host "InputDir: $InputDir"
Write-Host "OutputDir: $OutputDir"
Write-Host "Total PDF files detected: $($pdfs.Count)"

$sections = @()
$idx = 0
foreach ($pdf in $pdfs) {
    $idx++
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pdf.Name)
    $srcRegion = 'INPUT'
    $safeBaseName = ($baseName -replace '[^a-zA-Z0-9\-_.]+', '_')
    $imgOutDir = Join-Path $workDir ("img_{0:D3}_{1}_{2}" -f $idx, $srcRegion, $safeBaseName)
    New-Item -ItemType Directory -Path $imgOutDir -Force | Out-Null

    Write-Host ("[{0}/{1}] [{2}] {3}" -f $idx, $pdfs.Count, $srcRegion, $pdf.Name)

    # Extract service title from file name pattern: "Job <service> - Nomad"
    $service = $baseName
    if ($service -match '^Job\s+(.+?)\s+-\s+Nomad') {
        $service = $Matches[1]
    }

    # Convert PDF -> PNG (1 image per page), dedicated folder per file to avoid overwrite
    & cmd /c """$soffice"" --headless --convert-to png --outdir ""$imgOutDir"" ""$($pdf.FullName)""" | Out-Null

    $converted = Get-ChildItem -Path $imgOutDir -Filter '*.png' | Sort-Object Name

    # Fallback: get latest PNGs in that image folder
    if (-not $converted -or $converted.Count -eq 0) {
        $converted = Get-ChildItem -Path $imgOutDir -Filter '*.png' | Sort-Object LastWriteTime
    }

    $sections += [PSCustomObject]@{
        Service = $service
        Images  = $converted
    }

    if ($converted -and $converted.Count -gt 0) {
        Write-Host ("    -> OK ({0} image)" -f $converted.Count)
    } else {
        Write-Host "    -> WARNING: converted images were not found"
    }
}

$fodtPath = Join-Path $workDir 'report.fodt'
$htmlPath = Join-Path $workDir 'report.html'
$xmlEsc = [System.Security.SecurityElement]::Escape

$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$sb.AppendLine('<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.2">')
[void]$sb.AppendLine('  <office:automatic-styles>')
[void]$sb.AppendLine('    <style:style style:name="PTitle" style:family="paragraph"><style:paragraph-properties fo:margin-bottom="0.4cm"/><style:text-properties fo:font-size="18pt" fo:font-weight="bold"/></style:style>')
[void]$sb.AppendLine('    <style:style style:name="PService" style:family="paragraph"><style:paragraph-properties fo:margin-top="0.3cm" fo:margin-bottom="0.2cm"/><style:text-properties fo:font-size="14pt" fo:font-weight="bold"/></style:style>')
[void]$sb.AppendLine('    <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" style:print-orientation="portrait" fo:margin-top="1cm" fo:margin-bottom="1cm" fo:margin-left="1cm" fo:margin-right="1cm"/></style:page-layout>')
[void]$sb.AppendLine('    <style:master-page style:name="Standard" style:page-layout-name="pm1"/>')
[void]$sb.AppendLine('  </office:automatic-styles>')
[void]$sb.AppendLine('  <office:body><office:text>')
$titleLine = '    <text:p text:style-name="PTitle">' + $xmlEsc.Invoke($ReportTitle) + ' - ' + (Get-Date -Format 'dd MMMM yyyy') + '</text:p>'
[void]$sb.AppendLine($titleLine)

$imgCounter = 1
foreach ($sec in $sections) {
    $serviceLine = '    <text:p text:style-name="PService">' + $xmlEsc.Invoke($sec.Service) + '</text:p>'
    [void]$sb.AppendLine($serviceLine)
    foreach ($img in $sec.Images) {
        $uri = (New-Object System.Uri($img.FullName)).AbsoluteUri
        [void]$sb.AppendLine('    <text:p>')
        $frameLine = '      <draw:frame draw:name="img' + $imgCounter + '" text:anchor-type="paragraph" svg:width="6in" svg:height="8in" draw:z-index="0">'
        [void]$sb.AppendLine($frameLine)
        $imgLine = '        <draw:image xlink:href="' + $uri + '" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>'
        [void]$sb.AppendLine($imgLine)
        [void]$sb.AppendLine('      </draw:frame>')
        [void]$sb.AppendLine('    </text:p>')
        $imgCounter++
    }
    [void]$sb.AppendLine('    <text:p text:style-name="PService"/>')
}

[void]$sb.AppendLine('  </office:text></office:body></office:document>')
[System.IO.File]::WriteAllText($fodtPath, $sb.ToString(), [System.Text.Encoding]::UTF8)

# HTML fallback (if FODT conversion fails)
$h = New-Object System.Text.StringBuilder
[void]$h.AppendLine('<!doctype html><html><head><meta charset="utf-8">')
[void]$h.AppendLine('<style>@page{size:A4 portrait; margin:1cm;} body{font-family:Arial,sans-serif;font-size:11pt;} h2{margin:8px 0;} img{width:6in;height:8in;display:block;object-fit:contain;margin:8px auto;border:1px solid #ddd;}</style>')
[void]$h.AppendLine('</head><body>')
[void]$h.AppendLine('<h1>' + $xmlEsc.Invoke($ReportTitle) + ' - ' + (Get-Date -Format 'dd MMMM yyyy') + '</h1>')
foreach ($sec in $sections) {
    [void]$h.AppendLine('<h2>' + $xmlEsc.Invoke($sec.Service) + '</h2>')
    foreach ($img in $sec.Images) {
        $uri = (New-Object System.Uri($img.FullName)).AbsoluteUri
        [void]$h.AppendLine('<img src="' + $uri + '" />')
    }
}
[void]$h.AppendLine('</body></html>')
[System.IO.File]::WriteAllText($htmlPath, $h.ToString(), [System.Text.Encoding]::UTF8)

$outDir = $OutputDir

# Convert FODT -> DOCX
& cmd /c """$soffice"" --headless --convert-to docx:""MS Word 2007 XML"" --outdir ""$outDir"" ""$fodtPath""" | Out-Null

$generatedDocx = Join-Path $outDir 'report.docx'
if (-not (Test-Path $generatedDocx)) {
    $generatedDocx = Join-Path $outDir 'report.fodt.docx'
}

# Fallback conversion if FODT fails
if (-not (Test-Path $generatedDocx)) {
    & cmd /c """$soffice"" --headless --convert-to odt --outdir ""$outDir"" ""$htmlPath""" | Out-Null
    $htmlOdt = Join-Path $outDir 'report.odt'
    if (Test-Path $htmlOdt) {
        & cmd /c """$soffice"" --headless --convert-to docx:""MS Word 2007 XML"" --outdir ""$outDir"" ""$htmlOdt""" | Out-Null
        if (Test-Path (Join-Path $outDir 'report.docx')) {
            $generatedDocx = Join-Path $outDir 'report.docx'
        }
    }
}

if (-not (Test-Path $generatedDocx)) {
    throw "Failed to create DOCX. Check source files in $InputDir and ensure no LibreOffice document is currently locked/open."
}

if ((Test-Path $OutputDocx) -and (-not $ForceOverwrite)) {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    if ($OutputFileName) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($OutputFileName)
        $OutputDocx = Join-Path $OutputDir ($base + ' - ' + $timestamp + '.docx')
    } else {
        $OutputDocx = Join-Path $OutputDir ($OutputBaseName + ' - ' + $timestamp + '.docx')
    }
}
if (Test-Path $OutputDocx) { Remove-Item $OutputDocx -Force }
Move-Item -Path $generatedDocx -Destination $OutputDocx -Force

if (Test-Path $OutputDocx) {
    Write-Host "Evidence DOCX generated successfully: $OutputDocx"
}
