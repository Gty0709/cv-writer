<#
.SYNOPSIS
    Compiles a structured CV into Word (.docx) and exports to PDF (.pdf) with page budget verification.
.DESCRIPTION
    Takes the base OpenXML template, inserts structured paragraphs adhering to the studio's exact
    margins, tab stops, and borders, updates the docx package, opens it in Word COM to verify page count,
    and exports a pristine PDF.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$TemplatePath = "$PSScriptRoot\..\templates\template.docx",

    [Parameter(Mandatory=$true)]
    [string]$OutputDocx,

    [Parameter(Mandatory=$false)]
    [string]$OutputPdf,

    [Parameter(Mandatory=$true)]
    [string]$XmlBodyPath
)

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

if (-not (Test-Path $TemplatePath)) {
    Write-Error "Template not found at $TemplatePath"
    exit 1
}

if (-not $OutputPdf) {
    $OutputPdf = [System.IO.Path]::ChangeExtension($OutputDocx, ".pdf")
}

# Ensure parent directory exists
$parent = Split-Path $OutputDocx -Parent
if ($parent -and -not (Test-Path $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}

# 1. Copy base template
Copy-Item -Path $TemplatePath -Destination $OutputDocx -Force

# 2. Read body content
$bodyContent = [System.IO.File]::ReadAllText($XmlBodyPath, [System.Text.Encoding]::UTF8)

# 3. Assemble document.xml
$sectPr = "<w:sectPr><w:pgSz w:w=`"11909`" w:h=`"16834`" /><w:pgMar w:top=`"567`" w:right=`"1150`" w:bottom=`"567`" w:left=`"850`" w:header=`"708`" w:footer=`"708`" w:gutter=`"0`" /><w:cols w:space=`"720`" /><w:docGrid w:linePitch=`"360`" /></w:sectPr>"
$docXml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><w:document xmlns:wpc=`"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas`" xmlns:mc=`"http://schemas.openxmlformats.org/markup-compatibility/2006`" xmlns:o=`"urn:schemas-microsoft-com:office:office`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`" xmlns:m=`"http://schemas.openxmlformats.org/officeDocument/2006/math`" xmlns:v=`"urn:schemas-microsoft-com:vml`" xmlns:wp14=`"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing`" xmlns:wp=`"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing`" xmlns:w10=`"urn:schemas-microsoft-com:office:word`" xmlns:w=`"http://schemas.openxmlformats.org/wordprocessingml/2006/main`" xmlns:w14=`"http://schemas.microsoft.com/office/word/2010/wordml`" xmlns:wpg=`"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup`" xmlns:wpi=`"http://schemas.microsoft.com/office/word/2010/wordprocessingInk`" xmlns:wne=`"http://schemas.microsoft.com/office/word/2006/wordml`" xmlns:wps=`"http://schemas.microsoft.com/office/word/2010/wordprocessingShape`"><w:body>$bodyContent$sectPr</w:body></w:document>"

# 4. Inject into DOCX
$fs = [System.IO.File]::Open($OutputDocx, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)
$entry = $zip.GetEntry("word/document.xml")
if ($entry -ne $null) { $entry.Delete() }
$newEntry = $zip.CreateEntry("word/document.xml")
$writer = New-Object System.IO.StreamWriter($newEntry.Open(), [System.Text.Encoding]::UTF8)
$writer.Write($docXml)
$writer.Close()
$zip.Dispose()
$fs.Close()

Write-Output "Successfully updated Word document: $OutputDocx"

# 5. Export PDF & verify page count via Word COM
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($OutputDocx, $false, $true)
    
    $pages = $doc.ComputeStatistics(2)
    Write-Output "Document Page Count: $pages"
    
    # 17 = wdFormatPDF
    $doc.SaveAs([ref]$OutputPdf, [ref]17)
    Write-Output "Successfully exported PDF: $OutputPdf"
    
    $doc.Close($false)
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
} catch {
    Write-Warning "Word COM export failed or timed out: $($_.Exception.Message)"
}
