# Converts every .docx in a Windows folder to PDF with the installed Microsoft Word, to compare with DotOffice's output.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File word-to-pdf.ps1 -InputDir C:\Temp\dotoffice-check\in -OutputDir C:\Temp\dotoffice-check\out
#
# Starts its own hidden Word and closes only that one. Documents are opened read-only with macros disabled and are never saved.
# Existing PDFs of the same name in OutputDir are overwritten.
param(
    [Parameter(Mandatory = $true)][string]$InputDir,
    [Parameter(Mandatory = $true)][string]$OutputDir
)

$ErrorActionPreference = 'Stop'
$wdExportFormatPDF = 17
$wdDoNotSaveChanges = 0
$msoAutomationSecurityForceDisable = 3

$inputPath = (Resolve-Path -LiteralPath $InputDir).Path
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null
$outputPath = (Resolve-Path -LiteralPath $OutputDir).Path

# "~$" files are Word's lock files, not documents.
$files = @(Get-ChildItem -LiteralPath $inputPath -Filter *.docx -File | Where-Object { -not $_.Name.StartsWith('~$') })
if ($files.Count -eq 0) { Write-Output "No .docx files in $inputPath"; exit 0 }

$failed = 0
$word = New-Object -ComObject Word.Application
try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = $msoAutomationSecurityForceDisable
    Write-Output "Word version $($word.Version), build $($word.Build)"

    foreach ($file in $files) {
        $pdf = Join-Path $outputPath ($file.BaseName + '.pdf')
        $doc = $null
        try {
            # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
            $doc = $word.Documents.Open($file.FullName, $false, $true, $false)
            $doc.ExportAsFixedFormat($pdf, $wdExportFormatPDF)
            Write-Output "OK     $($file.Name)"
        }
        catch {
            $failed++
            Write-Output "FAILED $($file.Name): $($_.Exception.Message)"
        }
        finally {
            if ($null -ne $doc) {
                $doc.Close([ref]$wdDoNotSaveChanges)
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
            }
        }
    }
}
finally {
    $word.Quit([ref]$wdDoNotSaveChanges)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($failed -gt 0) { exit 1 }
