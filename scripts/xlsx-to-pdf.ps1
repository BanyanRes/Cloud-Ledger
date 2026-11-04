# CloudLedger - xlsx -> pdf via Excel COM
#
# Converts worksheets to PDF with Excel itself, so the output matches a manual
# "Save as PDF": the workbook's own print area, landscape and fit-to-one-page
# settings are honored (IgnorePrintAreas = $false). Used by the Time Allocation
# report on Jimmy's PC (server/timeallocation.js).
#
# Usage:  powershell -NoProfile -ExecutionPolicy Bypass -File xlsx-to-pdf.ps1 -PairsFile <tsv>
#   where <tsv> is a UTF-8 text file with one  "<xlsx path><TAB><pdf path>"  per line.
#   (TSV, not JSON: Windows PowerShell 5.1's ConvertFrom-Json collapses an array of
#    same-shaped objects into one object with array-valued properties.)
# Prints one "OK<TAB><pdf>" line per converted file; exits non-zero on failure.

param(
  [Parameter(Mandatory = $true)][string]$PairsFile
)
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $PairsFile)) { Write-Error "pairs file not found: $PairsFile"; exit 2 }
$pairs = @()
foreach ($line in (Get-Content -LiteralPath $PairsFile)) {
  if (-not $line -or -not $line.Trim()) { continue }
  $parts = $line -split "`t", 2
  if ($parts.Count -lt 2) { continue }
  $pairs += [pscustomobject]@{ xlsx = $parts[0]; pdf = $parts[1] }
}
if (-not $pairs -or $pairs.Count -eq 0) { Write-Output 'no pairs'; exit 0 }

$xlTypePDF = 0        # xlFixedFormatType.xlTypePDF
$xlQualityStandard = 0

$excel = $null
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.ScreenUpdating = $false
  $excel.AskToUpdateLinks = $false
  $excel.EnableEvents = $false
} catch {
  Write-Error "Excel is not available: $($_.Exception.Message)"; exit 3
}

$failed = 0
try {
  foreach ($p in $pairs) {
    $xlsx = [string]$p.xlsx
    $pdf = [string]$p.pdf
    if (-not (Test-Path -LiteralPath $xlsx)) { Write-Output ("MISSING`t" + $xlsx); $failed++; continue }
    $dir = Split-Path -Parent $pdf
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if (Test-Path -LiteralPath $pdf) { Remove-Item -LiteralPath $pdf -Force -ErrorAction SilentlyContinue }
    $wb = $null
    try {
      # Open(Filename, UpdateLinks=0, ReadOnly=$true)
      $wb = $excel.Workbooks.Open($xlsx, 0, $true)
      try { $excel.CalculateFull() } catch {}   # resolve any carried-forward formulas (e.g. Approval Date = Signature Date)
      # ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties=$true, IgnorePrintAreas=$false)
      $wb.ExportAsFixedFormat($xlTypePDF, $pdf, $xlQualityStandard, $true, $false)
      if (Test-Path -LiteralPath $pdf) { Write-Output ("OK`t" + $pdf) } else { Write-Output ("NOFILE`t" + $pdf); $failed++ }
    } catch {
      Write-Output ("ERROR`t" + $pdf + "`t" + $_.Exception.Message); $failed++
    } finally {
      if ($wb) { try { $wb.Close($false) } catch {}; try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null } catch {} }
    }
  }
} finally {
  if ($excel) {
    try { $excel.Quit() } catch {}
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($failed -gt 0) { exit 1 }
exit 0
