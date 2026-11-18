# docx2pdf.ps1 - high-fidelity docx -> PDF via Word COM
# Usage: powershell -ExecutionPolicy Bypass -File docx2pdf.ps1 -Src <docx> [-Dst <pdf>]
# Notes: deletes an existing target PDF first (Word may silently skip writing
#        when the target exists); uses Word's own rendering so SimSun / faux-bold /
#        tab stops match the editor exactly. Output messages are ASCII on purpose
#        (Windows PowerShell 5.1 mis-parses non-ASCII in BOM-less .ps1 files).
param(
    [Parameter(Mandatory=$true)][string]$Src,
    [string]$Dst = ""
)
if ($Dst -eq "") { $Dst = [System.IO.Path]::ChangeExtension($Src, ".pdf") }

if (Test-Path $Dst) { Remove-Item $Dst -Force }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($Src, $false, $true)   # ReadOnly
    $doc.ExportAsFixedFormat($Dst, 17)                  # 17 = wdExportFormatPDF
    $pages = $doc.ComputeStatistics(2)                  # 2 = wdStatisticPages
    $doc.Close(0)
    Write-Output ("PDF OK pages=$pages path=$Dst")
} finally {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}
if (Test-Path $Dst) {
    $f = Get-Item $Dst
    Write-Output ("verify size=$($f.Length) mtime=$($f.LastWriteTime)")
} else {
    Write-Output "PDF MISSING"
    exit 1
}
