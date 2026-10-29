$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$docxPath = Join-Path $taskRoot 'outputs\实习证明.docx'
$pdfPath = Join-Path $taskRoot 'work\certificate_render\certificate.pdf'
$wordApp = $null
$wordDoc = $null
try {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordDoc = $wordApp.Documents.Open($docxPath, $false, $true)
    $wordDoc.ExportAsFixedFormat($pdfPath, 17)
    Write-Output $pdfPath
} finally {
    if ($null -ne $wordDoc) { $wordDoc.Close(0) }
    if ($null -ne $wordApp) { $wordApp.Quit() }
}
