$pptxPath = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V3.pptx"
$pdfPath = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V3.pdf"
$destPdf = "D:\Project\GAWE\Paparan Roadmap STP\output_paparan\PAPARAN_WALIKOTA_STP_V3.pdf"

$pptApp = New-Object -ComObject PowerPoint.Application
try {
    $presentation = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $presentation.SaveAs($pdfPath, 32)
    $presentation.Close()
    Copy-Item $pdfPath $destPdf -Force
    Write-Host "[SUCCESS] PDF Master V3 exported to: $pdfPath"
    Write-Host "[SUCCESS] Synced to: $destPdf"
}
finally {
    $pptApp.Quit()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
