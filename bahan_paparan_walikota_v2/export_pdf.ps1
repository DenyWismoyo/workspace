$pptxPath = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v2\PAPARAN_ROADMAP_STP_WALIKOTA_MASTER.pptx"
$pdfPath = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v2\PAPARAN_ROADMAP_STP_WALIKOTA_MASTER.pdf"
$destPdf = "D:\Project\GAWE\Paparan Roadmap STP\output_paparan\PAPARAN_ROADMAP_STP_WALIKOTA_MASTER.pdf"

$pptApp = New-Object -ComObject PowerPoint.Application
try {
    $presentation = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    $presentation.SaveAs($pdfPath, 32)
    $presentation.Close()
    Copy-Item $pdfPath $destPdf -Force
    Write-Host "[SUCCESS] PDF Master exported to: $pdfPath"
    Write-Host "[SUCCESS] Synced to: $destPdf"
}
finally {
    $pptApp.Quit()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
