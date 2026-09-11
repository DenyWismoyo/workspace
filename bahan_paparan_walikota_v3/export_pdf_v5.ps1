$pptxPath = "d:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V5.pptx"
$pdfPath1 = "d:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V5.pdf"
$pdfPath2 = "d:\Project\GAWE\Paparan Roadmap STP\output_paparan\PAPARAN_WALIKOTA_STP_V5.pdf"

$pptApp = New-Object -ComObject PowerPoint.Application
$presentation = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

# ppSaveAsPDF = 32
$presentation.SaveAs($pdfPath1, 32)
$presentation.SaveAs($pdfPath2, 32)

$presentation.Close()
$pptApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($pptApp) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "PDF V5 successfully exported!"
