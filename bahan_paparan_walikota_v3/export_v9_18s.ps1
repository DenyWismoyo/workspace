$pptxPath = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V9_18S.pptx"
$pdfPath1 = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V9_18S.pdf"
$pdfPath2 = "D:\Project\GAWE\Paparan Roadmap STP\output_paparan\PAPARAN_WALIKOTA_STP_V9_18S.pdf"
$outDir = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\slides_18s_preview"

if (!(Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$pptApp = New-Object -ComObject PowerPoint.Application
try {
    $presentation = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    # Save PDF
    $presentation.SaveAs($pdfPath1, 32)
    $presentation.SaveAs($pdfPath2, 32)
    Write-Host "PDF 18S successfully exported to $pdfPath1 and $pdfPath2"
    
    # Save PNGs
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $outFile = Join-Path $outDir ("slide_18s_" + $i.ToString("D2") + ".png")
        $presentation.Slides.Item($i).Export($outFile, "PNG", 1920, 1080)
        Write-Host "Exported Slide $i to $outFile"
    }
    $presentation.Close()
}
finally {
    $pptApp.Quit()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
Write-Host "All 18S exports completed successfully!"
