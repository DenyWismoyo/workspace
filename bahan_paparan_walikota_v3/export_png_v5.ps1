$pptxPath = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\PAPARAN_WALIKOTA_STP_V5.pptx"
$outDir = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v3\slides_v5_preview"

if (!(Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$pptApp = New-Object -ComObject PowerPoint.Application
try {
    $presentation = $pptApp.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    for ($i = 1; $i -le $presentation.Slides.Count; $i++) {
        $outFile = Join-Path $outDir ("slide_v5_" + $i.ToString("D2") + ".png")
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
