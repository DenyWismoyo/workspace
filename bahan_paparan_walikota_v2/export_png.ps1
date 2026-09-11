$pptxPath = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v2\PAPARAN_ROADMAP_STP_WALIKOTA_MASTER.pptx"
$previewDir = "D:\Project\GAWE\Paparan Roadmap STP\bahan_paparan_walikota_v2\preview_slides"

if (!(Test-Path $previewDir)) {
    New-Item -ItemType Directory -Path $previewDir | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

# Export all slides as PNG
$pres.SaveAs($previewDir, 18)
$pres.Close()
$ppt.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null

Get-ChildItem $previewDir | Select-Object Name, Length
