param(
    [string]$pptxPath = "d:\Project\GAWE\Paparan Roadmap STP\output_paparan\PAPARAN_ROADMAP_STP_WALIKOTA.pptx",
    [string]$pdfPath = "d:\Project\GAWE\Paparan Roadmap STP\output_paparan\PAPARAN_ROADMAP_STP_WALIKOTA.pdf"
)

try {
    Write-Host "⏳ Membuka PowerPoint untuk konversi PDF..."
    $ppt = New-Object -ComObject PowerPoint.Application
    
    # Buka presentasi (ReadOnly, Untitled, WithWindow = False)
    $pres = $ppt.Presentations.Open($pptxPath, [Microsoft.Office.Core.MsoTriState]::msoTrue, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
    
    # 32 = ppSaveAsPDF
    $pres.SaveAs($pdfPath, 32)
    $pres.Close()
    $ppt.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    
    Write-Host "✅ Berhasil mengekspor PDF: $pdfPath"
} catch {
    Write-Error "❌ Gagal mengonversi ke PDF: $_"
    exit 1
}
