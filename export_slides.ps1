$pptxPath = "D:\AnToanvaBaoMatHeThongThongTin\GK\docs\Bao_Cao_Giua_Ky_DT05_Final.pptx"
$pdfPath = "D:\AnToanvaBaoMatHeThongThongTin\GK\docs\Bao_Cao_Giua_Ky_DT05_Final.pdf"
$imgDir = "D:\AnToanvaBaoMatHeThongThongTin\GK\docs\slide_images"

if (!(Test-Path $imgDir)) {
    New-Item -ItemType Directory -Path $imgDir -Force | Out-Null
}

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($pptxPath, 2, 0, 0) # 2 = ReadOnly, 0 = Untitled, 0 = WithWindow
    
    # Export to PDF
    Write-Output "Exporting to PDF: $pdfPath"
    $pres.SaveAs($pdfPath, 32) # 32 = ppSaveAsPDF
    
    # Export each slide to PNG
    $slideCount = $pres.Slides.Count
    Write-Output "Exporting $slideCount slides to PNG..."
    for ($i = 1; $i -le $slideCount; $i++) {
        $slide = $pres.Slides.Item($i)
        $imgPath = Join-Path $imgDir "slide_$($i.ToString('00')).png"
        $slide.Export($imgPath, "PNG", 1920, 1080)
        Write-Output "  Exported Slide $i -> $imgPath"
    }

    $pres.Close()
    $ppt.Quit()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Output "SUCCESS: Exported all slides and PDF."
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
