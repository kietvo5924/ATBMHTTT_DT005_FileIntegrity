$pptxPath = "D:\AnToanvaBaoMatHeThongThongTin\GK\docs\Bao_Cao_Giua_Ky_DT05_Final.pptx"
$pdfPath = "D:\AnToanvaBaoMatHeThongThongTin\GK\docs\Bao_Cao_Giua_Ky_DT05_Final.pdf"
$rootPdf = "D:\AnToanvaBaoMatHeThongThongTin\GK\Bao_Cao_Giua_Ky_DT05_Final.pdf"
$imgDir = "D:\AnToanvaBaoMatHeThongThongTin\GK\docs\slide_images"

try {
    Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | Stop-Process -Force
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($pptxPath, 2, 0, 0) # 2 = ReadOnly, 0 = Untitled, 0 = WithWindow
    
    # Export to PDF
    Write-Output "Exporting to PDF: $pdfPath"
    $pres.SaveAs($pdfPath, 32) # 32 = ppSaveAsPDF
    
    $pres.Close()
    $ppt.Quit()
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    
    Copy-Item -Path $pdfPath -Destination $rootPdf -Force
    Write-Output "Copied PDF to: $rootPdf"
    
    if (Test-Path $imgDir) {
        Remove-Item -Path $imgDir -Recurse -Force
        Write-Output "Cleaned up temporary slide_images folder."
    }
    
    Write-Output "SUCCESS: Exported and synchronized PDF."
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
}
