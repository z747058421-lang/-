# check_pages.ps1 — 用 Word COM 计算 docx 页数（被 build_lesson_plan.py 调用）
param(
    [Parameter(Mandatory=$true)]
    [string]$Path
)
try {
    # 确保路径是绝对路径且存在
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not (Test-Path $fullPath)) {
        Write-Output "ERR: file not found: $fullPath"
        exit 1
    }
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $doc = $w.Documents.Open($fullPath)
    $pages = $doc.ComputeStatistics(2)  # wdStatisticPages = 2
    $doc.Close([ref]$false)
    $w.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
    Write-Output $pages
} catch {
    Write-Output "ERR:$($_.Exception.Message)"
    try {
        if ($doc) { $doc.Close([ref]$false) }
        if ($w) { $w.Quit() }
    } catch {}
}
