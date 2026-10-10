param(
    [Parameter(Mandatory = $true)][string]$Pptx,
    [Parameter(Mandatory = $true)][string]$OutputDir
)
$ErrorActionPreference = 'Stop'
function Release-Com($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null } catch {}
    }
}
$renderDirectory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDir)
New-Item -ItemType Directory -Path $renderDirectory -Force | Out-Null
$report = [ordered]@{
    status = 'unavailable'; error = $null; tolerance_pt = 1
    summary = [ordered]@{ rendered_slides = 0; text_shapes_checked = 0; overflow_count = 0 }
    images = @(); text_bounds = @(); overflow = @()
}
$application = $null; $presentation = $null; $createdApplication = $false
try {
    $inputPath = (Resolve-Path -LiteralPath $Pptx).Path
    $report.pptx_sha256 = (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $existingPowerPoint = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
    try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
    catch {
        $application = New-Object -ComObject PowerPoint.Application
        $createdApplication = -not $existingPowerPoint
    }
    # ReadOnly = true; Untitled = false; WithWindow = false. Never hide the user's app.
    $presentation = $application.Presentations.Open($inputPath, -1, 0, 0)
    $height = [int][Math]::Round(1600 * $presentation.PageSetup.SlideHeight / $presentation.PageSetup.SlideWidth)
    for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
        $slide = $null
        try {
            $slide = $presentation.Slides.Item($slideIndex)
            $imagePath = Join-Path $renderDirectory ('slide_{0:D2}.png' -f $slideIndex)
            $slide.Export($imagePath, 'PNG', 1600, $height)
            $report.images += [ordered]@{ slide = $slideIndex; path = $imagePath; width = 1600; height = $height }
            $report.summary.rendered_slides++
            for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                $shape = $null; $frame = $null; $range = $null
                try {
                    $shape = $slide.Shapes.Item($shapeIndex)
                    if ($shape.Visible -eq 0 -or $shape.HasTextFrame -eq 0) { continue }
                    $frame = $shape.TextFrame2
                    if ($frame.HasText -eq 0) { continue }
                    $range = $frame.TextRange
                    if ([string]::IsNullOrWhiteSpace($range.Text)) { continue }
                    $shapeBounds = @([double]$shape.Left, [double]$shape.Top, [double]$shape.Width, [double]$shape.Height)
                    $textBounds = @([double]$range.BoundLeft, [double]$range.BoundTop, [double]$range.BoundWidth, [double]$range.BoundHeight)
                    $item = [ordered]@{
                        slide = $slideIndex; slidename = [string]$slide.Name; textboundshape = [string]$shape.Name
                        text = [string]$range.Text; bounds = $textBounds; shape_bounds = $shapeBounds
                    }
                    $report.text_bounds += $item
                    $report.summary.text_shapes_checked++
                    if ($textBounds[0] -lt ($shapeBounds[0] - 1) -or $textBounds[1] -lt ($shapeBounds[1] - 1) -or
                        ($textBounds[0] + $textBounds[2]) -gt ($shapeBounds[0] + $shapeBounds[2] + 1) -or
                        ($textBounds[1] + $textBounds[3]) -gt ($shapeBounds[1] + $shapeBounds[3] + 1)) {
                        $report.overflow += $item
                    }
                } finally { Release-Com $range; Release-Com $frame; Release-Com $shape }
            }
        } finally { Release-Com $slide }
    }
    $report.summary.overflow_count = $report.overflow.Count
    $report.status = if ($report.overflow.Count -eq 0) { 'passed' } else { 'failed' }
} catch {
    $report.status = if ($report.summary.rendered_slides -gt 0) { 'failed' } else { 'unavailable' }
    $report.error = 'PowerPoint render/text-bound check could not complete: ' + $_.Exception.Message
} finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { $report.status = 'failed'; $report.error = $_.Exception.Message }
        Release-Com $presentation
    }
    if ($createdApplication -and $null -ne $application) {
        try { if ($application.Presentations.Count -eq 0) { $application.Quit() } } catch {}
    }
    Release-Com $application
}
$reportPath = Join-Path $renderDirectory 'render_report.json'
$report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding UTF8
Write-Output ([ordered]@{ status = $report.status; error = $report.error; summary = $report.summary; report = $reportPath } | ConvertTo-Json -Depth 4 -Compress)
if ($report.status -ne 'passed') { exit 1 }
