$inputFolder = "D:\PPTs\input"
$outputFolder = "D:\PPTs\output"

# Create output folder if it doesn't exist
if (!(Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

$pptApp = New-Object -ComObject PowerPoint.Application
$pptApp.WindowState = 2   # Minimize PowerPoint safely

$files = Get-ChildItem $inputFolder -Filter *.pptx
$total = $files.Count
$count = 0

foreach ($file in $files) {
    $count++
    Write-Host "[$count / $total] Converting: $($file.Name)"

    $presentation = $pptApp.Presentations.Open(
        $file.FullName,
        $false,   # ReadOnly
        $false,   # Untitled
        $false    # WithWindow
    )

    $pdfPath = Join-Path $outputFolder ($file.BaseName + ".pdf")
    $presentation.SaveAs($pdfPath, 32)   # 32 = PDF format
    $presentation.Close()
}

$pptApp.Quit()

# Completion sound
[console]::beep(1000, 600)

Write-Host "`n Conversion complete! $total file(s) converted." -ForegroundColor Green

# Keep window open so you see the message
Read-Host "Press Enter to close"
