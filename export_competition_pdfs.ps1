param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\\competition-pdfs"
)

$ErrorActionPreference = "Stop"

function Invoke-ExcelCall {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action,

        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = 12,

        [Parameter(Mandatory = $false)]
        [int]$DelayMs = 250
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $Action
        }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq $MaxAttempts) {
                throw
            }
            Start-Sleep -Milliseconds ($DelayMs * $attempt)
        }
    }
}

function Get-SafeFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $safe = $Name
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($char, "_")
    }
    return $safe
}

function Get-LastHeaderColumn {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet
    )

    $lastCol = 0
    $usedCols = [Math]::Max((Invoke-ExcelCall { $Worksheet.UsedRange.Columns.Count }), 1)
    for ($col = 1; $col -le $usedCols; $col++) {
        $header = Invoke-ExcelCall { $Worksheet.Cells.Item(4, $col).Text }
        $isHidden = Invoke-ExcelCall { $Worksheet.Columns.Item($col).Hidden }
        if (-not $isHidden -and -not [string]::IsNullOrWhiteSpace($header)) {
            $lastCol = $col
        }
    }
    return $lastCol
}

function Get-LastDataRow {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $false)]
        [int]$StartRow = 5,

        [Parameter(Mandatory = $false)]
        [int]$EndRow = 1004
    )

    $lastRow = 0
    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $value = Invoke-ExcelCall { $Worksheet.Cells.Item($row, 1).Text }
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $lastRow = $row
        }
    }
    return $lastRow
}

$competitionSheets = @(
    "Junior Half Size Line Follower",
    "Senior Half Size Line Follower",
    "Junior Half Size Wall Follower",
    "Senior Half Size Wall Follower",
    "Junior Wall Follower",
    "Senior Wall Follower",
    "Junior Maze Solver",
    "Senior Maze Solver",
    "Junior Half Size Maze Solver",
    "Senior Half Size Maze Solver",
    "Junior Drag Race",
    "Senior Drag Race",
    "Junior Pursuit Race",
    "Senior Pursuit Race"
)

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$resolvedOutputFolder = (Resolve-Path -LiteralPath $OutputFolder).Path

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false

    $workbook = Invoke-ExcelCall { $excel.Workbooks.Open($resolvedWorkbookPath, 0, $true) }

    Invoke-ExcelCall { $excel.CalculateFullRebuild() }

    $availableSheetNames = @()
    foreach ($sheet in $workbook.Worksheets) {
        $availableSheetNames += $sheet.Name
    }

    foreach ($sheetName in $competitionSheets) {
        if ($availableSheetNames -notcontains $sheetName) {
            Write-Warning "Sheet not found: $sheetName"
            continue
        }

        $worksheet = $null
        try {
            $worksheet = Invoke-ExcelCall { $workbook.Worksheets.Item($sheetName) }
        } catch {
            Write-Warning "Unable to access sheet: $sheetName"
            continue
        }

        Invoke-ExcelCall { $worksheet.Calculate() }

        $firstRobot = Invoke-ExcelCall { $worksheet.Range("A5").Text }
        if ([string]::IsNullOrWhiteSpace($firstRobot)) {
            Write-Host "Skipping empty sheet: $sheetName"
            continue
        }

        $lastHeaderColumn = Get-LastHeaderColumn -Worksheet $worksheet
        $lastDataRow = Get-LastDataRow -Worksheet $worksheet

        if ($lastHeaderColumn -lt 1 -or $lastDataRow -lt 5) {
            Write-Host "Skipping sheet with no printable area: $sheetName"
            continue
        }

        $bottomRightAddress = Invoke-ExcelCall { $worksheet.Cells.Item($lastDataRow, $lastHeaderColumn).Address() }
        $printArea = '$A$1:' + $bottomRightAddress

        Invoke-ExcelCall { $worksheet.PageSetup.PrintArea = $printArea }
        Invoke-ExcelCall { $worksheet.PageSetup.Orientation = 2 }
        Invoke-ExcelCall { $worksheet.PageSetup.PaperSize = 9 }
        Invoke-ExcelCall { $worksheet.PageSetup.Zoom = $false }
        Invoke-ExcelCall { $worksheet.PageSetup.FitToPagesWide = 1 }
        Invoke-ExcelCall { $worksheet.PageSetup.FitToPagesTall = $false }
        Invoke-ExcelCall { $worksheet.PageSetup.CenterHorizontally = $true }
        Invoke-ExcelCall { $worksheet.PageSetup.CenterVertically = $false }

        $pdfName = "{0}.pdf" -f (Get-SafeFileName -Name $sheetName)
        $pdfPath = Join-Path $resolvedOutputFolder $pdfName

        Invoke-ExcelCall { $worksheet.ExportAsFixedFormat(0, $pdfPath) }
        Write-Host "Exported: $pdfPath"
    }
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
