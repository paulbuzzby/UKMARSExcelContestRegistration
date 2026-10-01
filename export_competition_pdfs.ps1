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
        }
        catch [System.Management.Automation.RuntimeException] {
            # Excel can briefly return a null object instead of a busy COM error.
            if ($attempt -eq $MaxAttempts -or
                (-not $_.Exception.Message.Contains("null-valued expression") -and
                 -not $_.Exception.Message.Contains("Excel returned a null range"))) {
                throw
            }
        }
        Start-Sleep -Milliseconds ($DelayMs * $attempt)
    }
}

function Get-RangeValues {
    param($Worksheet, [string]$Address)

    return Invoke-ExcelCall {
        $range = $null
        try {
            $range = $Worksheet.Range($Address)
            if ($null -eq $range) {
                throw "Excel returned a null range: $Address"
            }
            # Wrap the array so PowerShell does not flatten Excel's two-dimensional values.
            $values = $range.Value2
            if ($null -eq $values) {
                throw "Excel returned a null range value: $Address"
            }
            [PSCustomObject]@{ Values = $values }
        }
        finally {
            if ($null -ne $range) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($range)
            }
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
    $headers = (Get-RangeValues -Worksheet $Worksheet -Address "A4:$((Get-ColumnLetter $usedCols))4").Values
    for ($col = 1; $col -le $usedCols; $col++) {
        $header = if ($usedCols -eq 1) { $headers } else { $headers[1, $col] }
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
    $values = (Get-RangeValues -Worksheet $Worksheet -Address "A${StartRow}:A${EndRow}").Values
    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $value = if ($StartRow -eq $EndRow) { $values } else { $values[($row - $StartRow + 1), 1] }
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $lastRow = $row
        }
    }
    return $lastRow
}

function Get-ColumnLetter {
    param([int]$Column)

    $letter = ""
    while ($Column -gt 0) {
        $Column--
        $letter = [string][char](65 + ($Column % 26)) + $letter
        $Column = [int][Math]::Floor($Column / 26)
    }
    return $letter
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
    Invoke-ExcelCall { $excel.Visible = $false }
    Invoke-ExcelCall { $excel.DisplayAlerts = $false }
    Invoke-ExcelCall { $excel.ScreenUpdating = $false }
    Invoke-ExcelCall { $excel.EnableEvents = $false }

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

        $lastDataRow = Get-LastDataRow -Worksheet $worksheet
        if ($lastDataRow -lt 5) {
            Write-Host "Skipping empty sheet: $sheetName"
            continue
        }

        $lastHeaderColumn = Get-LastHeaderColumn -Worksheet $worksheet

        if ($lastHeaderColumn -lt 1 -or $lastDataRow -lt 5) {
            Write-Host "Skipping sheet with no printable area: $sheetName"
            continue
        }

        $printArea = '$A$1:$' + (Get-ColumnLetter $lastHeaderColumn) + '$' + $lastDataRow

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
        Invoke-ExcelCall { $workbook.Close($false) }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        Invoke-ExcelCall { $excel.Quit() }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
