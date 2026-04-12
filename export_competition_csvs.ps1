param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = ".\\competition-csvs"
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

function Get-LastDataRow {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $false)]
        [int]$StartRow = 5,

        [Parameter(Mandatory = $false)]
    [int]$EndRow = 1004
    )

    $candidate = Invoke-ExcelCall { $Worksheet.Cells.Item($Worksheet.Rows.Count, 1).End(-4162).Row }
    if ($candidate -lt $StartRow) {
        return 0
    }
    return [Math]::Min($candidate, $EndRow)
}

function Get-CellText {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$Row,

        [Parameter(Mandatory = $true)]
        [int]$Column
    )

    for ($attempt = 1; $attempt -le 12; $attempt++) {
        try {
            $cell = $Worksheet.Cells.Item($Row, $Column)
            if ($null -ne $cell) {
                return [string]$cell.Text
            }
        }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq 12) {
                throw
            }
        }
        catch [System.Management.Automation.RuntimeException] {
            if ($attempt -eq 12 -or -not $_.Exception.Message.Contains("null-valued expression")) {
                throw
            }
        }

        Start-Sleep -Milliseconds (250 * $attempt)
    }

    return ""
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

        $lastDataRow = Get-LastDataRow -Worksheet $worksheet
        if ($lastDataRow -lt 5) {
            Write-Host "Skipping sheet with no data rows: $sheetName"
            continue
        }

        $rows = New-Object System.Collections.Generic.List[object]
        for ($row = 4; $row -le $lastDataRow; $row++) {
            $robot = Get-CellText -Worksheet $worksheet -Row $row -Column 1
            $builder = Get-CellText -Worksheet $worksheet -Row $row -Column 2

            if ($row -gt 4 -and [string]::IsNullOrWhiteSpace($robot)) {
                continue
            }

            $rows.Add([PSCustomObject]@{
                'Robot Name' = $robot
                'Builder Name' = $builder
            }) | Out-Null
        }

        $csvName = "{0}.csv" -f (Get-SafeFileName -Name $sheetName)
        $csvPath = Join-Path $resolvedOutputFolder $csvName
        $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported: $csvPath"
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
