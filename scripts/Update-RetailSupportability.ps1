[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceWorkbookPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationWorkbookPath,

    [datetime]$AsOfDate = (Get-Date)
)

$ErrorActionPreference = 'Stop'
$targetStartRow = 4
$soldToParty = '5395802'

function Release-ComObject {
    param([AllowNull()][object]$ComObject)
    if ($null -ne $ComObject -and
        [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
            $ComObject
        )
    }
}

function Register-RetailExcelComFilter {
    if (-not ('RetailOleMessageFilter' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IRetailOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(
        int callType,
        IntPtr taskCaller,
        int tickCount,
        IntPtr interfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(
        IntPtr taskCallee,
        int tickCount,
        int rejectType);

    [PreserveSig]
    int MessagePending(
        IntPtr taskCallee,
        int tickCount,
        int pendingType);
}

public sealed class RetailOleMessageFilter : IRetailOleMessageFilter
{
    private static RetailOleMessageFilter current;
    private static IRetailOleMessageFilter previous;

    [DllImport("Ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        IRetailOleMessageFilter newFilter,
        out IRetailOleMessageFilter oldFilter);

    public static void Register()
    {
        if (current != null)
        {
            return;
        }

        current = new RetailOleMessageFilter();
        IRetailOleMessageFilter oldFilter;
        CoRegisterMessageFilter(current, out oldFilter);
        previous = oldFilter;
    }

    public static void Revoke()
    {
        IRetailOleMessageFilter ignored;
        CoRegisterMessageFilter(previous, out ignored);
        previous = null;
        current = null;
    }

    public int HandleInComingCall(
        int callType,
        IntPtr taskCaller,
        int tickCount,
        IntPtr interfaceInfo)
    {
        return 0;
    }

    public int RetryRejectedCall(
        IntPtr taskCallee,
        int tickCount,
        int rejectType)
    {
        bool retryable = rejectType == 1 || rejectType == 2;
        if (retryable && tickCount <= 9750)
        {
            return 250;
        }

        return -1;
    }

    public int MessagePending(
        IntPtr taskCallee,
        int tickCount,
        int pendingType)
    {
        return 2;
    }
}
'@
    }

    [RetailOleMessageFilter]::Register()
}

function Invoke-ExcelComRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,

        [Parameter(Mandatory = $true)]
        [string]$Context,

        [int]$TimeoutSeconds = 10
    )

    $deadline = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastComException = $null
    while ($true) {
        if ($null -ne $lastComException -and
            [datetime]::UtcNow -ge $deadline) {
            $hresultText = '0x{0:X8}' -f
                ($lastComException.HResult -band 0xFFFFFFFFL)
            throw "Excel COM operation '$Context' failed after $TimeoutSeconds seconds ($hresultText): $($lastComException.Message)"
        }

        try {
            return & $Operation
        }
        catch [Runtime.InteropServices.COMException] {
            $lastComException = $_.Exception
            $hresult = $_.Exception.HResult -band 0xFFFFFFFFL
            $retryable = $hresult -in @(
                0x80010001L,
                0x8001010AL,
                0x800AC472L
            )
            if (-not $retryable) {
                throw
            }
            $remainingMilliseconds =
                ($deadline - [datetime]::UtcNow).TotalMilliseconds
            if ($remainingMilliseconds -le 250) {
                $hresultText = '0x{0:X8}' -f $hresult
                throw "Excel COM operation '$Context' failed after $TimeoutSeconds seconds ($hresultText): $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds 250
        }
    }
}

function Get-Key {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim().ToUpperInvariant()
}

function Get-SoldTo {
    param([AllowNull()][object]$Value)
    $text = Get-Key -Value $Value
    if ($text -match '^\d+$') { return $text.TrimStart('0') }
    return $text
}

function ConvertTo-Number {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return [double]0
    }
    if ($Value -is [datetime]) {
        throw "Expected a numeric value for $Context, but found a date."
    }
    if ($Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        return [double]$Value
    }

    $normalized = ([string]$Value).Trim().
        Replace([char]0x00A0, ' ').Replace(' ', '')
    if ($normalized.StartsWith('(') -and $normalized.EndsWith(')')) {
        $normalized = '-' + $normalized.Substring(1, $normalized.Length - 2)
    }
    $number = [double]0
    $styles = [Globalization.NumberStyles]::Float -bor
        [Globalization.NumberStyles]::AllowThousands
    if (-not [double]::TryParse(
        $normalized,
        $styles,
        [Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )) {
        throw "Expected a numeric value for $Context, but found '$Value'."
    }
    return $number
}

function ConvertTo-Date {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Context
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }
    if ($Value -is [datetime]) { return ([datetime]$Value).Date }
    if ($Value -is [ValueType]) {
        return [datetime]::FromOADate([double]$Value).Date
    }
    $date = [datetime]::MinValue
    if ([datetime]::TryParse(
        ([string]$Value).Trim(),
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None,
        [ref]$date
    )) {
        return $date.Date
    }
    throw "Expected a date for $Context, but found '$Value'."
}

function Add-Total {
    param(
        [Parameter(Mandatory = $true)][Collections.Hashtable]$Map,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][double]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    if (-not $Map.ContainsKey($Key)) { $Map[$Key] = [double]0 }
    $Map[$Key] = [double]$Map[$Key] + $Value
}

function Get-MapValue {
    param(
        [Parameter(Mandatory = $true)][Collections.Hashtable]$Map,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Key
    )
    if ($Map.ContainsKey($Key)) { return [double]$Map[$Key] }
    return [double]0
}

function New-Map {
    return [Collections.Hashtable]::new(
        [StringComparer]::OrdinalIgnoreCase
    )
}

function Get-Friday {
    param([Parameter(Mandatory = $true)][datetime]$Date)
    $daysToFriday = (([int][DayOfWeek]::Friday - [int]$Date.DayOfWeek) + 7) % 7
    return $Date.Date.AddDays($daysToFriday)
}

function Get-RefreshFriday {
    param([Parameter(Mandatory = $true)][datetime]$RunDate)
    $daysSinceMonday = (([int]$RunDate.DayOfWeek + 6) % 7)
    return $RunDate.Date.AddDays(-$daysSinceMonday + 4)
}

function Get-WeeklyActualRange {
    param([Parameter(Mandatory = $true)][datetime]$RunDate)
    $daysSinceMonday = (([int]$RunDate.DayOfWeek + 6) % 7)
    $monday = $RunDate.Date.AddDays(-$daysSinceMonday)
    if ($RunDate.DayOfWeek -in @(
        [DayOfWeek]::Monday,
        [DayOfWeek]::Tuesday,
        [DayOfWeek]::Wednesday,
        [DayOfWeek]::Thursday
    )) {
        return [pscustomobject]@{
            StartDate = $monday.AddDays(-7)
            EndDate = $monday.AddDays(-1)
        }
    }
    return [pscustomobject]@{
        StartDate = $monday
        EndDate = $monday.AddDays(4)
    }
}

function Get-OpenOrderTotal {
    param(
        [Parameter(Mandatory = $true)][AllowNull()][object[,]]$Values,
        [Parameter(Mandatory = $true)][int]$LastRow
    )
    $map = New-Map
    for ($row = 9; $row -le $LastRow; $row++) {
        $sku = Get-Key -Value $Values[$row, 16]
        if ([string]::IsNullOrWhiteSpace($sku) -or $sku -match 'TOTAL$') {
            continue
        }
        $total = [double]0
        for ($column = 17; $column -le 22; $column++) {
            if ((Get-SoldTo -Value $Values[8, $column]) -eq $soldToParty) {
                $total += ConvertTo-Number `
                    -Value $Values[$row, $column] `
                    -Context "DS open order row $row, column $column"
            }
        }
        Add-Total -Map $map -Key $sku -Value $total
    }
    return ,$map
}

function Get-WeeklyOpenOrders {
    param(
        [Parameter(Mandatory = $true)][AllowNull()][object[,]]$Values,
        [Parameter(Mandatory = $true)][int]$LastRow,
        [Parameter(Mandatory = $true)][datetime]$RunDate
    )

    $pivotSoldTo = Get-SoldTo -Value $Values[4, 26]
    if ($pivotSoldTo -ne $soldToParty) {
        throw "HK OOR pivot Y:AP is filtered to SoldTo '$pivotSoldTo', not required SoldTo '$soldToParty'."
    }

    $refreshFriday = Get-RefreshFriday -RunDate $RunDate
    $current = New-Map
    $future = [Collections.Hashtable]::new()

    for ($column = 28; $column -le 42; $column++) {
        $header = Get-Key -Value $Values[7, $column]
        if ([string]::IsNullOrWhiteSpace($header) -or $header -eq 'GRAND TOTAL') {
            continue
        }
        $requestedDate = ConvertTo-Date `
            -Value $Values[7, $column] `
            -Context "HK OOR pivot header column $column"
        $bucketFriday = Get-Friday -Date $requestedDate
        $bucketMap = $current
        if ($requestedDate -gt $refreshFriday) {
            $bucketKey = $bucketFriday.ToString('yyyy-MM-dd')
            if (-not $future.ContainsKey($bucketKey)) {
                $future[$bucketKey] = New-Map
            }
            $bucketMap = $future[$bucketKey]
        }

        for ($row = 8; $row -le $LastRow; $row++) {
            $sku = Get-Key -Value $Values[$row, 27]
            if ([string]::IsNullOrWhiteSpace($sku) -or $sku -match 'TOTAL$') {
                continue
            }
            $quantity = ConvertTo-Number `
                -Value $Values[$row, $column] `
                -Context "weekly open order row $row, column $column"
            Add-Total -Map $bucketMap -Key $sku -Value $quantity
        }
    }

    $retained = [Collections.Generic.List[object]]::new()
    foreach ($bucketKey in @($future.Keys | Sort-Object)) {
        $bucketMap = $future[$bucketKey]
        $bucketTotal = [double]0
        foreach ($value in $bucketMap.Values) { $bucketTotal += [double]$value }
        if ([math]::Abs($bucketTotal) -gt 0.000001) {
            $retained.Add([pscustomobject]@{
                Friday = [datetime]::ParseExact(
                    $bucketKey,
                    'yyyy-MM-dd',
                    [Globalization.CultureInfo]::InvariantCulture
                )
                Values = $bucketMap
            })
        }
        if ($retained.Count -eq 7) { break }
    }

    return [pscustomobject]@{
        RefreshFriday = $refreshFriday
        Current = $current
        Future = $retained
    }
}

function Get-MonthlyActuals {
    param(
        [Parameter(Mandatory = $true)][AllowNull()][object[,]]$Values,
        [Parameter(Mandatory = $true)][int]$LastRow
    )

    $months = [Collections.Generic.List[string]]::new()
    $columnMonths = @{}
    $currentMonth = $null
    # Scan only DDR pivot A:AC so month blocks can move within that range.
    for ($column = 1; $column -le 29; $column++) {
        $header = Get-Key -Value $Values[7, $column]
        if ($header -eq 'GRAND TOTAL') { break }
        if ($header -match '^(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)') {
            $currentMonth = $header
            if (-not $months.Contains($currentMonth)) {
                $months.Add($currentMonth)
            }
        }
        if ($null -ne $currentMonth) { $columnMonths[$column] = $currentMonth }
    }

    $maps = @{}
    foreach ($month in $months) { $maps[$month] = New-Map }
    for ($row = 9; $row -le $LastRow; $row++) {
        $material = Get-Key -Value $Values[$row, 2]
        if ([string]::IsNullOrWhiteSpace($material) -or
            $material -match 'TOTAL$') {
            continue
        }
        foreach ($column in $columnMonths.Keys) {
            if ((Get-SoldTo -Value $Values[8, $column]) -ne $soldToParty) {
                continue
            }
            Add-Total `
                -Map $maps[$columnMonths[$column]] `
                -Key $material `
                -Value (ConvertTo-Number `
                    -Value $Values[$row, $column] `
                    -Context "monthly actual row $row, column $column")
        }
    }
    return [pscustomobject]@{ Months = $months; Maps = $maps }
}

function Get-WeeklyActuals {
    param(
        [Parameter(Mandatory = $true)][AllowNull()][object[,]]$Values,
        [Parameter(Mandatory = $true)][int]$LastRow,
        [Parameter(Mandatory = $true)][datetime]$RunDate
    )

    $range = Get-WeeklyActualRange -RunDate $RunDate
    $columns = [Collections.Generic.List[int]]::new()
    $currentSoldTo = ''
    for ($column = 39; $column -le 104; $column++) {
        $soldToHeader = Get-SoldTo -Value $Values[7, $column]
        if (-not [string]::IsNullOrWhiteSpace($soldToHeader)) {
            $currentSoldTo = $soldToHeader
        }
        if ($currentSoldTo -ne $soldToParty) { continue }
        $deliveryDate = ConvertTo-Date `
            -Value $Values[8, $column] `
            -Context "weekly actual header column $column"
        if ($null -ne $deliveryDate -and
            $deliveryDate -ge $range.StartDate -and
            $deliveryDate -le $range.EndDate) {
            $columns.Add($column)
        }
    }

    $map = New-Map
    for ($row = 9; $row -le $LastRow; $row++) {
        $material = Get-Key -Value $Values[$row, 36]
        if ([string]::IsNullOrWhiteSpace($material) -or
            $material -match 'TOTAL$') {
            continue
        }
        $total = [double]0
        foreach ($column in $columns) {
            $total += ConvertTo-Number `
                -Value $Values[$row, $column] `
                -Context "weekly actual row $row, column $column"
        }
        Add-Total -Map $map -Key $material -Value $total
    }
    return [pscustomobject]@{
        Values = $map
        StartDate = $range.StartDate
        EndDate = $range.EndDate
    }
}

function Get-RetailInventory {
    param(
        [Parameter(Mandatory = $true)][AllowNull()][object[,]]$Values,
        [Parameter(Mandatory = $true)][int]$LastRow
    )
    $maps = @(
        (New-Map)
        (New-Map)
        (New-Map)
    )
    for ($row = 2; $row -le $LastRow; $row++) {
        $material = Get-Key -Value $Values[$row, 1]
        if ([string]::IsNullOrWhiteSpace($material)) { continue }
        $maps[0][$material] = ConvertTo-Number `
            -Value $Values[$row, 8] -Context "Retail INV H$row"
        $maps[1][$material] = ConvertTo-Number `
            -Value $Values[$row, 9] -Context "Retail INV I$row"
        $maps[2][$material] = ConvertTo-Number `
            -Value $Values[$row, 10] -Context "Retail INV J$row"
    }
    return ,$maps
}

function Get-LastDataRow {
    param(
        [Parameter(Mandatory = $true)][object]$Worksheet,
        [Parameter(Mandatory = $true)][int]$Column
    )

    $bottomCell = $null
    $lastCell = $null
    try {
        $bottomCell = $Worksheet.Cells.Item(1048576, $Column)
        $lastCell = $bottomCell.End(-4162)
        return [int]$lastCell.Row
    }
    finally {
        Release-ComObject -ComObject $lastCell
        Release-ComObject -ComObject $bottomCell
    }
}

function Get-RowVisibilityRuns {
    param(
        [Parameter(Mandatory = $true)][object]$Worksheet,
        [Parameter(Mandatory = $true)][int]$StartRow,
        [Parameter(Mandatory = $true)][int]$EndRow
    )

    $runs = [Collections.Generic.List[object]]::new()
    $runStart = $StartRow
    $runHidden = $null
    for ($row = $StartRow; $row -le $EndRow; $row++) {
        $worksheetRow = $Worksheet.Rows.Item($row)
        try {
            $hidden = [bool]$worksheetRow.Hidden
        }
        finally {
            Release-ComObject -ComObject $worksheetRow
        }

        if ($null -eq $runHidden) {
            $runHidden = $hidden
            continue
        }
        if ($hidden -eq $runHidden) {
            continue
        }

        $runs.Add([pscustomobject]@{
            StartRow = $runStart
            EndRow = $row - 1
        })
        $runStart = $row
        $runHidden = $hidden
    }
    $runs.Add([pscustomobject]@{
        StartRow = $runStart
        EndRow = $EndRow
    })
    return $runs.ToArray()
}

function Get-ColumnVisibilityRuns {
    param(
        [Parameter(Mandatory = $true)][object]$Worksheet,
        [Parameter(Mandatory = $true)][int]$StartColumn,
        [Parameter(Mandatory = $true)][int]$EndColumn
    )

    $runs = [Collections.Generic.List[object]]::new()
    $runStart = $StartColumn
    $runHidden = $null
    for ($column = $StartColumn; $column -le $EndColumn; $column++) {
        $worksheetColumn = $Worksheet.Columns.Item($column)
        try {
            $hidden = [bool]$worksheetColumn.Hidden
        }
        finally {
            Release-ComObject -ComObject $worksheetColumn
        }

        if ($null -eq $runHidden) {
            $runHidden = $hidden
            continue
        }
        if ($hidden -eq $runHidden) {
            continue
        }

        $runs.Add([pscustomobject]@{
            StartColumn = $runStart
            EndColumn = $column - 1
        })
        $runStart = $column
        $runHidden = $hidden
    }
    $runs.Add([pscustomobject]@{
        StartColumn = $runStart
        EndColumn = $EndColumn
    })
    return $runs.ToArray()
}

function Set-RangeValues {
    param(
        [Parameter(Mandatory = $true)][object]$Worksheet,
        [Parameter(Mandatory = $true)][object]$Range,
        [Parameter(Mandatory = $true)][AllowNull()][object[,]]$Values,
        [Parameter(Mandatory = $true)][object[]]$RowRuns,
        [Parameter(Mandatory = $true)][string]$Context
    )

    $startRow = [int]$Range.Row
    $startColumn = [int]$Range.Column
    $columnCount = $Values.GetLength(1)
    $columnRuns = Get-ColumnVisibilityRuns `
        -Worksheet $Worksheet `
        -StartColumn $startColumn `
        -EndColumn ($startColumn + $columnCount - 1)
    foreach ($run in $RowRuns) {
        $runLength = $run.EndRow - $run.StartRow + 1
        $sourceStart = $run.StartRow - $startRow
        foreach ($columnRun in $columnRuns) {
            $runColumnCount =
                $columnRun.EndColumn - $columnRun.StartColumn + 1
            $sourceColumnStart = $columnRun.StartColumn - $startColumn
            $runValues = New-Object `
                'object[,]' $runLength, $runColumnCount
            for ($rowOffset = 0; $rowOffset -lt $runLength; $rowOffset++) {
                for (
                    $columnOffset = 0;
                    $columnOffset -lt $runColumnCount;
                    $columnOffset++
                ) {
                    $runValues[$rowOffset, $columnOffset] = $Values[
                        ($sourceStart + $rowOffset),
                        ($sourceColumnStart + $columnOffset)
                    ]
                }
            }

            Invoke-ExcelComRetry `
                -Context (
                    "$Context rows $($run.StartRow)-$($run.EndRow), " +
                    "columns $($columnRun.StartColumn)-$($columnRun.EndColumn)"
                ) `
                -Operation {
                    $firstCell = $null
                    $lastCell = $null
                    $runRange = $null
                    try {
                        $firstCell = $Worksheet.Cells.Item(
                            $run.StartRow,
                            $columnRun.StartColumn
                        )
                        $lastCell = $Worksheet.Cells.Item(
                            $run.EndRow,
                            $columnRun.EndColumn
                        )
                        $runRange = $Worksheet.Range($firstCell, $lastCell)
                        $runRange.Value2 = $runValues
                    }
                    finally {
                        Release-ComObject -ComObject $runRange
                        Release-ComObject -ComObject $lastCell
                        Release-ComObject -ComObject $firstCell
                    }
                }
        }
    }
}

function Assert-ArrayValues {
    param(
        [Parameter(Mandatory = $true)][AllowNull()][array]$Actual,
        [Parameter(Mandatory = $true)][AllowNull()][array]$Expected,
        [Parameter(Mandatory = $true)][string]$RangeName
    )

    $rowCount = $Expected.GetLength(0)
    $columnCount = $Expected.GetLength(1)
    if ($Actual.Rank -eq 1 -and
        $Actual.Length -ne ($rowCount * $columnCount)) {
        throw "$RangeName validation returned an unexpected value count."
    }

    for ($rowIndex = 1; $rowIndex -le $rowCount; $rowIndex++) {
        for (
            $columnIndex = 1;
            $columnIndex -le $columnCount;
            $columnIndex++
        ) {
            $actualValue = if ($Actual.Rank -eq 2) {
                $Actual[$rowIndex, $columnIndex]
            }
            else {
                $Actual[
                    (($rowIndex - 1) * $columnCount) + ($columnIndex - 1)
                ]
            }
            $numericActual = ConvertTo-Number `
                -Value $actualValue `
                -Context "$RangeName validation"
            $expectedValue = [double]$Expected[
                ($rowIndex - 1),
                ($columnIndex - 1)
            ]
            if ([math]::Abs($numericActual - $expectedValue) -gt 0.000001) {
                throw "$RangeName validation failed at relative row $rowIndex, column $columnIndex. Expected $expectedValue but found $numericActual."
            }
        }
    }
}

foreach ($path in @($SourceWorkbookPath, $DestinationWorkbookPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Workbook not found: $path"
    }
}

$sourcePath = (Resolve-Path -LiteralPath $SourceWorkbookPath).Path
$destinationPath = (Resolve-Path -LiteralPath $DestinationWorkbookPath).Path
$excel = $null
$sourceWorkbook = $null
$destinationWorkbook = $null
$excelProcess = $null
$refreshSucceeded = $false
$comFilterRegistered = $false

try {
    Register-RetailExcelComFilter
    $comFilterRegistered = $true

    if (-not ('RetailExcelProcessResolver' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class RetailExcelProcessResolver
{
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);
    public static int GetProcessId(IntPtr windowHandle)
    {
        uint processId;
        GetWindowThreadProcessId(windowHandle, out processId);
        return unchecked((int)processId);
    }
}
'@
    }

    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = [RetailExcelProcessResolver]::GetProcessId(
        [IntPtr]$excel.Hwnd
    )
    $excelProcess = [Diagnostics.Process]::GetProcessById($excelProcessId)
    [void]$excelProcess.Handle
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $excel.ScreenUpdating = $false

    $sourceWorkbook = $excel.Workbooks.Open($sourcePath, 0, $true)
    $destinationWorkbook = $excel.Workbooks.Open($destinationPath, 0, $false)
    if ([bool]$destinationWorkbook.ReadOnly) {
        throw "Destination workbook opened as read-only: $destinationPath"
    }

    $oorSheet = $sourceWorkbook.Worksheets.Item('HK OOR Pivot')
    $ddrSheet = $sourceWorkbook.Worksheets.Item('DDR pivot')
    $inventorySheet = $sourceWorkbook.Worksheets.Item('Retail INV')
    $targetSheet = $destinationWorkbook.Worksheets.Item('Supportability')

    $oorLastRow = [math]::Max(
        (Get-LastDataRow -Worksheet $oorSheet -Column 16),
        (Get-LastDataRow -Worksheet $oorSheet -Column 27)
    )
    $oorRange = $oorSheet.Range("A1:AP$oorLastRow")
    $oorValues = $oorRange.Value2

    $ddrLastRow = [math]::Max(
        (Get-LastDataRow -Worksheet $ddrSheet -Column 2),
        (Get-LastDataRow -Worksheet $ddrSheet -Column 36)
    )
    $ddrRange = $ddrSheet.Range("A1:CZ$ddrLastRow")
    $ddrValues = $ddrRange.Value2

    $inventoryLastRow = Get-LastDataRow -Worksheet $inventorySheet -Column 1
    $inventoryRange = $inventorySheet.Range("A1:J$inventoryLastRow")
    $inventoryValues = $inventoryRange.Value2

    $targetLastRow = Get-LastDataRow -Worksheet $targetSheet -Column 16
    if ($targetLastRow -lt $targetStartRow) {
        throw 'The destination Supportability sheet has no product rows.'
    }
    $productRange = $targetSheet.Range("P${targetStartRow}:P${targetLastRow}")
    $productValues = $productRange.Value2

    $dsOpenOrder = Get-OpenOrderTotal `
        -Values $oorValues -LastRow $oorLastRow
    $weeklyOpenOrder = Get-WeeklyOpenOrders `
        -Values $oorValues -LastRow $oorLastRow -RunDate $AsOfDate
    $monthly = Get-MonthlyActuals `
        -Values $ddrValues -LastRow $ddrLastRow
    $weeklyActual = Get-WeeklyActuals `
        -Values $ddrValues -LastRow $ddrLastRow -RunDate $AsOfDate
    $inventory = Get-RetailInventory `
        -Values $inventoryValues -LastRow $inventoryLastRow

    $targetRowCount = $targetLastRow - $targetStartRow + 1
    $monthlyOutput = New-Object 'object[,]' $targetRowCount, 3
    $openOrderOutput = New-Object 'object[,]' $targetRowCount, 9
    $weeklyActualOutput = New-Object 'object[,]' $targetRowCount, 1
    $inventoryHOutput = New-Object 'object[,]' $targetRowCount, 1
    $inventoryIOutput = New-Object 'object[,]' $targetRowCount, 1
    $inventoryJOutput = New-Object 'object[,]' $targetRowCount, 1

    for ($outputRow = 0; $outputRow -lt $targetRowCount; $outputRow++) {
        $productId = Get-Key -Value $productValues[($outputRow + 1), 1]
        for ($index = 0; $index -lt 3; $index++) {
            $monthValue = [double]0
            if ($index -lt $monthly.Months.Count) {
                $month = $monthly.Months[$index]
                $monthValue = Get-MapValue `
                    -Map $monthly.Maps[$month] -Key $productId
            }
            $monthlyOutput[$outputRow, $index] = $monthValue
        }

        $openOrderOutput[$outputRow, 0] =
            Get-MapValue -Map $dsOpenOrder -Key $productId
        $openOrderOutput[$outputRow, 1] = Get-MapValue `
            -Map $weeklyOpenOrder.Current -Key $productId

        for ($index = 0; $index -lt 7; $index++) {
            $futureValue = [double]0
            if ($index -lt $weeklyOpenOrder.Future.Count) {
                $futureValue = Get-MapValue `
                    -Map $weeklyOpenOrder.Future[$index].Values `
                    -Key $productId
            }
            $openOrderOutput[$outputRow, (2 + $index)] = $futureValue
        }

        $weeklyActualOutput[$outputRow, 0] =
            Get-MapValue -Map $weeklyActual.Values -Key $productId
        $inventoryHOutput[$outputRow, 0] =
            Get-MapValue -Map $inventory[0] -Key $productId
        $inventoryIOutput[$outputRow, 0] =
            Get-MapValue -Map $inventory[1] -Key $productId
        $inventoryJOutput[$outputRow, 0] =
            Get-MapValue -Map $inventory[2] -Key $productId
    }

    $monthlyTarget = $targetSheet.Range(
        "U${targetStartRow}:W${targetLastRow}"
    )
    $openOrderTarget = $targetSheet.Range(
        "AS${targetStartRow}:BA${targetLastRow}"
    )
    $weeklyActualTarget = $targetSheet.Range(
        "BT${targetStartRow}:BT${targetLastRow}"
    )
    $inventoryHTarget = $targetSheet.Range(
        "CJ${targetStartRow}:CJ${targetLastRow}"
    )
    $inventoryITarget = $targetSheet.Range(
        "CL${targetStartRow}:CL${targetLastRow}"
    )
    $inventoryJTarget = $targetSheet.Range(
        "CN${targetStartRow}:CN${targetLastRow}"
    )
    $rowRuns = Get-RowVisibilityRuns `
        -Worksheet $targetSheet `
        -StartRow $targetStartRow `
        -EndRow $targetLastRow

    Set-RangeValues -Worksheet $targetSheet -Range $monthlyTarget `
        -Values $monthlyOutput -RowRuns $rowRuns `
        -Context 'write Retail Supportability U:W'
    Set-RangeValues -Worksheet $targetSheet -Range $openOrderTarget `
        -Values $openOrderOutput -RowRuns $rowRuns `
        -Context 'write Retail Supportability AS:BA'
    Set-RangeValues -Worksheet $targetSheet -Range $weeklyActualTarget `
        -Values $weeklyActualOutput -RowRuns $rowRuns `
        -Context 'write Retail Supportability BT'
    Set-RangeValues -Worksheet $targetSheet -Range $inventoryHTarget `
        -Values $inventoryHOutput -RowRuns $rowRuns `
        -Context 'write Retail Supportability CJ'
    Set-RangeValues -Worksheet $targetSheet -Range $inventoryITarget `
        -Values $inventoryIOutput -RowRuns $rowRuns `
        -Context 'write Retail Supportability CL'
    Set-RangeValues -Worksheet $targetSheet -Range $inventoryJTarget `
        -Values $inventoryJOutput -RowRuns $rowRuns `
        -Context 'write Retail Supportability CN'

    for ($index = 0; $index -lt 7; $index++) {
        Invoke-ExcelComRetry `
            -Context "write future header column $(47 + $index)" `
            -Operation {
                $header = $null
                try {
                    $header = $targetSheet.Cells.Item(3, 47 + $index)
                    if ($index -lt $weeklyOpenOrder.Future.Count) {
                        $header.Value2 =
                            $weeklyOpenOrder.Future[$index].Friday.ToOADate()
                        $header.NumberFormat = 'mm/dd'
                    }
                    else {
                        $header.ClearContents()
                    }
                }
                finally {
                    Release-ComObject -ComObject $header
                }
            }
    }

    Invoke-ExcelComRetry -Context 'save Retail destination workbook' -Operation {
        $destinationWorkbook.Save()
    }

    Assert-ArrayValues -Actual $monthlyTarget.Value2 `
        -Expected $monthlyOutput -RangeName 'Supportability U:W'
    Assert-ArrayValues -Actual $openOrderTarget.Value2 `
        -Expected $openOrderOutput -RangeName 'Supportability AS:BA'
    Assert-ArrayValues -Actual $weeklyActualTarget.Value2 `
        -Expected $weeklyActualOutput -RangeName 'Supportability BT'
    Assert-ArrayValues -Actual $inventoryHTarget.Value2 `
        -Expected $inventoryHOutput -RangeName 'Supportability CJ'
    Assert-ArrayValues -Actual $inventoryITarget.Value2 `
        -Expected $inventoryIOutput -RangeName 'Supportability CL'
    Assert-ArrayValues -Actual $inventoryJTarget.Value2 `
        -Expected $inventoryJOutput -RangeName 'Supportability CN'

    for ($index = 0; $index -lt 7; $index++) {
        Invoke-ExcelComRetry `
            -Context "validate future header column $(47 + $index)" `
            -Operation {
                $header = $null
                try {
                    $header = $targetSheet.Cells.Item(3, 47 + $index)
                    if ($index -lt $weeklyOpenOrder.Future.Count) {
                        $expectedDate = $weeklyOpenOrder.Future[$index].Friday
                        $actualDate = ConvertTo-Date `
                            -Value $header.Value2 `
                            -Context "future weekly header column $(47 + $index)"
                        if ($actualDate -ne $expectedDate.Date) {
                            throw "Future weekly header column $(47 + $index) has an incorrect date."
                        }
                        if ([string]$header.NumberFormat -ne 'mm/dd') {
                            throw "Future weekly header column $(47 + $index) does not use mm/dd format."
                        }
                    }
                    elseif (-not [string]::IsNullOrEmpty(
                        [string]$header.Value2
                    )) {
                        throw "Unused future weekly header column $(47 + $index) is not blank."
                    }
                }
                finally {
                    Release-ComObject -ComObject $header
                }
            }
    }

    $refreshSucceeded = $true
}
finally {
    if ($null -ne $destinationWorkbook) {
        try { $destinationWorkbook.Close($false) }
        catch { Write-Warning "Failed to close destination: $($_.Exception.Message)" }
    }
    if ($null -ne $sourceWorkbook) {
        try { $sourceWorkbook.Close($false) }
        catch { Write-Warning "Failed to close source: $($_.Exception.Message)" }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() }
        catch { Write-Warning "Failed to quit Excel: $($_.Exception.Message)" }
    }
    foreach ($object in @(
        $inventoryJTarget, $inventoryITarget, $inventoryHTarget,
        $weeklyActualTarget, $openOrderTarget, $monthlyTarget,
        $productRange, $inventoryRange, $ddrRange, $oorRange, $targetSheet,
        $inventorySheet, $ddrSheet, $oorSheet, $destinationWorkbook,
        $sourceWorkbook, $excel
    )) { Release-ComObject -ComObject $object }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($comFilterRegistered) {
        [RetailOleMessageFilter]::Revoke()
    }
    if ($null -ne $excelProcess) {
        try {
            if (-not $excelProcess.HasExited -and
                -not $excelProcess.WaitForExit(5000)) {
                $excelProcess.Kill()
                $excelProcess.WaitForExit()
                Write-Warning "Terminated the script-owned Excel process $excelProcessId."
            }
        }
        finally { $excelProcess.Dispose() }
    }
}

if ($refreshSucceeded) {
    Write-Host (
        'Retail supportability report refresh succeeded. ' +
        "Weekly actual range: " +
        "$($weeklyActual.StartDate.ToString('yyyy-MM-dd')) through " +
        "$($weeklyActual.EndDate.ToString('yyyy-MM-dd'))."
    )
}
