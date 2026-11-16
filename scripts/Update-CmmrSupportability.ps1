[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceWorkbookPath,

    [ValidateNotNullOrEmpty()]
    [string]$DestinationWorkbookPath = $SourceWorkbookPath,

    [datetime]$AsOfDate = (Get-Date),

    [switch]$MonthlyOnly,

    [switch]$OpenOrdersOnly,

    [switch]$InventoryOnly,

    [switch]$WeeklyOnly,

    [string]$DestinationSheetName
)

$ErrorActionPreference = 'Stop'

if ($OpenOrdersOnly -and $MonthlyOnly) {
    throw 'OpenOrdersOnly cannot be combined with MonthlyOnly.'
}
if ($InventoryOnly -and ($MonthlyOnly -or $OpenOrdersOnly)) {
    throw 'InventoryOnly cannot be combined with another Only mode.'
}
if ($WeeklyOnly -and
    ($MonthlyOnly -or $OpenOrdersOnly -or $InventoryOnly)) {
    throw 'WeeklyOnly cannot be combined with another Only mode.'
}

$targetStartRow = 4
$targetProductIdColumn = 12
$soldToParties = @('5395800', '5395801', '5395799')

function Release-ComObject {
    param(
        [AllowNull()]
        [object]$ComObject
    )

    if ($null -ne $ComObject -and
        [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(
            $ComObject
        )
    }
}

function Register-CmmrExcelComFilter {
    if (-not ('CmmrOleMessageFilter' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface ICmmrOleMessageFilter
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

public sealed class CmmrOleMessageFilter : ICmmrOleMessageFilter
{
    private static CmmrOleMessageFilter current;
    private static ICmmrOleMessageFilter previous;

    [DllImport("Ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        ICmmrOleMessageFilter newFilter,
        out ICmmrOleMessageFilter oldFilter);

    public static void Register()
    {
        if (current != null)
        {
            return;
        }
        current = new CmmrOleMessageFilter();
        ICmmrOleMessageFilter oldFilter;
        CoRegisterMessageFilter(current, out oldFilter);
        previous = oldFilter;
    }

    public static void Revoke()
    {
        ICmmrOleMessageFilter ignored;
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

    [CmmrOleMessageFilter]::Register()
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
        catch {
            $comException = $_.Exception
            while ($comException -isnot [Runtime.InteropServices.COMException] -and
                $null -ne $comException.InnerException) {
                $comException = $comException.InnerException
            }
            if ($comException -isnot [Runtime.InteropServices.COMException]) {
                throw
            }
            $hresult = $comException.HResult -band 0xFFFFFFFFL
            if ($hresult -notin @(0x80010001L, 0x8001010AL, 0x800AC472L)) {
                throw
            }
            $lastComException = $comException
            $remainingMilliseconds =
                ($deadline - [datetime]::UtcNow).TotalMilliseconds
            if ($remainingMilliseconds -le 250) {
                $hresultText = '0x{0:X8}' -f $hresult
                throw "Excel COM operation '$Context' failed after $TimeoutSeconds seconds ($hresultText): $($comException.Message)"
            }
            Start-Sleep -Milliseconds 250
        }
    }
}

function Get-NormalizedKey {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim().ToUpperInvariant()
}

function Get-NormalizedSoldTo {
    param(
        [AllowNull()]
        [object]$Value
    )

    $text = Get-NormalizedKey -Value $Value
    if ($text -match '^\d+$') {
        return $text.TrimStart('0')
    }

    return $text
}

function ConvertTo-Number {
    param(
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return [double]0
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

    $normalized = ([string]$Value).Trim()
    $normalized = $normalized.Replace([char]0x00A0, ' ').Replace(' ', '')
    if ($normalized.StartsWith('(') -and $normalized.EndsWith(')')) {
        $normalized = '-' + $normalized.Substring(1, $normalized.Length - 2)
    }

    $number = [double]0
    $styles = [System.Globalization.NumberStyles]::Float -bor
        [System.Globalization.NumberStyles]::AllowThousands
    if (-not [double]::TryParse(
        $normalized,
        $styles,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )) {
        throw "Expected a numeric value for $Context, but found '$Value'."
    }

    return $number
}

function Add-ToMap {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]$Map,

        [Parameter(Mandatory = $true)]
        [string]$Key,

        [Parameter(Mandatory = $true)]
        [double]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Key)) {
        return
    }

    if (-not $Map.ContainsKey($Key)) {
        $Map[$Key] = [double]0
    }
    $Map[$Key] = [double]$Map[$Key] + $Value
}

function Get-PivotBlockTotals {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [int]$LastRow,

        [Parameter(Mandatory = $true)]
        [int]$SkuColumn,

        [Parameter(Mandatory = $true)]
        [int]$SoldToHeaderRow,

        [Parameter(Mandatory = $true)]
        [int]$FirstValueColumn,

        [Parameter(Mandatory = $true)]
        [int]$LastValueColumn,

        [Parameter(Mandatory = $true)]
        [string[]]$IncludedSoldTo
    )

    $included = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($soldTo in $IncludedSoldTo) {
        $included[(Get-NormalizedSoldTo -Value $soldTo)] = $true
    }

    $valueColumns = [System.Collections.Generic.List[int]]::new()
    for (
        $columnIndex = $FirstValueColumn;
        $columnIndex -le $LastValueColumn;
        $columnIndex++
    ) {
        $soldTo = Get-NormalizedSoldTo `
            -Value $Values[$SoldToHeaderRow, $columnIndex]
        if ($included.ContainsKey($soldTo)) {
            $valueColumns.Add($columnIndex)
        }
    }

    $totals = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    for ($rowIndex = $SoldToHeaderRow + 1; $rowIndex -le $LastRow; $rowIndex++) {
        $sku = Get-NormalizedKey -Value $Values[$rowIndex, $SkuColumn]
        if ([string]::IsNullOrWhiteSpace($sku) -or
            $sku -match 'TOTAL$') {
            continue
        }

        $total = [double]0
        foreach ($columnIndex in $valueColumns) {
            $total += ConvertTo-Number `
                -Value $Values[$rowIndex, $columnIndex] `
                -Context "pivot row $rowIndex, column $columnIndex"
        }
        Add-ToMap -Map $totals -Key $sku -Value $total
    }

    return ,$totals
}

function Get-DisplayedMonth {
    param(
        [AllowNull()]
        [object]$Value
    )

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    foreach ($format in @('MMM', 'MMMM')) {
        $month = [datetime]::MinValue
        if ([datetime]::TryParseExact(
            $text,
            $format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$month
        )) {
            return $text
        }
    }

    return $null
}

function Get-MonthlyActuals {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [int]$LastRow
    )

    $monthOrder = [System.Collections.Generic.List[string]]::new()
    $columnMonths = @{}
    $currentMonth = $null
    # Scan only DDR pivot A:AC so month blocks can move within that range.
    for ($columnIndex = 1; $columnIndex -le 29; $columnIndex++) {
        $header = Get-NormalizedKey -Value $Values[7, $columnIndex]
        if ($header -eq 'GRAND TOTAL') {
            break
        }

        $displayedMonth = Get-DisplayedMonth -Value $Values[7, $columnIndex]
        if ($null -ne $displayedMonth) {
            $currentMonth = $displayedMonth
            if (-not $monthOrder.Contains($currentMonth)) {
                $monthOrder.Add($currentMonth)
            }
        }

        if ($null -ne $currentMonth) {
            $columnMonths[$columnIndex] = $currentMonth
        }
    }

    $monthMaps = @{}
    foreach ($month in $monthOrder) {
        $monthMaps[$month] = [System.Collections.Hashtable]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
    }

    for ($rowIndex = 9; $rowIndex -le $LastRow; $rowIndex++) {
        $material = Get-NormalizedKey -Value $Values[$rowIndex, 2]
        if ([string]::IsNullOrWhiteSpace($material) -or
            $material -match 'TOTAL$') {
            continue
        }

        foreach ($columnIndex in $columnMonths.Keys) {
            $soldTo = Get-NormalizedSoldTo -Value $Values[8, $columnIndex]
            if ($soldTo -notin $soldToParties) {
                continue
            }

            $month = $columnMonths[$columnIndex]
            $quantity = ConvertTo-Number `
                -Value $Values[$rowIndex, $columnIndex] `
                -Context "monthly actual row $rowIndex, column $columnIndex"
            Add-ToMap -Map $monthMaps[$month] -Key $material -Value $quantity
        }
    }

    return [pscustomobject]@{
        MonthOrder = $monthOrder
        MonthMaps = $monthMaps
    }
}

function ConvertTo-ExcelDate {
    param(
        [AllowNull()]
        [object]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Context
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    if ($Value -is [datetime]) {
        return ([datetime]$Value).Date
    }

    if ($Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        return [datetime]::FromOADate([double]$Value).Date
    }

    $date = [datetime]::MinValue
    if ([datetime]::TryParse(
        ([string]$Value).Trim(),
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None,
        [ref]$date
    )) {
        return $date.Date
    }

    throw "Expected a date for $Context, but found '$Value'."
}

function Get-WeeklyDateRange {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$RunDate
    )

    $date = $RunDate.Date
    $daysSinceMonday = (([int]$date.DayOfWeek + 6) % 7)
    $currentMonday = $date.AddDays(-$daysSinceMonday)

    if ($date.DayOfWeek -in @(
        [DayOfWeek]::Monday,
        [DayOfWeek]::Tuesday,
        [DayOfWeek]::Wednesday,
        [DayOfWeek]::Thursday
    )) {
        return [pscustomobject]@{
            StartDate = $currentMonday.AddDays(-7)
            EndDate = $currentMonday.AddDays(-1)
        }
    }

    return [pscustomobject]@{
        StartDate = $currentMonday
        EndDate = $currentMonday.AddDays(3)
    }
}

function Get-WeeklyActuals {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [int]$LastRow,

        [Parameter(Mandatory = $true)]
        [datetime]$RunDate
    )

    $dateRange = Get-WeeklyDateRange -RunDate $RunDate
    $columns = [System.Collections.Generic.List[int]]::new()
    $currentSoldTo = ''

    for ($columnIndex = 39; $columnIndex -le 104; $columnIndex++) {
        $soldToHeader = Get-NormalizedSoldTo -Value $Values[7, $columnIndex]
        if (-not [string]::IsNullOrWhiteSpace($soldToHeader)) {
            $currentSoldTo = $soldToHeader
        }
        if ($currentSoldTo -notin $soldToParties) {
            continue
        }

        $deliveryDate = ConvertTo-ExcelDate `
            -Value $Values[8, $columnIndex] `
            -Context "weekly header column $columnIndex"
        if ($null -ne $deliveryDate -and
            $deliveryDate -ge $dateRange.StartDate -and
            $deliveryDate -le $dateRange.EndDate) {
            $columns.Add($columnIndex)
        }
    }

    $totals = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    for ($rowIndex = 9; $rowIndex -le $LastRow; $rowIndex++) {
        $material = Get-NormalizedKey -Value $Values[$rowIndex, 36]
        if ([string]::IsNullOrWhiteSpace($material) -or
            $material -match 'TOTAL$') {
            continue
        }

        $total = [double]0
        foreach ($columnIndex in $columns) {
            $total += ConvertTo-Number `
                -Value $Values[$rowIndex, $columnIndex] `
                -Context "weekly actual row $rowIndex, column $columnIndex"
        }
        Add-ToMap -Map $totals -Key $material -Value $total
    }

    return [pscustomobject]@{
        Totals = $totals
        StartDate = $dateRange.StartDate
        EndDate = $dateRange.EndDate
    }
}

function Get-InventoryTotals {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [int]$LastRow
    )

    $totals = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    for ($rowIndex = 2; $rowIndex -le $LastRow; $rowIndex++) {
        $material = Get-NormalizedKey -Value $Values[$rowIndex, 1]
        if ([string]::IsNullOrWhiteSpace($material)) {
            continue
        }

        $totals[$material] = ConvertTo-Number `
            -Value $Values[$rowIndex, 10] `
            -Context "CMMR INV row $rowIndex, column J"
    }

    return ,$totals
}

function Get-MapValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]$Map,

        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Key
    )

    if ($Map.ContainsKey($Key)) {
        return [double]$Map[$Key]
    }

    return [double]0
}

function Assert-ArrayValues {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Actual,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$RangeName
    )

    for ($rowIndex = 1; $rowIndex -le $Expected.GetLength(0); $rowIndex++) {
        for (
            $columnIndex = 1;
            $columnIndex -le $Expected.GetLength(1);
            $columnIndex++
        ) {
            $actualValue = ConvertTo-Number `
                -Value $Actual[$rowIndex, $columnIndex] `
                -Context "$RangeName validation"
            $expectedValue = [double]$Expected[
                ($rowIndex - 1),
                ($columnIndex - 1)
            ]
            if ([math]::Abs($actualValue - $expectedValue) -gt 0.000001) {
                throw "$RangeName validation failed at relative row $rowIndex, column $columnIndex. Expected $expectedValue but found $actualValue."
            }
        }
    }
}

function Get-LastDataRow {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$Column
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

    $runs = [System.Collections.Generic.List[object]]::new()
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

    $runs = [System.Collections.Generic.List[object]]::new()
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
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [object]$Range,

        [Parameter(Mandatory = $true)]
        [AllowNull()][object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [object[]]$RowRuns
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

foreach ($path in @($SourceWorkbookPath, $DestinationWorkbookPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Workbook not found: $path"
    }
}

$sourcePath = (Resolve-Path -LiteralPath $SourceWorkbookPath).Path
$destinationPath = (Resolve-Path -LiteralPath $DestinationWorkbookPath).Path
$sameWorkbook = $sourcePath -eq $destinationPath
$monthlyColumns = if ($sameWorkbook) { @('N', 'P') } else { @('O', 'Q') }
$openOrderColumns = if ($sameWorkbook) { @('R', 'U') } else { @('S', 'V') }
$stockWeeklyColumns = if ($sameWorkbook) {
    if ($InventoryOnly) { @('AK', 'AK') }
    elseif ($WeeklyOnly) { @('AL', 'AL') }
    else { @('AK', 'AL') }
} else {
    if ($InventoryOnly) { @('AM', 'AM') }
    elseif ($WeeklyOnly) { @('AN', 'AN') }
    else { @('AM', 'AN') }
}
if (-not $PSBoundParameters.ContainsKey('DestinationSheetName')) {
    $DestinationSheetName = if ($sameWorkbook) {
        'CMMR Supportability'
    }
    else {
        'Supportability'
    }
}
if ([string]::IsNullOrWhiteSpace($DestinationSheetName)) {
    throw 'DestinationSheetName cannot be blank.'
}
$excel = $null
$sourceWorkbook = $null
$destinationWorkbook = $null
$excelProcess = $null
$refreshSucceeded = $false
$comFilterRegistered = $false

try {
    Register-CmmrExcelComFilter
    $comFilterRegistered = $true

    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = 0
    $getProcessId = @'
using System;
using System.Runtime.InteropServices;
public static class ExcelProcessResolver
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
    if (-not ('ExcelProcessResolver' -as [type])) {
        Add-Type -TypeDefinition $getProcessId
    }
    $excelProcessId = [ExcelProcessResolver]::GetProcessId(
        [IntPtr]$excel.Hwnd
    )
    $excelProcess = [System.Diagnostics.Process]::GetProcessById(
        $excelProcessId
    )
    [void]$excelProcess.Handle

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $excel.ScreenUpdating = $false

    $sourceWorkbook = $excel.Workbooks.Open($sourcePath, 0, (-not $sameWorkbook))
    if ($sameWorkbook) {
        $destinationWorkbook = $sourceWorkbook
    }
    else {
        $destinationWorkbook = $excel.Workbooks.Open($destinationPath, 0, $false)
    }
    if ([bool]$destinationWorkbook.ReadOnly) {
        throw "Destination workbook opened as read-only: $destinationPath"
    }

    if (-not $MonthlyOnly -and -not $OpenOrdersOnly -and -not $WeeklyOnly) {
        $inventorySheet = $sourceWorkbook.Worksheets.Item('CMMR INV')
    }
    $targetSheet = $destinationWorkbook.Worksheets.Item($DestinationSheetName)

    if (-not $InventoryOnly) {
        if (-not $OpenOrdersOnly) {
            $ddrPivotSheet = $sourceWorkbook.Worksheets.Item('DDR pivot')
            $ddrLastRow = Get-LastDataRow -Worksheet $ddrPivotSheet -Column 2
            $ddrLastColumn = 'AC'
            if (-not $MonthlyOnly) {
                $ddrLastRow = [math]::Max(
                    $ddrLastRow,
                    (Get-LastDataRow -Worksheet $ddrPivotSheet -Column 36)
                )
                $ddrLastColumn = 'CZ'
            }
            $ddrRange = $ddrPivotSheet.Range("A1:${ddrLastColumn}$ddrLastRow")
            $ddrValues = $ddrRange.Value2
        }
        if (-not $MonthlyOnly -and -not $WeeklyOnly) {
            $oorSheet = $sourceWorkbook.Worksheets.Item('HK OOR Pivot')
            $oorLastRow = [math]::Max(
                (Get-LastDataRow -Worksheet $oorSheet -Column 3),
                (Get-LastDataRow -Worksheet $oorSheet -Column 16)
            )
            $oorRange = $oorSheet.Range("A1:V$oorLastRow")
            $oorValues = $oorRange.Value2
        }
    }

    if (-not $MonthlyOnly -and -not $OpenOrdersOnly -and -not $WeeklyOnly) {
        $inventoryLastRow = Get-LastDataRow -Worksheet $inventorySheet -Column 1
        $inventoryRange = $inventorySheet.Range("A1:J$inventoryLastRow")
        $inventoryValues = $inventoryRange.Value2
    }

    $targetLastRow = Get-LastDataRow -Worksheet $targetSheet -Column 12
    if ($targetLastRow -lt $targetStartRow) {
        throw "The destination '$DestinationSheetName' sheet has no product rows."
    }
    $targetProductRange = $targetSheet.Range(
        "L${targetStartRow}:L${targetLastRow}"
    )
    $targetProductValues = $targetProductRange.Value2

    if (-not $MonthlyOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        $columnS = Get-PivotBlockTotals `
            -Values $oorValues `
            -LastRow $oorLastRow `
            -SkuColumn 3 `
            -SoldToHeaderRow 8 `
            -FirstValueColumn 4 `
            -LastValueColumn 10 `
            -IncludedSoldTo @('5395800', '5395801')
        $columnT = Get-PivotBlockTotals `
            -Values $oorValues `
            -LastRow $oorLastRow `
            -SkuColumn 3 `
            -SoldToHeaderRow 8 `
            -FirstValueColumn 4 `
            -LastValueColumn 10 `
            -IncludedSoldTo @('5395799')
        $columnU = Get-PivotBlockTotals `
            -Values $oorValues `
            -LastRow $oorLastRow `
            -SkuColumn 15 `
            -SoldToHeaderRow 8 `
            -FirstValueColumn 16 `
            -LastValueColumn 22 `
            -IncludedSoldTo @('5395800', '5395799', '5395801')
        $columnV = Get-PivotBlockTotals `
            -Values $oorValues `
            -LastRow $oorLastRow `
            -SkuColumn 3 `
            -SoldToHeaderRow 8 `
            -FirstValueColumn 4 `
            -LastValueColumn 10 `
            -IncludedSoldTo @('1001089')
    }
    if (-not $MonthlyOnly -and -not $OpenOrdersOnly) {
        if (-not $InventoryOnly) {
            $weekly = Get-WeeklyActuals `
                -Values $ddrValues `
                -LastRow $ddrLastRow `
                -RunDate $AsOfDate
        }
        if (-not $WeeklyOnly) {
            $inventory = Get-InventoryTotals `
                -Values $inventoryValues `
                -LastRow $inventoryLastRow
        }
    }
    if (-not $OpenOrdersOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        $monthly = Get-MonthlyActuals -Values $ddrValues -LastRow $ddrLastRow
    }
    $targetRowCount = $targetLastRow - $targetStartRow + 1
    if (-not $OpenOrdersOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        $monthlyOutput = New-Object 'object[,]' $targetRowCount, 3
    }
    if (-not $MonthlyOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        $openOrderOutput = New-Object 'object[,]' $targetRowCount, 4
    }
    if (-not $MonthlyOnly -and -not $OpenOrdersOnly) {
        $stockWeeklyColumnCount = if ($InventoryOnly -or $WeeklyOnly) { 1 } else { 2 }
        $stockWeeklyOutput = New-Object 'object[,]' $targetRowCount, $stockWeeklyColumnCount
    }

    for ($outputRow = 0; $outputRow -lt $targetRowCount; $outputRow++) {
        $productId = Get-NormalizedKey `
            -Value $targetProductValues[($outputRow + 1), 1]
        if (-not $OpenOrdersOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
            for ($monthIndex = 0; $monthIndex -lt 3; $monthIndex++) {
                $monthValue = [double]0
                if ($monthIndex -lt $monthly.MonthOrder.Count) {
                    $month = $monthly.MonthOrder[$monthIndex]
                    $monthValue = Get-MapValue `
                        -Map $monthly.MonthMaps[$month] `
                        -Key $productId
                }
                $monthlyOutput[$outputRow, $monthIndex] = $monthValue
            }
        }
        if (-not $MonthlyOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
            $openOrderOutput[$outputRow, 0] =
                Get-MapValue -Map $columnS -Key $productId
            $openOrderOutput[$outputRow, 1] =
                Get-MapValue -Map $columnT -Key $productId
            $openOrderOutput[$outputRow, 2] =
                Get-MapValue -Map $columnU -Key $productId
            $openOrderOutput[$outputRow, 3] =
                Get-MapValue -Map $columnV -Key $productId
        }
        if (-not $MonthlyOnly -and -not $OpenOrdersOnly) {
            if ($WeeklyOnly) {
                $stockWeeklyOutput[$outputRow, 0] =
                    Get-MapValue -Map $weekly.Totals -Key $productId
            } else {
                $stockWeeklyOutput[$outputRow, 0] =
                    Get-MapValue -Map $inventory -Key $productId
                if (-not $InventoryOnly) {
                    $stockWeeklyOutput[$outputRow, 1] =
                        Get-MapValue -Map $weekly.Totals -Key $productId
                }
            }
        }
    }

    if (-not $OpenOrdersOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        $monthlyTarget = $targetSheet.Range(
            "$($monthlyColumns[0])${targetStartRow}:$($monthlyColumns[1])${targetLastRow}"
        )
    }
    if (-not $MonthlyOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        $openOrderTarget = $targetSheet.Range(
            "$($openOrderColumns[0])${targetStartRow}:$($openOrderColumns[1])${targetLastRow}"
        )
    }
    if (-not $MonthlyOnly -and -not $OpenOrdersOnly) {
        $stockWeeklyTarget = $targetSheet.Range(
            "$($stockWeeklyColumns[0])${targetStartRow}:$($stockWeeklyColumns[1])${targetLastRow}"
        )
    }
    $rowRuns = Get-RowVisibilityRuns `
        -Worksheet $targetSheet `
        -StartRow $targetStartRow `
        -EndRow $targetLastRow

    if (-not $OpenOrdersOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        Set-RangeValues `
            -Worksheet $targetSheet `
            -Range $monthlyTarget `
            -Values $monthlyOutput `
            -RowRuns $rowRuns
        if ($sameWorkbook) {
            $monthlyTarget.NumberFormat = '#,##0;-#,##0;;@'
        }
    }
    if (-not $MonthlyOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        Set-RangeValues `
            -Worksheet $targetSheet `
            -Range $openOrderTarget `
            -Values $openOrderOutput `
            -RowRuns $rowRuns
    }
    if (-not $MonthlyOnly -and -not $OpenOrdersOnly) {
        Set-RangeValues `
            -Worksheet $targetSheet `
            -Range $stockWeeklyTarget `
            -Values $stockWeeklyOutput `
            -RowRuns $rowRuns
    }

    Invoke-ExcelComRetry -Context 'restore Automatic calculation before destination save' -Operation {
        $excel.Calculation = -4105
    }
    $destinationWorkbook.Save()
    if ($excel.Calculation -ne -4105) {
        throw 'Excel calculation mode is not Automatic after saving the destination workbook.'
    }

    if (-not $OpenOrdersOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        Assert-ArrayValues `
            -Actual $monthlyTarget.Value2 `
            -Expected $monthlyOutput `
            -RangeName "$DestinationSheetName $($monthlyColumns -join ':')"
        if ($sameWorkbook -and
            $monthlyTarget.NumberFormat -ne '#,##0;-#,##0;;@') {
            throw 'CMMR monthly quantities must display numeric zero as blank.'
        }
    }
    if (-not $MonthlyOnly -and -not $InventoryOnly -and -not $WeeklyOnly) {
        Assert-ArrayValues `
            -Actual $openOrderTarget.Value2 `
            -Expected $openOrderOutput `
            -RangeName "$DestinationSheetName $($openOrderColumns -join ':')"
    }
    if (-not $MonthlyOnly -and -not $OpenOrdersOnly) {
        Assert-ArrayValues `
            -Actual $stockWeeklyTarget.Value2 `
            -Expected $stockWeeklyOutput `
            -RangeName "$DestinationSheetName $($stockWeeklyColumns -join ':')"
    }

    $refreshSucceeded = $true
}
finally {
    if ($sameWorkbook -and $null -ne $destinationWorkbook) {
        $sourceWorkbook = $null
    }
    if ($null -ne $destinationWorkbook) {
        try {
            $destinationWorkbook.Close($false)
        }
        catch {
            Write-Warning "Failed to close the destination workbook: $($_.Exception.Message)"
        }
    }
    if ($null -ne $sourceWorkbook) {
        try {
            $sourceWorkbook.Close($false)
        }
        catch {
            Write-Warning "Failed to close the source workbook: $($_.Exception.Message)"
        }
    }
    if ($null -ne $excel) {
        for ($quitAttempt = 1; $quitAttempt -le 2; $quitAttempt++) {
            try {
                $excel.Quit()
            }
            catch {
                Write-Warning "Failed to quit Excel (attempt $quitAttempt): $($_.Exception.Message)"
            }
            if ($null -eq $excelProcess -or
                $excelProcess.HasExited -or
                $excelProcess.WaitForExit(5000)) {
                break
            }
        }
    }

    foreach ($comObject in @(
        $stockWeeklyTarget,
        $openOrderTarget,
        $monthlyTarget,
        $targetProductRange,
        $inventoryRange,
        $ddrRange,
        $oorRange,
        $targetSheet,
        $inventorySheet,
        $ddrPivotSheet,
        $oorSheet,
        $destinationWorkbook,
        $sourceWorkbook,
        $excel
    )) {
        Release-ComObject -ComObject $comObject
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($comFilterRegistered) {
        [CmmrOleMessageFilter]::Revoke()
    }

    if ($null -ne $excelProcess) {
        try {
            if (-not $excelProcess.HasExited -and
                -not $excelProcess.WaitForExit(5000)) {
                Write-Warning "Excel process $excelProcessId is still running after graceful quit retries; it was not force-terminated."
            }
        }
        finally {
            $excelProcess.Dispose()
        }
    }
}

if ($refreshSucceeded) {
    if ($WeeklyOnly) {
        Write-Host (
            "CMMR weekly-only $($stockWeeklyColumns[0]) refresh succeeded. " +
            "Weekly range: $($weekly.StartDate.ToString('yyyy-MM-dd')) through " +
            "$($weekly.EndDate.ToString('yyyy-MM-dd'))."
        )
    }
    elseif ($InventoryOnly) {
        Write-Host "CMMR inventory-only $($stockWeeklyColumns[0]) refresh succeeded."
    }
    elseif ($OpenOrdersOnly) {
        Write-Host "CMMR open-orders-only $($openOrderColumns -join ':') refresh succeeded."
    }
    elseif ($MonthlyOnly) {
        Write-Host (
            "CMMR monthly-only $($monthlyColumns -join ':') refresh succeeded. " +
            "Displayed months: $($monthly.MonthOrder -join ', ')."
        )
    }
    else {
        Write-Host (
            'CMMR supportability report refresh succeeded. ' +
            "Weekly range: $($weekly.StartDate.ToString('yyyy-MM-dd')) through " +
            "$($weekly.EndDate.ToString('yyyy-MM-dd'))."
        )
    }
}
