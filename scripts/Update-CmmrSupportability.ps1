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

function Get-NormalizedIdentifier {
    param(
        [AllowNull()]
        [object]$Value
    )

    $text = Get-NormalizedKey -Value $Value
    if ($text -match '^\d+$') {
        $trimmed = $text.TrimStart('0')
        if ([string]::IsNullOrEmpty($trimmed)) {
            return '0'
        }
        return $trimmed
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

function Get-InventoryAvailableSupply {
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
        $material = Get-NormalizedIdentifier -Value $Values[$rowIndex, 1]
        if ([string]::IsNullOrWhiteSpace($material)) {
            continue
        }

        $totals[$material] = ConvertTo-Number `
            -Value $Values[$rowIndex, 9] `
            -Context "CMMR INV row $rowIndex, column I"
    }

    return ,$totals
}

function Get-CurrentCalendarMonday {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$RunDate
    )

    $date = $RunDate.Date
    $daysSinceMonday = (([int]$date.DayOfWeek + 6) % 7)
    return $date.AddDays(-$daysSinceMonday)
}

function Get-SkuDateKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Sku,

        [Parameter(Mandatory = $true)]
        [datetime]$Date
    )

    return "$Sku|$($Date.ToString('yyyy-MM-dd'))"
}

function Get-IbPoConfirmationTotals {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Values
    )

    $rowCount = $Values.GetLength(0)
    $columnCount = $Values.GetLength(1)
    $headerRow = 0
    $materialColumn = 0
    for ($rowIndex = 1; $rowIndex -le $rowCount; $rowIndex++) {
        for ($columnIndex = 1;
            $columnIndex -le $columnCount;
            $columnIndex++) {
            if ((Get-NormalizedKey -Value $Values[
                $rowIndex,
                $columnIndex
            ]) -eq 'MATERIAL') {
                $headerRow = $rowIndex
                $materialColumn = $columnIndex
                break
            }
        }
        if ($headerRow -ne 0) {
            break
        }
    }
    if ($headerRow -eq 0) {
        throw "IB PO pvt does not contain a 'Material' header."
    }

    $dateColumns = @{}
    for ($columnIndex = $materialColumn + 1;
        $columnIndex -le $columnCount;
        $columnIndex++) {
        $header = Get-NormalizedKey -Value $Values[
            $headerRow,
            $columnIndex
        ]
        if ($header -eq 'GRAND TOTAL') {
            break
        }

        $headerValue = $Values[$headerRow, $columnIndex]
        if ($headerValue -is [datetime] -or
            $headerValue -is [byte] -or
            $headerValue -is [int16] -or
            $headerValue -is [int32] -or
            $headerValue -is [int64] -or
            $headerValue -is [single] -or
            $headerValue -is [double] -or
            $headerValue -is [decimal]) {
            $dateColumns[$columnIndex] = ConvertTo-ExcelDate `
                -Value $headerValue `
                -Context "IB PO pvt header column $columnIndex"
        }
    }
    if ($dateColumns.Count -eq 0) {
        throw 'IB PO pvt does not contain any date columns.'
    }

    $totals = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $material = ''
    for ($rowIndex = $headerRow + 1; $rowIndex -le $rowCount; $rowIndex++) {
        $rowMaterial = Get-NormalizedIdentifier `
            -Value $Values[$rowIndex, $materialColumn]
        if ($rowMaterial -eq 'GRAND TOTAL') {
            break
        }
        if (-not [string]::IsNullOrWhiteSpace($rowMaterial)) {
            $material = $rowMaterial
        }
        if ([string]::IsNullOrWhiteSpace($material)) {
            continue
        }

        foreach ($columnIndex in $dateColumns.Keys) {
            $quantity = ConvertTo-Number `
                -Value $Values[$rowIndex, $columnIndex] `
                -Context "IB PO pvt row $rowIndex, column $columnIndex"
            $key = Get-SkuDateKey `
                -Sku $material `
                -Date $dateColumns[$columnIndex]
            Add-ToMap -Map $totals -Key $key -Value $quantity
        }
    }

    return ,$totals
}

function Get-ConfirmationWeeklyBuckets {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [int]$LastRow,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]$IbPoTotals,

        [Parameter(Mandatory = $true)]
        [System.Collections.Hashtable]$AvailableSupply,

        [Parameter(Mandatory = $true)]
        [datetime]$RunDate
    )

    $expectedHeaders = @{
        3 = 'SOLDTO NBR'
        24 = 'SKU'
        30 = 'TOTAL CONFIRMED QTY'
        45 = 'PLANT'
        76 = 'DATE TO RELEASE'
    }
    foreach ($columnIndex in $expectedHeaders.Keys) {
        $actualHeader = Get-NormalizedKey -Value $Values[1, $columnIndex]
        if ($actualHeader -ne $expectedHeaders[$columnIndex]) {
            throw "CMMR confirmation column $columnIndex should be '$($expectedHeaders[$columnIndex])', but found '$actualHeader'."
        }
    }

    $currentMonday = Get-CurrentCalendarMonday -RunDate $RunDate
    $confirmationTotals = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $dateByKey = @{}
    $skuByKey = @{}

    for ($rowIndex = 2; $rowIndex -le $LastRow; $rowIndex++) {
        $plant = Get-NormalizedIdentifier -Value $Values[$rowIndex, 45]
        $soldTo = Get-NormalizedSoldTo -Value $Values[$rowIndex, 3]
        if ($plant -ne '4815' -or $soldTo -notin $soldToParties) {
            continue
        }

        $sku = Get-NormalizedIdentifier -Value $Values[$rowIndex, 24]
        if ([string]::IsNullOrWhiteSpace($sku)) {
            continue
        }

        $releaseValue = $Values[$rowIndex, 76]
        $releaseText = Get-NormalizedKey -Value $releaseValue
        if ([string]::IsNullOrWhiteSpace($releaseText) -or
            $releaseText -in @('TBD', 'DELIVERED')) {
            continue
        }
        $releaseDate = ConvertTo-ExcelDate `
            -Value $releaseValue `
            -Context "CMMR confirmation row $rowIndex, column BX"
        if ($releaseDate -lt $currentMonday) {
            continue
        }

        $quantity = ConvertTo-Number `
            -Value $Values[$rowIndex, 30] `
            -Context "CMMR confirmation row $rowIndex, column AD"
        $key = Get-SkuDateKey -Sku $sku -Date $releaseDate
        Add-ToMap -Map $confirmationTotals -Key $key -Value $quantity
        $dateByKey[$key] = $releaseDate
        $skuByKey[$key] = $sku
    }

    $currentWeekMap = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $currentWeekTotals = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $futureWeekTotals = @{}
    foreach ($entry in $confirmationTotals.GetEnumerator()) {
        $key = [string]$entry.Key
        $sku = [string]$skuByKey[$key]
        $releaseDate = [datetime]$dateByKey[$key]
        $quantity = [double]$entry.Value
        $weekMonday = Get-CurrentCalendarMonday -RunDate $releaseDate
        $weekFriday = $weekMonday.AddDays(4)

        if ($weekMonday -eq $currentMonday) {
            Add-ToMap `
                -Map $currentWeekTotals `
                -Key $sku `
                -Value $quantity
            continue
        }

        $weekKey = $weekFriday.ToString('yyyy-MM-dd')
        if (-not $futureWeekTotals.ContainsKey($weekKey)) {
            $futureWeekTotals[$weekKey] =
                [System.Collections.Hashtable]::new(
                    [System.StringComparer]::OrdinalIgnoreCase
                )
        }
        Add-ToMap -Map $futureWeekTotals[$weekKey] -Key $sku -Value $quantity
    }

    $currentFriday = $currentMonday.AddDays(4)
    foreach ($entry in $currentWeekTotals.GetEnumerator()) {
        $sku = [string]$entry.Key
        $quantity = [double]$entry.Value
        $available = Get-MapValue -Map $AvailableSupply -Key $sku
        $ibPoFridayKey = Get-SkuDateKey -Sku $sku -Date $currentFriday
        $ibPoFridayQuantity = Get-MapValue `
            -Map $IbPoTotals `
            -Key $ibPoFridayKey
        if ($available -ge $quantity -or
            $ibPoFridayQuantity -ge $quantity) {
            $currentWeekMap[$sku] = $quantity
        }
    }

    $futureFridays = @(
        $futureWeekTotals.Keys |
            ForEach-Object {
                [datetime]::ParseExact(
                    $_,
                    'yyyy-MM-dd',
                    [System.Globalization.CultureInfo]::InvariantCulture
                )
            } |
            Sort-Object |
            Select-Object -First 4
    )
    $weekMaps = @($currentWeekMap)
    foreach ($futureFriday in $futureFridays) {
        $futureWeekMap = [System.Collections.Hashtable]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
        $weekKey = $futureFriday.ToString('yyyy-MM-dd')
        foreach ($entry in $futureWeekTotals[$weekKey].GetEnumerator()) {
            $sku = [string]$entry.Key
            $quantity = [double]$entry.Value
            $ibPoKey = Get-SkuDateKey -Sku $sku -Date $futureFriday
            $ibPoQuantity = Get-MapValue -Map $IbPoTotals -Key $ibPoKey
            if ($ibPoQuantity -ge $quantity) {
                $futureWeekMap[$sku] = $quantity
            }
        }
        $weekMaps += $futureWeekMap
    }
    while ($weekMaps.Count -lt 5) {
        $weekMaps += [System.Collections.Hashtable]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
    }

    return [pscustomobject]@{
        WeekMaps = $weekMaps
        StartDate = $currentMonday
        FutureFridays = $futureFridays
    }
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
$ibPoSourcePath = Join-Path `
    -Path (Split-Path -Parent $sourcePath) `
    -ChildPath 'HK ALL IB PO.XLSX'
if (-not (Test-Path -LiteralPath $ibPoSourcePath -PathType Leaf)) {
    throw "IB PO source workbook not found: $ibPoSourcePath"
}
$ibPoSourcePath = (Resolve-Path -LiteralPath $ibPoSourcePath).Path
$excel = $null
$sourceWorkbook = $null
$destinationWorkbook = $null
$ibPoSourceWorkbook = $null
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

    $ibPoSourceWorkbook = $excel.Workbooks.Open(
        $ibPoSourcePath,
        0,
        $true
    )
    $sourceWorkbook = $excel.Workbooks.Open($sourcePath, 0, $false)
    if ([bool]$sourceWorkbook.ReadOnly) {
        throw "Source workbook opened as read-only and cannot update IB PO: $sourcePath"
    }

    $ibPoRawSheet = $ibPoSourceWorkbook.Worksheets.Item(1)
    $ibPoLastRow = Get-LastDataRow -Worksheet $ibPoRawSheet -Column 16
    if ($ibPoLastRow -lt 2) {
        throw "IB PO source worksheet has no data rows: $ibPoSourcePath"
    }

    $ibPoHeaderCell = $ibPoRawSheet.Cells.Item(1, 23)
    $ibPoHeaderCell.Value2 = 'IB PO RDD'
    $ibPoRddRange = $ibPoRawSheet.Range("W2:W$ibPoLastRow")
    $ibPoRddRange.FormulaR1C1 =
        '=IF(RC[-7]="","",RC[-7]-WEEKDAY(RC[-7],2)+5)'
    $ibPoRddRange.Calculate()
    $ibPoRddRange.Value2 = $ibPoRddRange.Value2
    $ibPoRddRange.NumberFormat = 'yyyy-mm-dd'

    for ($sheetIndex = $sourceWorkbook.Worksheets.Count;
        $sheetIndex -ge 1;
        $sheetIndex--) {
        $candidateSheet = $sourceWorkbook.Worksheets.Item($sheetIndex)
        try {
            if ($candidateSheet.Name -eq 'IB PO') {
                $candidateSheet.Delete()
            }
        }
        finally {
            Release-ComObject -ComObject $candidateSheet
        }
    }

    $lastSourceSheet = $sourceWorkbook.Worksheets.Item(
        $sourceWorkbook.Worksheets.Count
    )
    try {
        $ibPoRawSheet.Copy([Type]::Missing, $lastSourceSheet)
    }
    finally {
        Release-ComObject -ComObject $lastSourceSheet
    }
    $ibPoReplacementSheet = $sourceWorkbook.ActiveSheet
    $ibPoReplacementSheet.Name = 'IB PO'

    $ibPoPivotSheetForRefresh =
        $sourceWorkbook.Worksheets.Item('IB PO pvt')
    $ibPoPivotForRefresh =
        $ibPoPivotSheetForRefresh.PivotTables().Item(1)
    if (-not [bool]$ibPoPivotForRefresh.RefreshTable()) {
        throw 'Failed to refresh the IB PO pvt PivotTable.'
    }

    $sourceWorkbook.Save()
    $sourceWorkbook.Close($false)
    $ibPoSourceWorkbook.Close($false)
    foreach ($comObject in @(
        $ibPoPivotForRefresh,
        $ibPoPivotSheetForRefresh,
        $ibPoReplacementSheet,
        $ibPoRddRange,
        $ibPoHeaderCell,
        $ibPoRawSheet,
        $sourceWorkbook,
        $ibPoSourceWorkbook
    )) {
        Release-ComObject -ComObject $comObject
    }
    $ibPoPivotForRefresh = $null
    $ibPoPivotSheetForRefresh = $null
    $ibPoReplacementSheet = $null
    $ibPoRddRange = $null
    $ibPoHeaderCell = $null
    $ibPoRawSheet = $null
    $sourceWorkbook = $null
    $ibPoSourceWorkbook = $null

    $sourceWorkbook = $excel.Workbooks.Open($sourcePath, 0, $true)
    $destinationWorkbook = $excel.Workbooks.Open($destinationPath, 0, $false)
    if ([bool]$destinationWorkbook.ReadOnly) {
        throw "Destination workbook opened as read-only: $destinationPath"
    }

    $oorSheet = $sourceWorkbook.Worksheets.Item('HK OOR Pivot')
    $ddrPivotSheet = $sourceWorkbook.Worksheets.Item('DDR pivot')
    $inventorySheet = $sourceWorkbook.Worksheets.Item('CMMR INV')
    $confirmationSheet = $sourceWorkbook.Worksheets.Item('CMMR confirmation')
    $ibPoSheet = $sourceWorkbook.Worksheets.Item('IB PO pvt')
    $targetSheet = $destinationWorkbook.Worksheets.Item('Supportability')

    $oorLastRow = [math]::Max(
        (Get-LastDataRow -Worksheet $oorSheet -Column 3),
        (Get-LastDataRow -Worksheet $oorSheet -Column 16)
    )
    $oorRange = $oorSheet.Range("A1:V$oorLastRow")
    $oorValues = $oorRange.Value2

    $ddrLastRow = [math]::Max(
        (Get-LastDataRow -Worksheet $ddrPivotSheet -Column 2),
        (Get-LastDataRow -Worksheet $ddrPivotSheet -Column 36)
    )
    $ddrRange = $ddrPivotSheet.Range("A1:CZ$ddrLastRow")
    $ddrValues = $ddrRange.Value2

    $inventoryLastRow = Get-LastDataRow -Worksheet $inventorySheet -Column 1
    $inventoryRange = $inventorySheet.Range("A1:J$inventoryLastRow")
    $inventoryValues = $inventoryRange.Value2

    $confirmationLastRow = Get-LastDataRow `
        -Worksheet $confirmationSheet `
        -Column 24
    $confirmationRange = $confirmationSheet.Range(
        "A1:BX$confirmationLastRow"
    )
    $confirmationValues = $confirmationRange.Value2

    $ibPoPlantRange = $ibPoSheet.Range('A1:B3')
    $ibPoPlantValues = $ibPoPlantRange.Value2
    $ibPoPlant = ''
    for ($rowIndex = 1; $rowIndex -le 3; $rowIndex++) {
        if ((Get-NormalizedKey -Value $ibPoPlantValues[$rowIndex, 1]) -eq
            'PLANT') {
            $ibPoPlant = Get-NormalizedIdentifier `
                -Value $ibPoPlantValues[$rowIndex, 2]
            break
        }
    }
    if ($ibPoPlant -ne '4815') {
        throw "IB PO pvt must be filtered for Plant 4815, but found '$ibPoPlant'."
    }
    $ibPoPivotTable = $ibPoSheet.PivotTables().Item(1)
    $ibPoRange = $ibPoPivotTable.TableRange1
    $ibPoValues = $ibPoRange.Value2

    $targetLastRow = Get-LastDataRow -Worksheet $targetSheet -Column 12
    if ($targetLastRow -lt $targetStartRow) {
        throw 'The destination Supportability sheet has no product rows.'
    }
    $targetProductRange = $targetSheet.Range(
        "L${targetStartRow}:L${targetLastRow}"
    )
    $targetProductValues = $targetProductRange.Value2

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
        -SkuColumn 16 `
        -SoldToHeaderRow 8 `
        -FirstValueColumn 17 `
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
    $monthly = Get-MonthlyActuals -Values $ddrValues -LastRow $ddrLastRow
    $weekly = Get-WeeklyActuals `
        -Values $ddrValues `
        -LastRow $ddrLastRow `
        -RunDate $AsOfDate
    $inventory = Get-InventoryTotals `
        -Values $inventoryValues `
        -LastRow $inventoryLastRow
    $availableSupply = Get-InventoryAvailableSupply `
        -Values $inventoryValues `
        -LastRow $inventoryLastRow
    $ibPoTotals = Get-IbPoConfirmationTotals -Values $ibPoValues
    $confirmationWeekly = Get-ConfirmationWeeklyBuckets `
        -Values $confirmationValues `
        -LastRow $confirmationLastRow `
        -IbPoTotals $ibPoTotals `
        -AvailableSupply $availableSupply `
        -RunDate $AsOfDate

    $targetRowCount = $targetLastRow - $targetStartRow + 1
    $monthlyOutput = New-Object 'object[,]' $targetRowCount, 3
    $openOrderOutput = New-Object 'object[,]' $targetRowCount, 4
    $confirmationOutput = New-Object 'object[,]' $targetRowCount, 5
    $stockWeeklyOutput = New-Object 'object[,]' $targetRowCount, 2

    for ($outputRow = 0; $outputRow -lt $targetRowCount; $outputRow++) {
        $productId = Get-NormalizedKey `
            -Value $targetProductValues[($outputRow + 1), 1]

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

        $openOrderOutput[$outputRow, 0] =
            Get-MapValue -Map $columnS -Key $productId
        $openOrderOutput[$outputRow, 1] =
            Get-MapValue -Map $columnT -Key $productId
        $openOrderOutput[$outputRow, 2] =
            Get-MapValue -Map $columnU -Key $productId
        $openOrderOutput[$outputRow, 3] =
            Get-MapValue -Map $columnV -Key $productId
        $stockWeeklyOutput[$outputRow, 0] =
            Get-MapValue -Map $inventory -Key $productId
        $stockWeeklyOutput[$outputRow, 1] =
            Get-MapValue -Map $weekly.Totals -Key $productId
        $confirmationProductId = Get-NormalizedIdentifier `
            -Value $targetProductValues[($outputRow + 1), 1]
        for ($weekIndex = 0; $weekIndex -lt 5; $weekIndex++) {
            $confirmationOutput[$outputRow, $weekIndex] = Get-MapValue `
                -Map $confirmationWeekly.WeekMaps[$weekIndex] `
                -Key $confirmationProductId
        }
    }

    $monthlyTarget = $targetSheet.Range(
        "O${targetStartRow}:Q${targetLastRow}"
    )
    $openOrderTarget = $targetSheet.Range(
        "S${targetStartRow}:V${targetLastRow}"
    )
    $confirmationTarget = $targetSheet.Range(
        "X${targetStartRow}:AB${targetLastRow}"
    )
    $futureConfirmationTarget = $targetSheet.Range(
        "Y${targetStartRow}:AB${targetLastRow}"
    )
    $confirmationHeaderTarget = $targetSheet.Range('Y3:AB3')
    $confirmationHeaderOutput = New-Object 'object[,]' 1, 4
    for ($headerIndex = 0; $headerIndex -lt 4; $headerIndex++) {
        if ($headerIndex -lt $confirmationWeekly.FutureFridays.Count) {
            $confirmationHeaderOutput[0, $headerIndex] =
                $confirmationWeekly.FutureFridays[$headerIndex].ToOADate()
        }
        else {
            $confirmationHeaderOutput[0, $headerIndex] = $null
        }
    }
    $stockWeeklyTarget = $targetSheet.Range(
        "AL${targetStartRow}:AM${targetLastRow}"
    )
    $rowRuns = Get-RowVisibilityRuns `
        -Worksheet $targetSheet `
        -StartRow $targetStartRow `
        -EndRow $targetLastRow

    Set-RangeValues `
        -Worksheet $targetSheet `
        -Range $monthlyTarget `
        -Values $monthlyOutput `
        -RowRuns $rowRuns
    Set-RangeValues `
        -Worksheet $targetSheet `
        -Range $openOrderTarget `
        -Values $openOrderOutput `
        -RowRuns $rowRuns
    Set-RangeValues `
        -Worksheet $targetSheet `
        -Range $confirmationTarget `
        -Values $confirmationOutput `
        -RowRuns $rowRuns
    $futureConfirmationTarget.NumberFormat = '#,##0;-#,##0;;@'
    $confirmationHeaderTarget.Value2 = $confirmationHeaderOutput
    Set-RangeValues `
        -Worksheet $targetSheet `
        -Range $stockWeeklyTarget `
        -Values $stockWeeklyOutput `
        -RowRuns $rowRuns

    $destinationWorkbook.Save()

    Assert-ArrayValues `
        -Actual $monthlyTarget.Value2 `
        -Expected $monthlyOutput `
        -RangeName 'Supportability O:Q'
    Assert-ArrayValues `
        -Actual $openOrderTarget.Value2 `
        -Expected $openOrderOutput `
        -RangeName 'Supportability S:V'
    Assert-ArrayValues `
        -Actual $confirmationTarget.Value2 `
        -Expected $confirmationOutput `
        -RangeName 'Supportability X:AB'
    Assert-ArrayValues `
        -Actual $confirmationHeaderTarget.Value2 `
        -Expected $confirmationHeaderOutput `
        -RangeName 'Supportability Y3:AB3'
    Assert-ArrayValues `
        -Actual $stockWeeklyTarget.Value2 `
        -Expected $stockWeeklyOutput `
        -RangeName 'Supportability AL:AM'

    $refreshSucceeded = $true
}
finally {
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
    if ($null -ne $ibPoSourceWorkbook) {
        try {
            $ibPoSourceWorkbook.Close($false)
        }
        catch {
            Write-Warning "Failed to close the IB PO source workbook: $($_.Exception.Message)"
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Warning "Failed to quit Excel: $($_.Exception.Message)"
        }
    }

    foreach ($comObject in @(
        $stockWeeklyTarget,
        $confirmationHeaderTarget,
        $futureConfirmationTarget,
        $confirmationTarget,
        $openOrderTarget,
        $monthlyTarget,
        $targetProductRange,
        $ibPoRange,
        $ibPoPivotTable,
        $ibPoPlantRange,
        $confirmationRange,
        $inventoryRange,
        $ddrRange,
        $oorRange,
        $targetSheet,
        $ibPoSheet,
        $confirmationSheet,
        $inventorySheet,
        $ddrPivotSheet,
        $oorSheet,
        $destinationWorkbook,
        $sourceWorkbook,
        $ibPoSourceWorkbook,
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
                $excelProcess.Kill()
                $excelProcess.WaitForExit()
                Write-Warning "Terminated the script-owned Excel process $excelProcessId."
            }
        }
        finally {
            $excelProcess.Dispose()
        }
    }
}

if ($refreshSucceeded) {
    Write-Host (
        'CMMR supportability report refresh succeeded. ' +
        "Weekly range: $($weekly.StartDate.ToString('yyyy-MM-dd')) through " +
        "$($weekly.EndDate.ToString('yyyy-MM-dd')); confirmation buckets: " +
        "$($confirmationWeekly.StartDate.ToString('yyyy-MM-dd')) current week" +
        $(if ($confirmationWeekly.FutureFridays.Count -gt 0) {
            '; future Fridays: ' +
            (($confirmationWeekly.FutureFridays | ForEach-Object {
                $_.ToString('yyyy-MM-dd')
            }) -join ', ') + '.'
        }
        else {
            '; no future weeks with AD data.'
        })
    )
}
