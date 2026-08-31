[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath
)

$ErrorActionPreference = 'Stop'

$outputHeaders = @(
    'Material',
    'Description',
    'Plant',
    'SLoc',
    'QtyUnrestr',
    'QtyOpenDeliveries',
    'QtyOpenOrders',
    'DTV Supply',
    'DTV Supply without open DO',
    'DTV Supply without open DO & Open Order'
)

function ConvertFrom-DelimitedField {
    param(
        [AllowEmptyString()]
        [string]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    $normalized = $Value.Trim()
    if ($normalized.Length -ge 2 -and
        $normalized[0] -eq '"' -and
        $normalized[$normalized.Length - 1] -eq '"') {
        $normalized = $normalized.Substring(1, $normalized.Length - 2).Replace('""', '"')
    }

    return $normalized.Trim()
}

function ConvertFrom-SapNumber {
    param(
        [AllowEmptyString()]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$ColumnName,

        [Parameter(Mandatory = $true)]
        [int]$LineNumber,

        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $normalized = ConvertFrom-DelimitedField -Value $Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return [decimal]0
    }

    $normalized = $normalized.Replace([char]0x00A0, ' ').Replace(' ', '')
    if ($normalized.StartsWith('(') -and $normalized.EndsWith(')')) {
        $normalized = '-' + $normalized.Substring(1, $normalized.Length - 2)
    }

    $number = [decimal]0
    $styles = [System.Globalization.NumberStyles]::Float -bor
        [System.Globalization.NumberStyles]::AllowThousands
    $parsed = [decimal]::TryParse(
        $normalized,
        $styles,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )

    if (-not $parsed) {
        throw "Invalid SAP number '$Value' in column '$ColumnName' at line $LineNumber in '$SourcePath'."
    }

    return $number
}

function Get-FieldValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string[]]$Fields,

        [Parameter(Mandatory = $true)]
        [int]$Index
    )

    if ($Index -ge $Fields.Length) {
        return ''
    }

    return ConvertFrom-DelimitedField -Value $Fields[$Index]
}

function Read-InventorySource {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $aggregates = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $materialOrder = [System.Collections.Generic.List[string]]::new()
    $reader = $null

    try {
        $reader = [System.IO.StreamReader]::new(
            $SourcePath,
            [System.Text.Encoding]::Default,
            $true
        )

        $headerLine = $null
        $lineNumber = 0
        while (-not $reader.EndOfStream -and [string]::IsNullOrWhiteSpace($headerLine)) {
            $headerLine = $reader.ReadLine()
            $lineNumber++
        }

        if ([string]::IsNullOrWhiteSpace($headerLine)) {
            throw "Source file is empty: $SourcePath"
        }

        $headerFields = $headerLine.Split([char]9)
        $headerIndexes = [System.Collections.Hashtable]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )

        for ($columnIndex = 0; $columnIndex -lt $headerFields.Length; $columnIndex++) {
            $headerName = ConvertFrom-DelimitedField -Value $headerFields[$columnIndex]
            if ($columnIndex -eq 0) {
                $headerName = $headerName.TrimStart([char]0xFEFF)
            }

            if (-not [string]::IsNullOrWhiteSpace($headerName)) {
                $headerIndexes[$headerName] = $columnIndex
            }
        }

        $requiredHeaders = @(
            'Material',
            'Description',
            'Plant',
            'SLoc',
            'SpecStkInd',
            'QtyUnrestr',
            'QtyOpenDeliveries',
            'QtyOpenOrders'
        )

        foreach ($requiredHeader in $requiredHeaders) {
            if (-not $headerIndexes.ContainsKey($requiredHeader)) {
                throw "Required column '$requiredHeader' was not found in '$SourcePath'."
            }
        }

        while (-not $reader.EndOfStream) {
            $line = $reader.ReadLine()
            $lineNumber++

            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            $fields = $line.Split([char]9)
            $specStkInd = Get-FieldValue `
                -Fields $fields `
                -Index $headerIndexes['SpecStkInd']
            if (-not [string]::IsNullOrWhiteSpace($specStkInd)) {
                continue
            }

            $material = Get-FieldValue -Fields $fields -Index $headerIndexes['Material']
            if ([string]::IsNullOrWhiteSpace($material)) {
                continue
            }

            if (-not $aggregates.ContainsKey($material)) {
                $aggregates[$material] = [pscustomobject]@{
                    Material = $material
                    Description = Get-FieldValue `
                        -Fields $fields `
                        -Index $headerIndexes['Description']
                    Plant = Get-FieldValue `
                        -Fields $fields `
                        -Index $headerIndexes['Plant']
                    HasTargetSloc = $false
                    QtyUnrestr = [decimal]0
                    QtyOpenDeliveries = [decimal]0
                    QtyOpenOrders = [decimal]0
                }
                $materialOrder.Add($material)
            }

            $aggregate = $aggregates[$material]
            $aggregate.QtyOpenDeliveries += ConvertFrom-SapNumber `
                -Value (Get-FieldValue `
                    -Fields $fields `
                    -Index $headerIndexes['QtyOpenDeliveries']) `
                -ColumnName 'QtyOpenDeliveries' `
                -LineNumber $lineNumber `
                -SourcePath $SourcePath
            $aggregate.QtyOpenOrders += ConvertFrom-SapNumber `
                -Value (Get-FieldValue `
                    -Fields $fields `
                    -Index $headerIndexes['QtyOpenOrders']) `
                -ColumnName 'QtyOpenOrders' `
                -LineNumber $lineNumber `
                -SourcePath $SourcePath

            $sloc = Get-FieldValue -Fields $fields -Index $headerIndexes['SLoc']
            if ($sloc -match '^0*10$') {
                $aggregate.HasTargetSloc = $true
                $aggregate.QtyUnrestr += ConvertFrom-SapNumber `
                    -Value (Get-FieldValue `
                        -Fields $fields `
                        -Index $headerIndexes['QtyUnrestr']) `
                    -ColumnName 'QtyUnrestr' `
                    -LineNumber $lineNumber `
                    -SourcePath $SourcePath
            }
        }
    }
    finally {
        if ($null -ne $reader) {
            $reader.Dispose()
        }
    }

    $outputRows = [System.Collections.Generic.List[object]]::new()
    foreach ($material in $materialOrder) {
        $aggregate = $aggregates[$material]
        $hasOpenQuantity = $aggregate.QtyOpenDeliveries -ne 0 -or
            $aggregate.QtyOpenOrders -ne 0
        if (-not $aggregate.HasTargetSloc -and -not $hasOpenQuantity) {
            continue
        }

        $qtyUnrestr = if ($aggregate.HasTargetSloc) {
            $aggregate.QtyUnrestr
        }
        else {
            [decimal]0
        }

        $outputRows.Add([pscustomobject]@{
            Material = $aggregate.Material
            Description = $aggregate.Description
            Plant = $aggregate.Plant
            SLoc = 10
            QtyUnrestr = $qtyUnrestr
            QtyOpenDeliveries = $aggregate.QtyOpenDeliveries
            QtyOpenOrders = $aggregate.QtyOpenOrders
            DtvSupply = $qtyUnrestr
            DtvSupplyWithoutOpenDo = $qtyUnrestr - $aggregate.QtyOpenDeliveries
            DtvSupplyWithoutOpenDoAndOpenOrder = $qtyUnrestr + $aggregate.QtyOpenOrders
        })
    }

    return ,$outputRows
}

function ConvertTo-WorksheetArray {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[object]]$Rows,

        [Parameter(Mandatory = $true)]
        [string[]]$Headers
    )

    $values = New-Object 'object[,]' ($Rows.Count + 1), $Headers.Count
    for ($columnIndex = 0; $columnIndex -lt $Headers.Count; $columnIndex++) {
        $values[0, $columnIndex] = $Headers[$columnIndex]
    }

    for ($rowIndex = 0; $rowIndex -lt $Rows.Count; $rowIndex++) {
        $row = $Rows[$rowIndex]
        $targetRow = $rowIndex + 1
        $values[$targetRow, 0] = [string]$row.Material
        $values[$targetRow, 1] = [string]$row.Description
        $values[$targetRow, 2] = [string]$row.Plant
        $values[$targetRow, 3] = [double]$row.SLoc
        $values[$targetRow, 4] = [double]$row.QtyUnrestr
        $values[$targetRow, 5] = [double]$row.QtyOpenDeliveries
        $values[$targetRow, 6] = [double]$row.QtyOpenOrders
        $values[$targetRow, 7] = [double]$row.DtvSupply
        $values[$targetRow, 8] = [double]$row.DtvSupplyWithoutOpenDo
        $values[$targetRow, 9] = [double]$row.DtvSupplyWithoutOpenDoAndOpenOrder
    }

    return ,$values
}

function Get-OleColor {
    param(
        [Parameter(Mandatory = $true)]
        [int]$Red,

        [Parameter(Mandatory = $true)]
        [int]$Green,

        [Parameter(Mandatory = $true)]
        [int]$Blue
    )

    return $Red + ($Green * 256) + ($Blue * 65536)
}

function Release-ComObject {
    param(
        [AllowNull()]
        [object]$ComObject
    )

    if ($null -ne $ComObject -and
        [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Register-ExcelComRetryFilter {
    if (-not ('ExcelComRetryFilter' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
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

public sealed class ExcelComRetryFilter : IOleMessageFilter
{
    private const int ServerCallRejected = 1;
    private const int ServerCallRetryLater = 2;
    private const int PendingMessageWaitDefault = 2;
    private const int MaximumRetryDurationMilliseconds = 10000;
    private const int RetryDelayMilliseconds = 250;

    private static ExcelComRetryFilter currentFilter;
    private static IOleMessageFilter previousFilter;

    [DllImport("Ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        IOleMessageFilter newFilter,
        out IOleMessageFilter oldFilter);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(
        IntPtr windowHandle,
        out uint processId);

    public static void Register()
    {
        if (currentFilter != null)
        {
            return;
        }

        currentFilter = new ExcelComRetryFilter();
        IOleMessageFilter oldFilter;
        CoRegisterMessageFilter(currentFilter, out oldFilter);
        previousFilter = oldFilter;
    }

    public static void Revoke()
    {
        IOleMessageFilter ignored;
        CoRegisterMessageFilter(previousFilter, out ignored);
        previousFilter = null;
        currentFilter = null;
    }

    public static int GetProcessId(IntPtr windowHandle)
    {
        uint processId;
        GetWindowThreadProcessId(windowHandle, out processId);
        return unchecked((int)processId);
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
        bool isRetryable =
            rejectType == ServerCallRejected ||
            rejectType == ServerCallRetryLater;

        if (isRetryable && tickCount < MaximumRetryDurationMilliseconds)
        {
            return RetryDelayMilliseconds;
        }

        return -1;
    }

    public int MessagePending(
        IntPtr taskCallee,
        int tickCount,
        int pendingType)
    {
        return PendingMessageWaitDefault;
    }
}

public sealed class ExcelShutdownWatchdog : IDisposable
{
    private readonly Process process;
    private readonly Timer timer;
    private int terminated;

    private ExcelShutdownWatchdog(Process process, int timeoutMilliseconds)
    {
        this.process = process;
        timer = new Timer(OnTimeout, null, timeoutMilliseconds, Timeout.Infinite);
    }

    public bool Terminated
    {
        get { return Interlocked.CompareExchange(ref terminated, 0, 0) == 1; }
    }

    public static ExcelShutdownWatchdog Start(
        Process process,
        int timeoutMilliseconds)
    {
        return new ExcelShutdownWatchdog(process, timeoutMilliseconds);
    }

    private void OnTimeout(object state)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill();
                Interlocked.Exchange(ref terminated, 1);
            }
        }
        catch (InvalidOperationException)
        {
        }
    }

    public void Dispose()
    {
        timer.Dispose();
    }
}
'@
    }

    [ExcelComRetryFilter]::Register()
}

function Save-WorkbookWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Workbook,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookName,

        [int]$MaximumAttempts = 3
    )

    for ($attempt = 1; $attempt -le $MaximumAttempts; $attempt++) {
        try {
            $Workbook.Save()
            return
        }
        catch [System.Runtime.InteropServices.COMException] {
            $hresult = '0x{0:X8}' -f ($_.Exception.HResult -band 0xFFFFFFFFL)
            if ($attempt -eq $MaximumAttempts) {
                throw "Failed to save '$WorkbookName' after $MaximumAttempts attempts. Excel COM error $hresult`: $($_.Exception.Message)"
            }

            Start-Sleep -Seconds (2 * $attempt)
        }
    }
}

function Remove-Worksheet {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Workbook,

        [Parameter(Mandatory = $true)]
        [string]$WorksheetName
    )

    for ($index = $Workbook.Worksheets.Count; $index -ge 1; $index--) {
        $worksheet = $Workbook.Worksheets.Item($index)
        if ($worksheet.Name -eq $WorksheetName) {
            $worksheet.Delete()
        }
    }
}

function Set-WorksheetContent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application,

        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [int]$HeaderColor
    )

    $rowCount = $Values.GetLength(0)
    $columnCount = $Values.GetLength(1)
    $lastCell = $null
    $outputRange = $null
    $headerRange = $null
    $numericRange = $null
    $allColumns = $null
    $groupColumns = $null
    $firstRow = $null
    $filterRange = $null
    $window = $null
    $parentWorkbook = $null

    try {
        $lastCell = $Worksheet.Cells.Item($rowCount, $columnCount)
        $outputRange = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $lastCell)
        $outputRange.Value2 = $Values

        $headerRange = $Worksheet.Range('A1:J1')
        $headerRange.Interior.Color = $HeaderColor
        $headerRange.Font.Bold = $true
        $headerRange.Font.Color = 0
        $headerRange.WrapText = $true

        $numericRange = $Worksheet.Range("D1:J$rowCount")
        $numericRange.NumberFormat = '0'

        $allColumns = $Worksheet.Range('A:J')
        $allColumns.ColumnWidth = 23

        $firstRow = $Worksheet.Rows.Item(1)
        $firstRow.RowHeight = 30

        $filterRange = $Worksheet.Range("A1:J$rowCount")
        $filterRange.AutoFilter() | Out-Null

        $groupColumns = $Worksheet.Range('C:G').EntireColumn
        $groupColumns.Group() | Out-Null
        $groupColumns.Hidden = $true

        $parentWorkbook = $Worksheet.Parent
        $parentWorkbook.Activate()
        $Worksheet.Activate()
        $window = $parentWorkbook.Windows.Item(1)
        $window.FreezePanes = $false
        $window.SplitColumn = 0
        $window.SplitRow = 1
        $window.FreezePanes = $true
    }
    finally {
        foreach ($comObject in @(
            $filterRange,
            $firstRow,
            $groupColumns,
            $allColumns,
            $numericRange,
            $headerRange,
            $outputRange,
            $lastCell
        )) {
            Release-ComObject -ComObject $comObject
        }
    }
}

function Assert-NearlyEqual {
    param(
        [Parameter(Mandatory = $true)]
        [double]$Actual,

        [Parameter(Mandatory = $true)]
        [double]$Expected,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if ([math]::Abs($Actual - $Expected) -gt 0.000001) {
        throw "$Message Expected $Expected but found $Actual."
    }
}

function Assert-Worksheet {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Application,

        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [int]$DataRowCount,

        [Parameter(Mandatory = $true)]
        [int]$ExpectedHeaderColor
    )

    $headerRange = $null
    $numericRange = $null
    $dataRange = $null
    $window = $null
    $firstRow = $null
    $parentWorkbook = $null

    try {
        $headerRange = $Worksheet.Range('A1:J1')
        if (-not [bool]$headerRange.Font.Bold) {
            throw "Worksheet '$($Worksheet.Name)' header is not bold."
        }
        if ([int]$headerRange.Font.Color -ne 0) {
            throw "Worksheet '$($Worksheet.Name)' header font is not black."
        }
        if ([int]$headerRange.Interior.Color -ne $ExpectedHeaderColor) {
            throw "Worksheet '$($Worksheet.Name)' header fill color is incorrect."
        }
        if (-not [bool]$headerRange.WrapText) {
            throw "Worksheet '$($Worksheet.Name)' header wrapping is not enabled."
        }
        if (-not [bool]$Worksheet.AutoFilterMode) {
            throw "Worksheet '$($Worksheet.Name)' does not have filters enabled."
        }

        $parentWorkbook = $Worksheet.Parent
        $parentWorkbook.Activate()
        $Worksheet.Activate()
        $window = $parentWorkbook.Windows.Item(1)
        if (-not [bool]$window.FreezePanes -or [int]$window.SplitRow -ne 1) {
            throw "Worksheet '$($Worksheet.Name)' does not have row 1 frozen."
        }

        for ($columnIndex = 1; $columnIndex -le 10; $columnIndex++) {
            $column = $Worksheet.Columns.Item($columnIndex)
            $wasHidden = [bool]$column.Hidden
            try {
                if ($wasHidden) {
                    $column.Hidden = $false
                }
                Assert-NearlyEqual `
                    -Actual ([double]$column.ColumnWidth) `
                    -Expected 23 `
                    -Message "Worksheet '$($Worksheet.Name)' column $columnIndex has the wrong width."
            }
            finally {
                if ($wasHidden) {
                    $column.Hidden = $true
                }
                Release-ComObject -ComObject $column
            }
        }

        $firstRow = $Worksheet.Rows.Item(1)
        Assert-NearlyEqual `
            -Actual ([double]$firstRow.RowHeight) `
            -Expected 30 `
            -Message "Worksheet '$($Worksheet.Name)' row 1 has the wrong height."

        for ($columnIndex = 3; $columnIndex -le 7; $columnIndex++) {
            $column = $Worksheet.Columns.Item($columnIndex)
            try {
                if ([int]$column.OutlineLevel -lt 2 -or -not [bool]$column.Hidden) {
                    throw "Worksheet '$($Worksheet.Name)' columns C:G are not grouped and collapsed."
                }
            }
            finally {
                Release-ComObject -ComObject $column
            }
        }

        $lastRow = $DataRowCount + 1
        $numericRange = $Worksheet.Range("D1:J$lastRow")
        if ([string]$numericRange.NumberFormat -ne '0') {
            throw "Worksheet '$($Worksheet.Name)' numeric columns do not use a zero-decimal format."
        }

        if ($DataRowCount -eq 0) {
            return
        }

        $dataRange = $Worksheet.Range("D2:J$lastRow")
        $values = $dataRange.Value2
        for ($rowIndex = 1; $rowIndex -le $DataRowCount; $rowIndex++) {
            $sloc = $values[$rowIndex, 1]
            if ($sloc -is [string] -or [double]$sloc -ne 10) {
                throw "Worksheet '$($Worksheet.Name)' row $($rowIndex + 1) has a non-numeric or invalid SLoc value."
            }

            $qtyUnrestr = [double]$values[$rowIndex, 2]
            $qtyOpenDeliveries = [double]$values[$rowIndex, 3]
            $qtyOpenOrders = [double]$values[$rowIndex, 4]
            Assert-NearlyEqual `
                -Actual ([double]$values[$rowIndex, 5]) `
                -Expected $qtyUnrestr `
                -Message "Worksheet '$($Worksheet.Name)' row $($rowIndex + 1) has an invalid DTV Supply value."
            Assert-NearlyEqual `
                -Actual ([double]$values[$rowIndex, 6]) `
                -Expected ($qtyUnrestr - $qtyOpenDeliveries) `
                -Message "Worksheet '$($Worksheet.Name)' row $($rowIndex + 1) has an invalid DTV Supply without open DO value."
            Assert-NearlyEqual `
                -Actual ([double]$values[$rowIndex, 7]) `
                -Expected ($qtyUnrestr + $qtyOpenOrders) `
                -Message "Worksheet '$($Worksheet.Name)' row $($rowIndex + 1) has an invalid DTV Supply without open DO & Open Order value."
        }
    }
    finally {
        foreach ($comObject in @(
            $firstRow,
            $dataRange,
            $numericRange,
            $headerRange
        )) {
            Release-ComObject -ComObject $comObject
        }
    }
}

function Assert-WorksheetNameCount {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Workbook,

        [Parameter(Mandatory = $true)]
        [string]$WorksheetName
    )

    $matchCount = 0
    for ($index = 1; $index -le $Workbook.Worksheets.Count; $index++) {
        $worksheet = $Workbook.Worksheets.Item($index)
        if ($worksheet.Name -eq $WorksheetName) {
            $matchCount++
        }
    }

    if ($matchCount -ne 1) {
        throw "Expected worksheet '$WorksheetName' to exist exactly once, but found $matchCount."
    }
}

function Get-CellText {
    param(
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).Trim()
}

function ConvertFrom-WorksheetNumber {
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

    $normalized = ([string]$Value).Trim().Replace([char]0x00A0, ' ').Replace(' ', '')
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

function Get-OorHeaderIndexes {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object[,]]$Values,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookName
    )

    $headers = [System.Collections.Hashtable]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    for ($columnIndex = 1; $columnIndex -le $Values.GetLength(1); $columnIndex++) {
        $header = Get-CellText -Value $Values[1, $columnIndex]
        if (-not [string]::IsNullOrWhiteSpace($header)) {
            $headers[$header] = $columnIndex
        }
    }

    foreach ($requiredHeader in @(
        'Order Nbr',
        'SKU',
        'Sched Line #',
        'Plant',
        'Delivered Qty',
        'Material Avail.Date',
        'ShipTo Ctry'
    )) {
        if (-not $headers.ContainsKey($requiredHeader)) {
            throw "Required column '$requiredHeader' was not found in '$WorkbookName'."
        }
    }

    return ,$headers
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

function Set-OrAssert-RowColors {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [int[]]$ExpectedColors,

        [Parameter(Mandatory = $true)]
        [int]$LastRow,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookName,

        [switch]$Validate
    )

    $runStart = 0
    $runColor = 0
    for ($rowIndex = 2; $rowIndex -le ($LastRow + 1); $rowIndex++) {
        $color = if ($rowIndex -le $LastRow) {
            $ExpectedColors[$rowIndex]
        }
        else {
            0
        }

        if ($color -eq $runColor) {
            continue
        }

        if ($runColor -ne 0) {
            $rowRange = $Worksheet.Range(
                "A${runStart}:BX$($rowIndex - 1)"
            )
            try {
                if ($Validate) {
                    if ([int]$rowRange.Interior.Color -ne $runColor) {
                        throw "Rows $runStart-$($rowIndex - 1) in '$WorkbookName' have the wrong highlight color."
                    }
                }
                else {
                    $rowRange.Interior.Color = $runColor
                }
            }
            finally {
                Release-ComObject -ComObject $rowRange
            }
        }

        $runStart = if ($color -ne 0) { $rowIndex } else { 0 }
        $runColor = $color
    }
}

function Set-SystemConfirmationContent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Workbook,

        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookName,

        [Parameter(Mandatory = $true)]
        [int]$LightGreen,

        [Parameter(Mandatory = $true)]
        [int]$LightBlue
    )

    $headerRange = $null
    $sourceRange = $null
    $plantRange = $null
    $plantStart = $null
    $plantEnd = $null
    $releaseRange = $null
    $releaseHeader = $null
    $pivotCaches = $null

    try {
        $headerRange = $Worksheet.Range('A1:BX1')
        $headerValues = $headerRange.Value2
        $headerIndexes = Get-OorHeaderIndexes `
            -Values $headerValues `
            -WorkbookName $WorkbookName
        $rowCount = [math]::Max(
            (Get-LastDataRow `
                -Worksheet $Worksheet `
                -Column $headerIndexes['Order Nbr']),
            (Get-LastDataRow `
                -Worksheet $Worksheet `
                -Column $headerIndexes['SKU'])
        )
        if ($rowCount -lt 2) {
            throw "Workbook '$WorkbookName' does not contain any data rows."
        }

        $sourceRange = $Worksheet.Range("A1:BX$rowCount")
        $values = $sourceRange.Value2
        $groups = [System.Collections.Hashtable]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
        $scheduleIsZero = New-Object 'bool[]' ($rowCount + 1)

        for ($rowIndex = 2; $rowIndex -le $rowCount; $rowIndex++) {
            $orderNumber = Get-CellText -Value $values[$rowIndex, $headerIndexes['Order Nbr']]
            $sku = Get-CellText -Value $values[$rowIndex, $headerIndexes['SKU']]
            if ([string]::IsNullOrWhiteSpace($orderNumber) -and
                [string]::IsNullOrWhiteSpace($sku)) {
                continue
            }

            $scheduleLine = ConvertFrom-WorksheetNumber `
                -Value $values[$rowIndex, $headerIndexes['Sched Line #']] `
                -Context "'Sched Line #' at row $rowIndex in '$WorkbookName'"
            $isZero = [math]::Abs($scheduleLine) -lt 0.000001
            $scheduleIsZero[$rowIndex] = $isZero
            $groupKey = $orderNumber + [char]0x001F + $sku

            if (-not $groups.ContainsKey($groupKey)) {
                $groups[$groupKey] = [pscustomobject]@{
                    HasZero = $false
                    HasNonZero = $false
                }
            }

            if ($isZero) {
                $groups[$groupKey].HasZero = $true
            }
            else {
                $groups[$groupKey].HasNonZero = $true
            }
        }

        $plantValues = New-Object 'object[,]' ($rowCount - 1), 1
        $releaseValues = New-Object 'object[,]' ($rowCount - 1), 1
        $expectedColors = New-Object 'int[]' ($rowCount + 1)
        $previousPlant = $null

        for ($rowIndex = 2; $rowIndex -le $rowCount; $rowIndex++) {
            $outputIndex = $rowIndex - 2
            $plantValue = $values[$rowIndex, $headerIndexes['Plant']]
            if ($null -eq $plantValue -or [string]::IsNullOrWhiteSpace([string]$plantValue)) {
                if ($null -eq $previousPlant) {
                    throw "The first Plant value is blank at row $rowIndex in '$WorkbookName'."
                }
                $plantNumber = $previousPlant
            }
            else {
                $plantNumber = ConvertFrom-WorksheetNumber `
                    -Value $plantValue `
                    -Context "'Plant' at row $rowIndex in '$WorkbookName'"
                $previousPlant = $plantNumber
            }
            $plantValues[$outputIndex, 0] = [double]$plantNumber

            $orderNumber = Get-CellText -Value $values[$rowIndex, $headerIndexes['Order Nbr']]
            $sku = Get-CellText -Value $values[$rowIndex, $headerIndexes['SKU']]
            $groupKey = $orderNumber + [char]0x001F + $sku
            $clearReleaseValue = $false
            if ($groups.ContainsKey($groupKey)) {
                $group = $groups[$groupKey]
                $clearReleaseValue = $scheduleIsZero[$rowIndex] -and
                    $group.HasZero -and
                    $group.HasNonZero
            }

            if ($clearReleaseValue -or
                ([string]::IsNullOrWhiteSpace($orderNumber) -and
                [string]::IsNullOrWhiteSpace($sku))) {
                $releaseValues[$outputIndex, 0] = ''
            }
            else {
                $deliveredQuantity = ConvertFrom-WorksheetNumber `
                    -Value $values[$rowIndex, $headerIndexes['Delivered Qty']] `
                    -Context "'Delivered Qty' at row $rowIndex in '$WorkbookName'"
                $materialAvailableValue = $values[
                    $rowIndex,
                    $headerIndexes['Material Avail.Date']
                ]

                if ([math]::Abs($deliveredQuantity) -ge 0.000001) {
                    $releaseValues[$outputIndex, 0] = 'Delivered'
                }
                elseif ($null -eq $materialAvailableValue -or
                    [string]::IsNullOrWhiteSpace([string]$materialAvailableValue) -or
                    [math]::Abs((ConvertFrom-WorksheetNumber `
                        -Value $materialAvailableValue `
                        -Context "'Material Avail.Date' at row $rowIndex in '$WorkbookName'")) `
                        -lt 0.000001) {
                    $releaseValues[$outputIndex, 0] = 'TBD'
                }
                else {
                    $releaseValues[$outputIndex, 0] =
                        "=BH$rowIndex+MOD(5-WEEKDAY(BH$rowIndex,2),7)"
                }
            }

            if ($scheduleIsZero[$rowIndex]) {
                $shipToCountry = Get-CellText `
                    -Value $values[$rowIndex, $headerIndexes['ShipTo Ctry']]
                if ($shipToCountry -eq 'HK') {
                    $expectedColors[$rowIndex] = $LightGreen
                }
                else {
                    $expectedColors[$rowIndex] = $LightBlue
                }
            }
        }

        $releaseHeader = $Worksheet.Cells.Item(1, 76)
        $releaseHeader.Value2 = 'Date to release'

        $plantStart = $Worksheet.Cells.Item(2, $headerIndexes['Plant'])
        $plantEnd = $Worksheet.Cells.Item($rowCount, $headerIndexes['Plant'])
        $plantRange = $Worksheet.Range($plantStart, $plantEnd)
        $plantRange.NumberFormat = '0'
        $plantRange.Value2 = $plantValues

        $releaseRange = $Worksheet.Range("BX2:BX$rowCount")
        $releaseRange.Formula = $releaseValues
        [void]$releaseRange.Calculate()

        Set-OrAssert-RowColors `
            -Worksheet $Worksheet `
            -ExpectedColors $expectedColors `
            -LastRow $rowCount `
            -WorkbookName $WorkbookName

        $pivotCaches = $Workbook.PivotCaches()
        for ($cacheIndex = 1; $cacheIndex -le $pivotCaches.Count; $cacheIndex++) {
            $pivotCache = $pivotCaches.Item($cacheIndex)
            try {
                $pivotCache.MissingItemsLimit = 0
                [void]$pivotCache.Refresh()
            }
            finally {
                Release-ComObject -ComObject $pivotCache
            }
        }

        return [pscustomobject]@{
            RowCount = $rowCount
            PlantColumnIndex = $headerIndexes['Plant']
            PlantValues = $plantValues
            ReleaseValues = $releaseValues
            ScheduleIsZero = $scheduleIsZero
            ExpectedColors = $expectedColors
        }
    }
    finally {
        foreach ($comObject in @(
            $pivotCaches,
            $releaseHeader,
            $releaseRange,
            $plantRange,
            $plantEnd,
            $plantStart,
            $sourceRange,
            $headerRange
        )) {
            Release-ComObject -ComObject $comObject
        }
    }
}

function Assert-SystemConfirmationContent {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Worksheet,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookName,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Expected
    )

    $releaseHeader = $null
    $plantRange = $null
    $plantStart = $null
    $plantEnd = $null
    $releaseRange = $null

    try {
        $releaseHeader = $Worksheet.Cells.Item(1, 76)
        if ([string]$releaseHeader.Value2 -ne 'Date to release') {
            throw "Workbook '$WorkbookName' does not have the expected BX header."
        }

        $plantStart = $Worksheet.Cells.Item(2, $Expected.PlantColumnIndex)
        $plantEnd = $Worksheet.Cells.Item(
            $Expected.RowCount,
            $Expected.PlantColumnIndex
        )
        $plantRange = $Worksheet.Range($plantStart, $plantEnd)
        $actualPlants = $plantRange.Value2
        $releaseRange = $Worksheet.Range("BX2:BX$($Expected.RowCount)")
        $actualReleaseValues = $releaseRange.Formula

        for ($rowIndex = 2; $rowIndex -le $Expected.RowCount; $rowIndex++) {
            $arrayIndex = $rowIndex - 1
            $actualPlant = $actualPlants[$arrayIndex, 1]
            if ($actualPlant -is [string]) {
                throw "Plant at row $rowIndex in '$WorkbookName' is still stored as text."
            }
            Assert-NearlyEqual `
                -Actual ([double]$actualPlant) `
                -Expected ([double]$Expected.PlantValues[($rowIndex - 2), 0]) `
                -Message "Plant at row $rowIndex in '$WorkbookName' is incorrect."

            $expectedReleaseValue = [string]$Expected.ReleaseValues[($rowIndex - 2), 0]
            $actualReleaseValue = [string]$actualReleaseValues[$arrayIndex, 1]
            if ($actualReleaseValue -ne $expectedReleaseValue) {
                throw "Date to release at row $rowIndex in '$WorkbookName' is incorrect."
            }

        }

        Set-OrAssert-RowColors `
            -Worksheet $Worksheet `
            -ExpectedColors $Expected.ExpectedColors `
            -LastRow $Expected.RowCount `
            -WorkbookName $WorkbookName `
            -Validate
    }
    finally {
        foreach ($comObject in @(
            $releaseRange,
            $plantRange,
            $plantEnd,
            $plantStart,
            $releaseHeader
        )) {
            Release-ComObject -ComObject $comObject
        }
    }
}

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceDirectory = Split-Path -Parent $resolvedWorkbookPath
$cmmrSourcePath = Join-Path $sourceDirectory 'CMMR_INV_RAW_DATA.XLS'
$retailSourcePath = Join-Path $sourceDirectory 'RETAIL_INV_RAW_DATA.xls'
$cmmrOorPath = Join-Path $sourceDirectory 'CMMR allregion OOR.xlsx'
$rtOorPath = Join-Path $sourceDirectory 'RT ALLRG OOR.xlsx'

foreach ($sourcePath in @(
    $cmmrSourcePath,
    $retailSourcePath,
    $cmmrOorPath,
    $rtOorPath
)) {
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Source file not found: $sourcePath"
    }
}

$cmmrRows = Read-InventorySource -SourcePath $cmmrSourcePath
$retailRows = Read-InventorySource -SourcePath $retailSourcePath
$cmmrValues = ConvertTo-WorksheetArray -Rows $cmmrRows -Headers $outputHeaders
$retailValues = ConvertTo-WorksheetArray -Rows $retailRows -Headers $outputHeaders

$lightGreen = Get-OleColor -Red 198 -Green 239 -Blue 206
$lightBlue = Get-OleColor -Red 221 -Green 235 -Blue 247
$excel = $null
$workbook = $null
$cmmrWorksheet = $null
$retailWorksheet = $null
$cmmrOorWorkbook = $null
$cmmrOorWorksheet = $null
$rtOorWorkbook = $null
$rtOorWorksheet = $null
$excelComFilterRegistered = $false
$excelProcessId = 0
$excelProcess = $null

try {
    Register-ExcelComRetryFilter
    $excelComFilterRegistered = $true

    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = [ExcelComRetryFilter]::GetProcessId(
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

    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath, 0, $false)
    if ([bool]$workbook.ReadOnly) {
        throw "Workbook opened as read-only and cannot be updated: $resolvedWorkbookPath"
    }
    $cmmrOorWorkbook = $excel.Workbooks.Open($cmmrOorPath, 0, $false)
    if ([bool]$cmmrOorWorkbook.ReadOnly) {
        throw "Workbook opened as read-only and cannot be updated: $cmmrOorPath"
    }
    $rtOorWorkbook = $excel.Workbooks.Open($rtOorPath, 0, $false)
    if ([bool]$rtOorWorkbook.ReadOnly) {
        throw "Workbook opened as read-only and cannot be updated: $rtOorPath"
    }
    $cmmrOorWorksheet = $cmmrOorWorkbook.Worksheets.Item(1)
    $rtOorWorksheet = $rtOorWorkbook.Worksheets.Item(1)
    $excel.Calculation = -4135

    Remove-Worksheet -Workbook $workbook -WorksheetName '__CMMR_INV_TEMP__'
    Remove-Worksheet -Workbook $workbook -WorksheetName '__RETAIL_INV_TEMP__'

    $temporarySuffix = [guid]::NewGuid().ToString('N').Substring(0, 8)
    $cmmrWorksheet = $workbook.Worksheets.Add()
    $cmmrWorksheet.Name = "__CMMR_$temporarySuffix"
    $retailWorksheet = $workbook.Worksheets.Add()
    $retailWorksheet.Name = "__RETAIL_$temporarySuffix"

    Remove-Worksheet -Workbook $workbook -WorksheetName 'CMMR INV'
    Remove-Worksheet -Workbook $workbook -WorksheetName 'Retail INV'

    $cmmrWorksheet.Name = 'CMMR INV'
    $retailWorksheet.Name = 'Retail INV'

    Set-WorksheetContent `
        -Application $excel `
        -Worksheet $cmmrWorksheet `
        -Values $cmmrValues `
        -HeaderColor $lightGreen
    Set-WorksheetContent `
        -Application $excel `
        -Worksheet $retailWorksheet `
        -Values $retailValues `
        -HeaderColor $lightBlue

    $cmmrOorExpected = Set-SystemConfirmationContent `
        -Workbook $cmmrOorWorkbook `
        -Worksheet $cmmrOorWorksheet `
        -WorkbookName 'CMMR allregion OOR.xlsx' `
        -LightGreen $lightGreen `
        -LightBlue $lightBlue
    $rtOorExpected = Set-SystemConfirmationContent `
        -Workbook $rtOorWorkbook `
        -Worksheet $rtOorWorksheet `
        -WorkbookName 'RT ALLRG OOR.xlsx' `
        -LightGreen $lightGreen `
        -LightBlue $lightBlue

    Save-WorkbookWithRetry `
        -Workbook $workbook `
        -WorkbookName (Split-Path -Leaf $resolvedWorkbookPath)
    Save-WorkbookWithRetry `
        -Workbook $cmmrOorWorkbook `
        -WorkbookName (Split-Path -Leaf $cmmrOorPath)
    Save-WorkbookWithRetry `
        -Workbook $rtOorWorkbook `
        -WorkbookName (Split-Path -Leaf $rtOorPath)

    Assert-WorksheetNameCount -Workbook $workbook -WorksheetName 'CMMR INV'
    Assert-WorksheetNameCount -Workbook $workbook -WorksheetName 'Retail INV'
    Assert-Worksheet `
        -Application $excel `
        -Worksheet $cmmrWorksheet `
        -DataRowCount $cmmrRows.Count `
        -ExpectedHeaderColor $lightGreen
    Assert-Worksheet `
        -Application $excel `
        -Worksheet $retailWorksheet `
        -DataRowCount $retailRows.Count `
        -ExpectedHeaderColor $lightBlue
    Assert-SystemConfirmationContent `
        -Worksheet $cmmrOorWorksheet `
        -WorkbookName 'CMMR allregion OOR.xlsx' `
        -Expected $cmmrOorExpected
    Assert-SystemConfirmationContent `
        -Worksheet $rtOorWorksheet `
        -WorkbookName 'RT ALLRG OOR.xlsx' `
        -Expected $rtOorExpected

    Write-Host 'CMMR-RT inventory and system confirmation update succeeded.'
}
finally {
    $shutdownWatchdog = $null
    if ($null -ne $excelProcess -and -not $excelProcess.HasExited) {
        $shutdownWatchdog = [ExcelShutdownWatchdog]::Start(
            $excelProcess,
            15000
        )
    }

    if ($null -ne $rtOorWorkbook) {
        try {
            $rtOorWorkbook.Close($false)
        }
        catch {
            Write-Warning "Failed to close RT ALLRG OOR.xlsx: $($_.Exception.Message)"
        }
    }
    if ($null -ne $cmmrOorWorkbook) {
        try {
            $cmmrOorWorkbook.Close($false)
        }
        catch {
            Write-Warning "Failed to close CMMR allregion OOR.xlsx: $($_.Exception.Message)"
        }
    }
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
            Write-Warning "Failed to close the inventory workbook: $($_.Exception.Message)"
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Warning "Failed to quit Excel cleanly: $($_.Exception.Message)"
        }
    }

    foreach ($comObject in @(
        $rtOorWorksheet,
        $rtOorWorkbook,
        $cmmrOorWorksheet,
        $cmmrOorWorkbook,
        $retailWorksheet,
        $cmmrWorksheet,
        $workbook,
        $excel
    )) {
        Release-ComObject -ComObject $comObject
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($excelComFilterRegistered) {
        [ExcelComRetryFilter]::Revoke()
    }

    if ($null -ne $excelProcess) {
        try {
            if (-not $excelProcess.HasExited -and
                -not $excelProcess.WaitForExit(5000)) {
                $excelProcess.Kill()
                $excelProcess.WaitForExit()
                Write-Warning "Terminated the unresponsive Excel process $excelProcessId."
            }
        }
        catch [System.InvalidOperationException] {
        }
        finally {
            if ($null -ne $shutdownWatchdog) {
                if ($shutdownWatchdog.Terminated) {
                    Write-Warning "The Excel shutdown watchdog terminated process $excelProcessId."
                }
                $shutdownWatchdog.Dispose()
            }
            $excelProcess.Dispose()
        }
    }
}
