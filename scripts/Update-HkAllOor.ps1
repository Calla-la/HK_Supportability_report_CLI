[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetWorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$ReplacementWorkbookPath
)

$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([object]$ComObject)

    if ($null -ne $ComObject -and
        [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(
            $ComObject
        )
    }
}

function Register-HkAllOorExcelComFilter {
    if (-not ('HkAllOorOleMessageFilter' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IHkAllOorOleMessageFilter
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

public sealed class HkAllOorOleMessageFilter : IHkAllOorOleMessageFilter
{
    private static HkAllOorOleMessageFilter current;
    private static IHkAllOorOleMessageFilter previous;

    [DllImport("Ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        IHkAllOorOleMessageFilter newFilter,
        out IHkAllOorOleMessageFilter oldFilter);

    public static void Register()
    {
        if (current != null)
        {
            return;
        }

        current = new HkAllOorOleMessageFilter();
        IHkAllOorOleMessageFilter oldFilter;
        CoRegisterMessageFilter(current, out oldFilter);
        previous = oldFilter;
    }

    public static void Revoke()
    {
        IHkAllOorOleMessageFilter ignored;
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

    [HkAllOorOleMessageFilter]::Register()
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
    while ($true) {
        try {
            return & $Operation
        }
        catch [Runtime.InteropServices.COMException] {
            $hresult = $_.Exception.HResult -band 0xFFFFFFFFL
            if ($hresult -notin @(
                0x80010001L,
                0x8001010AL,
                0x800AC472L
            )) {
                throw
            }
            if ([datetime]::UtcNow.AddMilliseconds(250) -ge $deadline) {
                $hresultText = '0x{0:X8}' -f $hresult
                throw "Excel COM operation '$Context' failed after $TimeoutSeconds seconds ($hresultText): $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds 250
        }
    }
}

foreach ($path in @($TargetWorkbookPath, $ReplacementWorkbookPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Workbook not found: $path"
    }
}

$targetPath = (Resolve-Path -LiteralPath $TargetWorkbookPath).Path
$replacementPath = (Resolve-Path -LiteralPath $ReplacementWorkbookPath).Path
if ($targetPath -eq $replacementPath) {
    throw 'Target and replacement workbook paths must be different.'
}

$excel = $null
$targetWorkbook = $null
$replacementWorkbook = $null
$targetSheet = $null
$replacementSheet = $null
$replacementSearchRange = $null
$replacementLastCell = $null
$replacementRange = $null
$targetClearRange = $null
$targetPasteRange = $null
$targetExistingRange = $null
$targetLastCell = $null
$filterRange = $null
$hkOorPivotSheet = $null
$pivotTables = $null
$deliveryDatePivot = $null
$scheduleLineField = $null
$soldToField = $null
$scheduleLineCell = $null
$soldToCell = $null
$comFilterRegistered = $false

try {
    Register-HkAllOorExcelComFilter
    $comFilterRegistered = $true

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    $excel.ScreenUpdating = $false

    $replacementWorkbook = $excel.Workbooks.Open(
        $replacementPath,
        0,
        $true
    )
    $targetWorkbook = $excel.Workbooks.Open($targetPath, 0, $false)
    if ([bool]$targetWorkbook.ReadOnly) {
        throw "Target workbook opened as read-only: $targetPath"
    }

    $replacementSheet = $replacementWorkbook.Worksheets.Item(1)
    $targetSheet = $targetWorkbook.Worksheets.Item('HK all OOR')

    $replacementSearchRange = $replacementSheet.Range('A:BW')
    $replacementLastCell = $replacementSearchRange.Find(
        '*',
        $replacementSearchRange.Cells.Item(1, 1),
        -4163,
        2,
        1,
        2,
        $false,
        $false,
        $false
    )
    if ($null -eq $replacementLastCell) {
        throw "Replacement worksheet has no data: $replacementPath"
    }

    $replacementLastRow = [int]$replacementLastCell.Row
    $replacementRange = $replacementSheet.Range(
        "A1:BW$replacementLastRow"
    )
    $replacementValues = $replacementRange.Value2

    $targetExistingRange = $targetSheet.Range('A:D')
    $targetLastCell = $targetExistingRange.Find(
        '*',
        $targetExistingRange.Cells.Item(1, 1),
        -4163,
        2,
        1,
        2,
        $false,
        $false,
        $false
    )
    $targetLastRow = if ($null -eq $targetLastCell) {
        1
    }
    else {
        [int]$targetLastCell.Row
    }

    $targetClearRange = $targetSheet.Range('E:CA')
    $targetClearRange.ClearContents()

    $targetPasteRange = $targetSheet.Range(
        "E1:CA$replacementLastRow"
    )
    $targetPasteRange.Value2 = $replacementValues

    if ([bool]$targetSheet.AutoFilterMode) {
        $targetSheet.AutoFilterMode = $false
    }
    $filterLastRow = [math]::Max($targetLastRow, $replacementLastRow)
    $filterRange = $targetSheet.Range("A1:CA$filterLastRow")
    [void]$filterRange.AutoFilter()

    $hkOorPivotSheet = $targetWorkbook.Worksheets.Item('HK OOR Pivot')
    $pivotTables = $hkOorPivotSheet.PivotTables()
    for ($pivotIndex = 1; $pivotIndex -le $pivotTables.Count; $pivotIndex++) {
        $pivotTable = $pivotTables.Item($pivotIndex)
        try {
            $pivotRefreshed = Invoke-ExcelComRetry `
                -Context "refresh PivotTable '$($pivotTable.Name)'" `
                -Operation {
                    [bool]$pivotTable.RefreshTable()
                }
            if (-not $pivotRefreshed) {
                throw "Failed to refresh PivotTable '$($pivotTable.Name)'."
            }
        }
        finally {
            Release-ComObject -ComObject $pivotTable
        }
    }

    $deliveryDatePivot = $pivotTables.Item('DeliveryDateDetail')
    $scheduleLineField = $deliveryDatePivot.PivotFields('Sched Line #')
    $soldToField = $deliveryDatePivot.PivotFields('SoldTo Nbr')

    $scheduleLineCell = $hkOorPivotSheet.Range('Z3')
    $soldToCell = $hkOorPivotSheet.Range('Z4')
    $scheduleLineValue = Invoke-ExcelComRetry `
        -Context 'read HK OOR Pivot Z3' `
        -Operation { [string]$scheduleLineCell.Value2 }
    $soldToValue = Invoke-ExcelComRetry `
        -Context 'read HK OOR Pivot Z4' `
        -Operation { [string]$soldToCell.Value2 }

    if ($scheduleLineValue -ne '0') {
        Invoke-ExcelComRetry -Context 'restore Sched Line # filter' -Operation {
            $scheduleLineField.EnableMultiplePageItems = $false
            $scheduleLineField.CurrentPage = '0'
        }
        $scheduleLineValue = Invoke-ExcelComRetry `
            -Context 'validate HK OOR Pivot Z3' `
            -Operation { [string]$scheduleLineCell.Value2 }
    }
    if ($soldToValue -ne '5395802') {
        Invoke-ExcelComRetry -Context 'restore SoldTo Nbr filter' -Operation {
            $soldToField.EnableMultiplePageItems = $false
            $soldToField.CurrentPage = '5395802'
        }
        $soldToValue = Invoke-ExcelComRetry `
            -Context 'validate HK OOR Pivot Z4' `
            -Operation { [string]$soldToCell.Value2 }
    }

    if ($scheduleLineValue -ne '0') {
        throw "Expected HK OOR Pivot Z3 to be '0' after refresh."
    }
    if ($soldToValue -ne '5395802') {
        throw "Expected HK OOR Pivot Z4 to be '5395802' after refresh."
    }

    Invoke-ExcelComRetry -Context 'save source workbook' -Operation {
        $targetWorkbook.Save()
    }
}
finally {
    if ($null -ne $targetWorkbook) {
        try {
            Invoke-ExcelComRetry -Context 'close source workbook' -Operation {
                $targetWorkbook.Close($false)
            }
        }
        catch {
            Write-Warning "Failed to close target workbook: $($_.Exception.Message)"
        }
    }
    if ($null -ne $replacementWorkbook) {
        try {
            Invoke-ExcelComRetry `
                -Context 'close replacement workbook' `
                -Operation {
                    $replacementWorkbook.Close($false)
                }
        }
        catch {
            Write-Warning "Failed to close replacement workbook: $($_.Exception.Message)"
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
        $soldToCell,
        $scheduleLineCell,
        $soldToField,
        $scheduleLineField,
        $deliveryDatePivot,
        $pivotTables,
        $hkOorPivotSheet,
        $filterRange,
        $targetLastCell,
        $targetExistingRange,
        $targetPasteRange,
        $targetClearRange,
        $replacementRange,
        $replacementLastCell,
        $replacementSearchRange,
        $targetSheet,
        $replacementSheet,
        $targetWorkbook,
        $replacementWorkbook,
        $excel
    )) {
        Release-ComObject -ComObject $comObject
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($comFilterRegistered) {
        [HkAllOorOleMessageFilter]::Revoke()
    }
}

Write-Host 'HK all OOR source data refresh succeeded.'
