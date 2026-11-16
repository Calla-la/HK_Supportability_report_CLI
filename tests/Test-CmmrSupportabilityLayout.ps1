$ErrorActionPreference = 'Stop'

class CmmrLayoutTestSheet {
    [hashtable]$Ranges = @{}

    [object] Range([string]$address) {
        if (-not $this.Ranges.ContainsKey($address)) {
            $this.Ranges[$address] = [pscustomobject]@{
                Address = $address
                Value2 = $null
                NumberFormat = '#,##0'
            }
        }
        return $this.Ranges[$address]
    }
}

$scriptPath = Join-Path $PSScriptRoot '..\scripts\Update-CmmrSupportability.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    $scriptPath, [ref]$tokens, [ref]$parseErrors
)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }

$layoutAssignments = @($ast.EndBlock.Statements | Where-Object {
    $_ -is [Management.Automation.Language.AssignmentStatementAst] -and
    $_.Left.VariablePath.UserPath -in @(
        'monthlyColumns', 'openOrderColumns', 'stockWeeklyColumns'
    )
})
$refreshTry = $ast.EndBlock.Statements | Where-Object {
    $_ -is [Management.Automation.Language.TryStatementAst]
} | Select-Object -First 1
$rangeStatements = @($refreshTry.Body.Statements | Where-Object {
    $_ -is [Management.Automation.Language.IfStatementAst] -and
    $_.Extent.Text -match '^if \(-not \$ConfirmationOnly\)' -and
    $_.Extent.Text -match '\$(monthlyTarget|stockWeeklyTarget)'
})
if ($layoutAssignments.Count -ne 3 -or $rangeStatements.Count -ne 5) {
    throw 'Could not locate the CMMR layout, write, and validation blocks.'
}
$configureLayout = [scriptblock]::Create(
    (($layoutAssignments | ForEach-Object { $_.Extent.Text }) -join "`n")
)
$writeAndValidate = [scriptblock]::Create(
    (($rangeStatements | ForEach-Object { $_.Extent.Text }) -join "`n")
)

function Set-RangeValues($Worksheet, $Range, $Values, $RowRuns) {
    $Range.Value2 = $Values
    $writes.Add($Range.Address)
}

function Assert-ArrayValues($Actual, $Expected, $RangeName) {
    if (-not [object]::ReferenceEquals($Actual, $Expected)) {
        throw "Unexpected payload for $RangeName"
    }
    $validated.Add($RangeName)
}

foreach ($mode in @('Full', 'MonthlyOnly', 'OpenOrdersOnly', 'InventoryOnly', 'WeeklyOnly')) {
$MonthlyOnly = $mode -eq 'MonthlyOnly'
$OpenOrdersOnly = $mode -eq 'OpenOrdersOnly'
$InventoryOnly = $mode -eq 'InventoryOnly'
$WeeklyOnly = $mode -eq 'WeeklyOnly'
foreach ($unified in @($true, $false)) {
    foreach ($confirmationOnly in @($false, $true)) {
        $sameWorkbook = $unified
        $DestinationSheetName = if ($unified) {
            'CMMR Supportability'
        } else { 'Supportability' }
        $targetStartRow = 4
        $targetLastRow = 5
        $targetSheet = [CmmrLayoutTestSheet]::new()
        $writes = [Collections.Generic.List[string]]::new()
        $validated = [Collections.Generic.List[string]]::new()
        $rowRuns = @()
        $monthlyOutput = [object[,]]::new(2, 3)
        $openOrderOutput = [object[,]]::new(2, 4)
        $stockWeeklyOutput = [object[,]]::new(
            2, $(if ($InventoryOnly -or $WeeklyOnly) { 1 } else { 2 })
        )
        $monthlyOutput[0, 0] = 10
        $monthlyOutput[0, 1] = 11
        $monthlyOutput[0, 2] = 12
        $openOrderOutput[0, 0] = 100
        $openOrderOutput[0, 1] = 200
        $openOrderOutput[0, 2] = 300
        $openOrderOutput[0, 3] = 400
        $stockWeeklyOutput[0, 0] = 500
        if (-not $InventoryOnly -and -not $WeeklyOnly) {
            $stockWeeklyOutput[0, 1] = 600
        }

        . $configureLayout
        . $writeAndValidate

        if ($confirmationOnly) {
            if ($writes.Count -or $validated.Count -or $targetSheet.Ranges.Count) {
                throw 'Confirmation-only mode accessed non-confirmation ranges.'
            }
            continue
        }
        $expectedAddresses = if ($unified) {
            @('N4:P5', 'R4:U5', 'AK4:AL5')
        } else { @('O4:Q5', 'S4:V5', 'AM4:AN5') }
        $expectedLabels = if ($unified) {
            @('N:P', 'R:U', 'AK:AL')
        } else { @('O:Q', 'S:V', 'AM:AN') }
        if ($MonthlyOnly) {
            $expectedAddresses = @($expectedAddresses[0])
            $expectedLabels = @($expectedLabels[0])
        } elseif ($OpenOrdersOnly) {
            $expectedAddresses = @($expectedAddresses[1])
            $expectedLabels = @($expectedLabels[1])
        } elseif ($InventoryOnly) {
            $expectedAddresses = if ($unified) { @('AK4:AK5') } else { @('AM4:AM5') }
            $expectedLabels = if ($unified) { @('AK:AK') } else { @('AM:AM') }
        } elseif ($WeeklyOnly) {
            $expectedAddresses = if ($unified) { @('AL4:AL5') } else { @('AN4:AN5') }
            $expectedLabels = if ($unified) { @('AL:AL') } else { @('AN:AN') }
        }
        if (($writes -join ',') -ne ($expectedAddresses -join ',')) {
            throw "Incorrect destination writes: $($writes -join ',')"
        }
        $expectedValidation = $expectedLabels | ForEach-Object {
            "$DestinationSheetName $_"
        }
        if (($validated -join ',') -ne ($expectedValidation -join ',')) {
            throw "Incorrect validation labels: $($validated -join ',')"
        }
        if ($targetSheet.Ranges.Count -ne $expectedAddresses.Count) {
            throw 'Unexpected destination range access.'
        }
        $expectedMonthlyFormat = if ($unified) {
            '#,##0;-#,##0;;@'
        } else { '#,##0' }
        if (-not $OpenOrdersOnly -and -not $InventoryOnly -and -not $WeeklyOnly -and
            $monthlyTarget.NumberFormat -ne $expectedMonthlyFormat) {
            throw 'Monthly zero-display formatting changed the wrong destination.'
        }
    }
    }
}

foreach ($definition in $ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -in @(
        'Get-NormalizedKey', 'Get-NormalizedSoldTo', 'ConvertTo-Number',
        'Add-ToMap', 'Get-DisplayedMonth', 'Get-MonthlyActuals', 'Get-MapValue'
    )
}, $false)) {
    Invoke-Expression $definition.Extent.Text
}
$monthlyLoop = $ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.ForStatementAst] -and
    $node.Extent.Text -match '^for \(\$monthIndex ='
}, $false) | Select-Object -First 1
if ($null -eq $monthlyLoop) { throw 'Could not locate monthly output mapping.' }
$mapMonthlyOutput = [scriptblock]::Create($monthlyLoop.Extent.Text)
$soldToParties = @('5395800', '5395801', '5395799')
foreach ($monthCount in @(1, 2, 3)) {
    $values = [Array]::CreateInstance(
        [object], [int[]]@(12, 29), [int[]]@(1, 1)
    )
    $values[7, 3] = '<7/2/2026'
    $values[8, 3] = '5395800'
    $values[9, 3] = 9999
    $values[9, 2] = 'sku-a'
    $values[10, 2] = 'SKU-A'
    $values[11, 2] = 'SKU-B'
    $values[12, 2] = 'Grand Total'
    for ($m = 0; $m -lt $monthCount; $m++) {
        $startColumn = 4 + $m * 5
        $values[7, $startColumn] = @('Oct', 'Nov', 'Dec')[$m]
        for ($s = 0; $s -lt 5; $s++) {
            $column = $startColumn + $s
            $values[8, $column] = @(
                '05395800', '5395801', '5395799', '5395802', 'Total'
            )[$s]
            $values[9, $column] = ($m + 1) * ($s + 1)
            $values[10, $column] = 10
            $values[11, $column] = 20
            $values[12, $column] = 9999
        }
    }
    $values[7, (4 + $monthCount * 5)] = 'Grand Total'
    $monthly = Get-MonthlyActuals -Values $values -LastRow 12
    if ($monthly.MonthOrder.Count -ne $monthCount) {
        throw 'Monthly aggregation included a non-month pivot item.'
    }
    foreach ($productId in @('SKU-A', 'SKU-B', 'MISSING-SKU')) {
        $outputRow = 0
        $monthlyOutput = [object[,]]::new(1, 3)
        . $mapMonthlyOutput
        for ($m = 0; $m -lt 3; $m++) {
            $expected = if ($m -ge $monthCount -or $productId -eq 'MISSING-SKU') {
                0
            } elseif ($productId -eq 'SKU-A') {
                6 * ($m + 1) + 30
            } else { 60 }
            if ($monthlyOutput[0, $m] -isnot [double] -or
                $monthlyOutput[0, $m] -ne $expected) {
                throw "Incorrect monthly quantity for $productId, month $m."
            }
        }
    }
}

$retryDefinition = $ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
    $node.Name -eq 'Invoke-ExcelComRetry'
}, $false) | Select-Object -First 1
if ($null -eq $retryDefinition) { throw 'Could not locate COM retry helper.' }
Invoke-Expression $retryDefinition.Extent.Text

foreach ($wrapped in @($false, $true)) {
    $state = @{ Attempts = 0 }
    $result = Invoke-ExcelComRetry -Context 'restore Automatic calculation' -Operation {
        $state.Attempts++
        if ($state.Attempts -eq 1) {
            $busy = [Runtime.InteropServices.COMException]::new(
                'Excel is busy', -2146777998
            )
            if ($wrapped) {
                throw [Management.Automation.SetValueInvocationException]::new(
                    'Exception setting Calculation', $busy
                )
            }
            throw $busy
        }
        return 'ready'
    }
    if ($state.Attempts -ne 2 -or $result -ne 'ready') {
        throw 'Transient Excel busy error was not retried successfully.'
    }
}

foreach ($failure in @(
    [Runtime.InteropServices.COMException]::new('Access denied', -2147024891),
    [InvalidOperationException]::new('Invalid operation')
)) {
    $state = @{ Attempts = 0 }
    $caught = $false
    try {
        Invoke-ExcelComRetry -Context 'non-transient failure' -Operation {
            $state.Attempts++
            throw $failure
        }
    }
    catch {
        $caught = $true
        if ($_.Exception.Message -notlike "*$($failure.Message)*") { throw }
    }
    if (-not $caught -or $state.Attempts -ne 1) {
        throw 'Non-transient error was swallowed or retried.'
    }
}

$state = @{ Attempts = 0 }
$caught = $false
try {
    Invoke-ExcelComRetry -Context 'busy timeout' -TimeoutSeconds 1 -Operation {
        $state.Attempts++
        throw [Runtime.InteropServices.COMException]::new('Excel is busy', -2146777998)
    }
}
catch {
    $caught = $true
    if ($_.Exception.Message -notlike "*busy timeout*0x800AC472*") { throw }
}
if (-not $caught -or $state.Attempts -gt 4) {
    throw 'Persistent Excel busy error did not stop at the retry deadline.'
}

Write-Host 'CMMR layout, monthly aggregation, and bounded COM retry tests passed.'
