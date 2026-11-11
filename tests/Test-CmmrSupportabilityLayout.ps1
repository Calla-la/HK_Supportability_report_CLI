$ErrorActionPreference = 'Stop'

class CmmrLayoutTestSheet {
    [hashtable]$Ranges = @{}

    [object] Range([string]$address) {
        if (-not $this.Ranges.ContainsKey($address)) {
            $this.Ranges[$address] = [pscustomobject]@{
                Address = $address
                Value2 = $null
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
        $stockWeeklyOutput = [object[,]]::new(2, 2)
        $monthlyOutput[0, 0] = 10
        $monthlyOutput[0, 1] = 11
        $monthlyOutput[0, 2] = 12
        $openOrderOutput[0, 0] = 100
        $openOrderOutput[0, 1] = 200
        $openOrderOutput[0, 2] = 300
        $openOrderOutput[0, 3] = 400
        $stockWeeklyOutput[0, 0] = 500
        $stockWeeklyOutput[0, 1] = 600

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
        if (($writes -join ',') -ne ($expectedAddresses -join ',')) {
            throw "Incorrect destination writes: $($writes -join ',')"
        }
        $expectedValidation = $expectedLabels | ForEach-Object {
            "$DestinationSheetName $_"
        }
        if (($validated -join ',') -ne ($expectedValidation -join ',')) {
            throw "Incorrect validation labels: $($validated -join ',')"
        }
        if ($targetSheet.Ranges.Count -ne 3) {
            throw 'Unexpected destination range access.'
        }
    }
}

Write-Host 'CMMR unified/legacy layout and confirmation-only isolation tests passed.'
