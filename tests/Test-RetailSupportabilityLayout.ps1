$ErrorActionPreference = 'Stop'

class RetailLayoutTestRange {
    [string]$Address
    [object]$Value2
    [object]$Formula2
    [bool]$MergeCells = $false
    [string]$NumberFormat = ''

    RetailLayoutTestRange([string]$address) { $this.Address = $address }
    [void] ClearContents() { $this.Value2 = $null }
}

class RetailLayoutTestCells {
    [hashtable]$Items = @{}

    [RetailLayoutTestRange] Item([int]$row, [int]$column) {
        $key = "$row,$column"
        if (-not $this.Items.ContainsKey($key)) {
            $this.Items[$key] = [RetailLayoutTestRange]::new($key)
        }
        return $this.Items[$key]
    }
}

class RetailLayoutTestSheet {
    [string]$Name
    [int]$ProductColumn
    [hashtable]$Ranges = @{}
    [RetailLayoutTestCells]$Cells = [RetailLayoutTestCells]::new()
    [Collections.Generic.List[string]]$Writes =
        [Collections.Generic.List[string]]::new()

    [RetailLayoutTestRange] Range([string]$address) {
        if (-not $this.Ranges.ContainsKey($address)) {
            $range = [RetailLayoutTestRange]::new($address)
            $range.Value2 = [object[,]]::new(3, 105)
            $this.Ranges[$address] = $range
        }
        return $this.Ranges[$address]
    }
}

class RetailLayoutTestSheets {
    [hashtable]$Items = @{}

    [RetailLayoutTestSheet] Item([string]$name) {
        if (-not $this.Items.ContainsKey($name)) {
            throw "Unexpected worksheet access: $name"
        }
        return $this.Items[$name]
    }
}

class RetailLayoutTestWorkbook {
    [RetailLayoutTestSheets]$Worksheets = [RetailLayoutTestSheets]::new()
    [int]$SaveCount

    [void] Save() { $this.SaveCount++ }
}

$scriptPath = Join-Path $PSScriptRoot '..\scripts\Update-RetailSupportability.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile(
    $scriptPath, [ref]$tokens, [ref]$parseErrors
)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
foreach ($definition in $ast.FindAll({
    param($node)
    $node -is [Management.Automation.Language.FunctionDefinitionAst]
}, $false)) {
    Invoke-Expression $definition.Extent.Text
}
$layoutAssignments = $ast.EndBlock.Statements | Where-Object {
    $_ -is [Management.Automation.Language.AssignmentStatementAst] -and
    $_.Left.VariablePath.UserPath -in @(
        'targetLayout', 'monthlyColumns', 'openOrderColumns'
    )
}
$configureLayout = [scriptblock]::Create(
    (($layoutAssignments | ForEach-Object { $_.Extent.Text }) -join "`n")
)
$refreshTry = $ast.EndBlock.Statements | Where-Object {
    $_ -is [Management.Automation.Language.TryStatementAst]
} | Select-Object -First 1
$firstReportStatement = $refreshTry.Body.Statements | Where-Object {
    $_.Extent.Text -match '^\$oorSheet\s*='
} | Select-Object -First 1
if ($layoutAssignments.Count -ne 3 -or $null -eq $firstReportStatement) {
    throw 'The Retail refresh test could not locate the report execution block.'
}
$refreshBody = [scriptblock]::Create(
    (($refreshTry.Body.Statements | Where-Object {
        $_.Extent.StartOffset -ge $firstReportStatement.Extent.StartOffset
    } | ForEach-Object { $_.Extent.Text }) -join "`n")
)

function New-TestExcelArray([int]$Rows, [int]$Columns) {
    return ,([Array]::CreateInstance(
        [object], [int[]]@($Rows, $Columns), [int[]]@(1, 1)
    ))
}

function Get-LastDataRow($Worksheet, [int]$Column) {
    if ($Worksheet.Name -eq 'Destination') {
        if ($Column -ne $Worksheet.ProductColumn) {
            throw "Product matching used the wrong column: $Column"
        }
        return 5
    }
    if ($Worksheet.Name -eq 'Retail INV') { return 4 }
    return 2
}

function New-TestStockSheet {
    $sheet = [RetailLayoutTestSheet]::new()
    $sheet.Name = 'Retail INV'
    $sheet.Cells.Item(1, 1).Value2 = 'Material'
    $values = New-TestExcelArray 4 10
    $values[1, 1] = 'Material'
    $values[2, 1] = 'sku-a'
    $values[2, 10] = 31
    $values[3, 1] = 'SKU-A'
    $values[3, 10] = 99
    $values[4, 1] = 'SKU-OTHER'
    $values[4, 10] = 17
    $sheet.Range('A1:J4').Value2 = $values
    return $sheet
}

function Get-OpenOrderTotal($Values, $LastRow) {
    return @{'SKU-A' = 5}
}

function Get-WeeklyOpenOrders($Values, $LastRow, $RunDate) {
    return [pscustomobject]@{
        Current = @{'SKU-A' = 7}
        Future = @([pscustomobject]@{
            Friday = [datetime]'2026-11-06'
            Values = @{'SKU-A' = 11}
        })
    }
}

function Get-MonthlyActuals($Values, $LastRow) {
    return [pscustomobject]@{
        Months = @('Oct', 'Nov')
        Maps = @{Oct = @{'SKU-A' = 17}; Nov = @{'SKU-A' = 19}}
    }
}

function Get-WeeklyActuals($Values, $LastRow, $RunDate) {
    return [pscustomobject]@{
        Values = @{'SKU-A' = 23}
        StartDate = [datetime]'2026-10-19'
        EndDate = [datetime]'2026-10-25'
    }
}

function Get-RowVisibilityRuns($Worksheet, $StartRow, $EndRow) {
    return [pscustomobject]@{StartRow = $StartRow; EndRow = $EndRow}
}

function Set-RangeValues($Worksheet, $Range, $Values, $RowRuns, $Context) {
    $output = New-TestExcelArray $Values.GetLength(0) $Values.GetLength(1)
    for ($r = 0; $r -lt $Values.GetLength(0); $r++) {
        for ($c = 0; $c -lt $Values.GetLength(1); $c++) {
            $output[($r + 1), ($c + 1)] = $Values[$r, $c]
        }
    }
    $Range.Value2 = $output
    $Worksheet.Writes.Add($Range.Address)
}

function Test-RetailLayout([bool]$Unified, [bool]$InvalidHeader = $false) {
    $sameWorkbook = $Unified
    $DestinationSheetName = if ($Unified) {
        'Retail supportability'
    }
    else { 'Supportability' }
    $targetStartRow = 4
    $AsOfDate = [datetime]'2026-10-29'
    $excel = [pscustomobject]@{Calculation = -4135}
    $sourceWorkbook = [RetailLayoutTestWorkbook]::new()
    $destinationWorkbook = if ($Unified) { $sourceWorkbook }
        else { [RetailLayoutTestWorkbook]::new() }
    $destination = [RetailLayoutTestSheet]::new()
    $destination.Name = 'Destination'
    $destination.ProductColumn = if ($Unified) { 14 } else { 16 }
    $productColumn = if ($Unified) { 'N' } else { 'P' }
    $productHeader = $destination.Cells.Item(3, $destination.ProductColumn)
    $productHeader.Value2 = if ($InvalidHeader) { 'Wrong layout' }
        else { 'Product ID' }
    $products = New-TestExcelArray 2 1
    $products[1, 1] = 'SKU-A'
    $products[2, 1] = 'SKU-MISSING'
    $destination.Range("${productColumn}4:${productColumn}5").Value2 = $products
    $destination.Cells.Item(1, 2).Value2 = '=SUBTOTAL(9,B6:B1048576)'
    $currentHeaderColumn = if ($Unified) { 30 } else { 46 }
    $futureHeaderColumn = $currentHeaderColumn + 1
    $destination.Cells.Item(3, $currentHeaderColumn).Value2 = 'Keep current header'
    for ($i = 0; $i -lt 7; $i++) {
        $destination.Cells.Item(3, $futureHeaderColumn + $i).Value2 = 'Old date'
    }
    $destination.Range('BN96:BP98').Value2 = 'Keep existing formulas'
    $destination.Range('CF4:CO5').Value2 = 'Keep former inventory cells'
    $destinationWorkbook.Worksheets.Items[$DestinationSheetName] = $destination
    foreach ($name in @('HK OOR Pivot', 'DDR pivot')) {
        $sourceWorkbook.Worksheets.Items[$name] = [RetailLayoutTestSheet]::new()
    }
    $destination.Cells.Item(3, 54).Value2 = 'Keep BB header'
    $destination.Cells.Item(4, 54).Formula2 = '=KEEP_LEGACY_BB()'
    $destination.Cells.Item(4, 54).Value2 = 31
    $destination.Cells.Item(5, 54).Value2 = 0
    $destination.Cells.Item(4, 54).NumberFormat = '#'
    $destination.Cells.Item(5, 54).NumberFormat = '#'
    if ($Unified) {
        $sourceWorkbook.Worksheets.Items['Retail INV'] = New-TestStockSheet
    }
    . $configureLayout
    if ($InvalidHeader) {
        try {
            . $refreshBody
            throw 'An invalid destination layout was accepted.'
        }
        catch {
            if ($_.Exception.Message -notlike '*expected Product ID in N3*') {
                throw
            }
        }
        if ($destination.Writes.Count -or $destinationWorkbook.SaveCount) {
            throw 'The invalid layout was modified or saved.'
        }
        return
    }
    . $refreshBody
    $monthlyAddress = if ($Unified) { 'P4:R5' } else { 'U4:W5' }
    $orderAddress = if ($Unified) { 'AC4:AK5' } else { 'AS4:BA5' }
    $actualAddress = if ($Unified) { 'BA4:BA5' } else { 'BT4:BT5' }
    $expectedWrites = @($monthlyAddress, $orderAddress, $actualAddress)
    if (($destination.Writes -join ',') -ne ($expectedWrites -join ',')) {
        throw "Unexpected destination writes: $($destination.Writes -join ',')"
    }
    $monthlyValues = $destination.Range($monthlyAddress).Value2
    $orderValues = $destination.Range($orderAddress).Value2
    $actualValues = $destination.Range($actualAddress).Value2
    if ($monthlyValues[1, 1] -ne 17 -or $monthlyValues[1, 2] -ne 19 -or
        $monthlyValues[1, 3] -ne 0 -or $orderValues[1, 1] -ne 5 -or
        $orderValues[1, 2] -ne 7 -or $orderValues[1, 3] -ne 11 -or
        $actualValues[1, 1] -ne 23) {
        throw 'Matched product values were not written to the correct fields.'
    }
    foreach ($address in @($monthlyAddress, $orderAddress, $actualAddress)) {
        $written = $destination.Range($address).Value2
        for ($c = 1; $c -le $written.GetLength(1); $c++) {
            if ($written[2, $c] -ne 0) { throw 'Unmatched product did not receive zero.' }
        }
    }
    for ($i = 0; $i -lt 7; $i++) {
        $header = $destination.Cells.Item(3, $futureHeaderColumn + $i)
        if ($i -eq 0) {
            if ($header.Value2 -ne ([datetime]'2026-11-06').ToOADate() -or
                $header.NumberFormat -ne 'mm/dd') {
                throw 'The future weekly header was not moved with its data.'
            }
        }
        elseif ($null -ne $header.Value2) { throw 'An unused future header was not cleared.' }
    }
    if ($destination.Cells.Item(1, 2).Value2 -ne '=SUBTOTAL(9,B6:B1048576)' -or
        $destination.Cells.Item(3, $currentHeaderColumn).Value2 -ne 'Keep current header' -or
        $destination.Range('BN96:BP98').Value2 -ne 'Keep existing formulas' -or
        $destination.Range('CF4:CO5').Value2 -ne 'Keep former inventory cells') {
        throw 'An unrelated formula or protected header was changed.'
    }
    if ($destinationWorkbook.SaveCount -ne 1 -or $excel.Calculation -ne -4105) {
        throw 'The destination was not saved once with Automatic calculation.'
    }
    if ($destination.Cells.Item(3, 54).Value2 -ne 'Keep BB header') {
        throw 'The BB header changed.'
    }
    if ($Unified) {
        foreach ($row in 4..5) {
            $expected = '=XLOOKUP(N{0},INDIRECT("''Retail INV''!A:A"),INDIRECT("''Retail INV''!J:J"),0)' -f $row
            if ($destination.Cells.Item($row, 54).Formula2 -cne $expected -or
                $destination.Cells.Item($row, 54).NumberFormat -ne '#') {
                throw 'The compact BB lookup formula or formatting is incorrect.'
            }
        }
        if ($stockLookups.Count -ne 2 -or $stockLookups[0].Expected -ne 31 -or
            $stockLookups[1].Expected -ne 0) {
            throw 'The BB lookup must use the first case-insensitive match or zero.'
        }
    }
    elseif ($destination.Cells.Item(4, 54).Formula2 -ne '=KEEP_LEGACY_BB()') {
        throw 'The separate legacy report BB formula changed.'
    }
}

Test-RetailLayout -Unified $true
Test-RetailLayout -Unified $false
Test-RetailLayout -Unified $true -InvalidHeader $true

$target = [RetailLayoutTestSheet]::new()
$source = New-TestStockSheet
$products = New-TestExcelArray 3 1
$products[1, 1] = 'SKU-A'
$products[3, 1] = 'SKU-MISSING'
$target.Cells.Item(4, 54).Value2 = 31
$target.Cells.Item(5, 54).Formula2 = '=KEEP_BLANK_PRODUCT_ROW()'
$target.Cells.Item(6, 54).Value2 = 0
$checks = @(Set-RetailStockLookup -Worksheet $target `
    -SourceWorksheet $source -ProductValues $products -StartRow 4)
Assert-RetailStockLookup -Worksheet $target -Checks $checks
if ($checks.Count -ne 2 -or $checks[1].Row -ne 6 -or
    $target.Cells.Item(5, 54).Formula2 -ne '=KEEP_BLANK_PRODUCT_ROW()') {
    throw 'Blank product rows must be preserved without shifting the lookup rows.'
}
foreach ($change in @('Formula2', 'Value2', 'NumberFormat')) {
    $cell = $target.Cells.Item(4, 54)
    $original = $cell.$change
    $cell.$change = if ($change -eq 'Value2') { 99 } else { 'Incorrect' }
    try {
        Assert-RetailStockLookup -Worksheet $target -Checks $checks
        throw "BB validation accepted an incorrect $change."
    }
    catch {
        if ($_.Exception.Message -notlike '*Retail BB4*') { throw }
    }
    finally { $cell.$change = $original }
}
$source.Cells.Item(1, 1).Value2 = 'Wrong header'
try {
    Set-RetailStockLookup -Worksheet $target -SourceWorksheet $source `
        -ProductValues $products -StartRow 4
    throw 'An invalid inventory header was accepted.'
}
catch {
    if ($_.Exception.Message -notlike '*expected Material in A1*') { throw }
}
Write-Output 'Retail layout regression tests passed.'
