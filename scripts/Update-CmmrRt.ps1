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

        $Worksheet.Activate()
        $window = $Application.ActiveWindow
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

        $Worksheet.Activate()
        $window = $Application.ActiveWindow
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

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceDirectory = Split-Path -Parent $resolvedWorkbookPath
$cmmrSourcePath = Join-Path $sourceDirectory 'CMMR_INV_RAW_DATA.XLS'
$retailSourcePath = Join-Path $sourceDirectory 'RETAIL_INV_RAW_DATA.xls'

foreach ($sourcePath in @($cmmrSourcePath, $retailSourcePath)) {
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

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Open($resolvedWorkbookPath, 0, $false)
    if ([bool]$workbook.ReadOnly) {
        throw "Workbook opened as read-only and cannot be updated: $resolvedWorkbookPath"
    }

    $cmmrWorksheet = $workbook.Worksheets.Add()
    $cmmrWorksheet.Name = '__CMMR_INV_TEMP__'
    $retailWorksheet = $workbook.Worksheets.Add()
    $retailWorksheet.Name = '__RETAIL_INV_TEMP__'

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

    $workbook.Save()

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

    Write-Host "CMMR-RT update and final validation succeeded: $resolvedWorkbookPath"
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }

    foreach ($comObject in @($retailWorksheet, $cmmrWorksheet, $workbook, $excel)) {
        Release-ComObject -ComObject $comObject
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
