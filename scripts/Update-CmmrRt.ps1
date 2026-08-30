[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$WorkbookPath
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceDirectory = Split-Path -Parent $resolvedWorkbookPath

$cmmrSourcePath = Join-Path $sourceDirectory 'CMMR_INV_RAW_DATA.XLS'
$retailSourcePath = Join-Path $sourceDirectory 'RETAIL_INV_RAW_DATA.xls'

if (-not (Test-Path -LiteralPath $cmmrSourcePath -PathType Leaf)) {
    throw "Source file not found: $cmmrSourcePath"
}

if (-not (Test-Path -LiteralPath $retailSourcePath -PathType Leaf)) {
    throw "Source file not found: $retailSourcePath"
}

throw 'CMMR-RT update logic has not been implemented yet.'
