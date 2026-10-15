[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceWorkbookPath,

    [ValidateNotNullOrEmpty()]
    [string]$CmmrDestinationWorkbookPath = $SourceWorkbookPath,

    [ValidateNotNullOrEmpty()]
    [string]$RetailDestinationWorkbookPath = $SourceWorkbookPath,

    [datetime]$AsOfDate = (Get-Date)
)

$ErrorActionPreference = 'Stop'
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

& (Join-Path $scriptDirectory 'Update-CmmrSupportability.ps1') `
    -SourceWorkbookPath $SourceWorkbookPath `
    -DestinationWorkbookPath $CmmrDestinationWorkbookPath `
    -AsOfDate $AsOfDate `
    -SkipConfirmation

& (Join-Path $scriptDirectory 'Update-RetailSupportability.ps1') `
    -SourceWorkbookPath $SourceWorkbookPath `
    -DestinationWorkbookPath $RetailDestinationWorkbookPath `
    -AsOfDate $AsOfDate

Write-Host 'CMMR and Retail supportability report refresh succeeded.'
