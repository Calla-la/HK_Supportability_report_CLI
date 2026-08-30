[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceWorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$CmmrDestinationWorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$RetailDestinationWorkbookPath,

    [datetime]$AsOfDate = (Get-Date)
)

$ErrorActionPreference = 'Stop'
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

& (Join-Path $scriptDirectory 'Update-CmmrSupportability.ps1') `
    -SourceWorkbookPath $SourceWorkbookPath `
    -DestinationWorkbookPath $CmmrDestinationWorkbookPath `
    -AsOfDate $AsOfDate

& (Join-Path $scriptDirectory 'Update-RetailSupportability.ps1') `
    -SourceWorkbookPath $SourceWorkbookPath `
    -DestinationWorkbookPath $RetailDestinationWorkbookPath `
    -AsOfDate $AsOfDate

Write-Host 'CMMR and Retail supportability report refresh succeeded.'
