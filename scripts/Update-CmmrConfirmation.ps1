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
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

& (Join-Path $scriptDirectory 'Update-CmmrSupportability.ps1') `
    -SourceWorkbookPath $SourceWorkbookPath `
    -DestinationWorkbookPath $DestinationWorkbookPath `
    -AsOfDate $AsOfDate `
    -ConfirmationOnly
