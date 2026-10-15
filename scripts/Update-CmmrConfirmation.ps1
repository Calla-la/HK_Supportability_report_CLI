[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceWorkbookPath,

    [ValidateNotNullOrEmpty()]
    [string]$DestinationWorkbookPath = $SourceWorkbookPath,

    [datetime]$AsOfDate = (Get-Date),

    [string]$DestinationSheetName
)

$ErrorActionPreference = 'Stop'
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$sheetParameters = @{}
if ($PSBoundParameters.ContainsKey('DestinationSheetName')) {
    $sheetParameters.DestinationSheetName = $DestinationSheetName
}

& (Join-Path $scriptDirectory 'Update-CmmrSupportability.ps1') `
    -SourceWorkbookPath $SourceWorkbookPath `
    -DestinationWorkbookPath $DestinationWorkbookPath `
    -AsOfDate $AsOfDate `
    -ConfirmationOnly @sheetParameters
