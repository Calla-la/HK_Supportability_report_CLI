$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot '..\scripts\Update-CmmrRt.ps1'
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $scriptPath, [ref]$tokens, [ref]$parseErrors
)
if ($parseErrors.Count -gt 0) {
    throw ($parseErrors | Out-String)
}
$definition = $ast.Find({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'Set-WorksheetHeaderFreeze'
}, $true)
if ($null -eq $definition) {
    throw 'Set-WorksheetHeaderFreeze was not found.'
}
Invoke-Expression $definition.Extent.Text

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public sealed class HeaderFreezeTestWindow
{
    public int FailuresRemaining;
    public bool AlwaysFail;
    public int ErrorCode = unchecked((int)0x800A03EC);
    public int Attempts;
    public int Activations;
    public int SplitColumn;
    public int SplitRow;
    private bool freezePanes;

    public void Activate() { Activations++; }

    public bool FreezePanes
    {
        get { return freezePanes; }
        set
        {
            Attempts++;
            if (AlwaysFail || FailuresRemaining-- > 0)
                throw new COMException("Simulated pane-setting failure", ErrorCode);
            freezePanes = value;
        }
    }
}

public sealed class HeaderFreezeTestWorksheet
{
    public string Name = "Inventory test";
    public int Activations;
    public void Activate() { Activations++; }
}
'@

foreach ($failureCount in @(0, 2)) {
    $window = New-Object HeaderFreezeTestWindow
    $window.FailuresRemaining = $failureCount
    $window.SplitColumn = 3
    $window.SplitRow = 5
    $sheet = New-Object HeaderFreezeTestWorksheet
    Set-WorksheetHeaderFreeze -Worksheet $sheet -Window $window
    if (-not $window.FreezePanes -or $window.SplitColumn -ne 0 -or
        $window.SplitRow -ne 1 -or $window.Activations -ne ($failureCount + 1) -or
        $sheet.Activations -ne $window.Activations) {
        throw "Incorrect frozen-header state after $failureCount transient failures."
    }
}

foreach ($errorCode in @(-2146827284, -2147024891)) {
    $window = New-Object HeaderFreezeTestWindow
    $window.AlwaysFail = $true
    $window.ErrorCode = $errorCode
    $sheet = New-Object HeaderFreezeTestWorksheet
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $caught = $false
    try {
        Set-WorksheetHeaderFreeze -Worksheet $sheet -Window $window
    }
    catch [System.Runtime.InteropServices.COMException] {
        if ($_.Exception.HResult -ne $errorCode) { throw }
        $caught = $true
    }
    if (-not $caught) {
        throw "Pane-setting error $errorCode was swallowed."
    }
    if ($errorCode -eq -2146827284) {
        if ($timer.ElapsedMilliseconds -lt 10000 -or
            $timer.ElapsedMilliseconds -gt 15000 -or $window.Attempts -lt 2) {
            throw 'Transient pane-setting failure did not respect the retry deadline.'
        }
    }
    elseif ($window.Attempts -ne 1) {
        throw 'An unrelated COM error was retried.'
    }
}

Write-Output 'Header freeze regression tests passed.'
