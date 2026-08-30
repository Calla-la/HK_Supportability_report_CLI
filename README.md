# HK Supportability Report CLI

Automates the CMMR-RT inventory and system confirmation update workflow.

## Files

- `.github/instructions/cmmr-rt.instructions.md` — inventory and system
  confirmation workflow requirements
- `scripts/Update-CmmrRt.ps1` — PowerShell implementation
- `.github/instructions/cmmr-supportability.instructions.md` — CMMR
  Supportability refresh requirements
- `scripts/Update-CmmrSupportability.ps1` — CMMR Supportability refresh
  implementation

## Run

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrRt.ps1 -WorkbookPath "<workbook-path>"
```
