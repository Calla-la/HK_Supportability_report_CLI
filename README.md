# HK Supportability Report CLI

Automates the CMMR and Retail inventory update workflow.

## Files

- `.github/instructions/cmmr-rt.instructions.md` — workflow requirements
- `scripts/Update-CmmrRt.ps1` — PowerShell implementation

## Run

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrRt.ps1 -WorkbookPath "<workbook-path>"
