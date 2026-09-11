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
- `scripts/Update-CmmrConfirmation.ps1` — CMMR confirmation-only `X:AB`
  refresh entry point
- `.github/instructions/retail-supportability.instructions.md` — Retail and
  combined Supportability refresh requirements
- `scripts/Update-RetailSupportability.ps1` — Retail Supportability refresh
  implementation
- `scripts/Update-HkAllOor.ps1` — `HK all OOR` source replacement and
  `HK OOR Pivot` refresh implementation
- `scripts/Update-SupportabilityReports.ps1` — combined CMMR and Retail entry
  point; refreshes `HK all OOR` before both reports

## Run

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrRt.ps1 -WorkbookPath "<workbook-path>"
```
