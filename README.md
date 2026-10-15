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
- `scripts/Update-CmmrConfirmation.ps1` — CMMR confirmation-only `X:AA`
  refresh entry point
- `.github/instructions/retail-supportability.instructions.md` — Retail and
  combined Supportability refresh requirements
- `scripts/Update-RetailSupportability.ps1` — Retail Supportability refresh
  implementation
- `scripts/Update-HkAllOor.ps1` — `HK all OOR` source replacement and
  `HK OOR Pivot` refresh implementation
- `scripts/Update-SupportabilityReports.ps1` — combined CMMR and Retail entry
  point; expects `HK all OOR` and `HK OOR Pivot` to be refreshed beforehand

## Run

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrRt.ps1 -WorkbookPath "<workbook-path>"
```

## Refresh all HK Surface reports

The shortcut `refresh HK all surface reports` follows
`.github/copilot-instructions.md`: synchronize `main`, ensure Excel is closed,
run the inventory/system confirmation update first, then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\Update-SupportabilityReports.ps1 `
  -SourceWorkbookPath "C:\Users\v-capan\OneDrive - Microsoft\disk C\HK ACCN COM online\3_CLI\1_HK CMMR+RT Order and Supply data CLI.xlsx"
```

The source workbook now also contains the report destinations:
`CMMR Supportability` and `Retail supportability`. The original CMMR
`PL 0922` dependency is copied as `CMMR PL`, with related CMMR formulas
pointing to that local sheet. The price list is not refreshed by the
report scripts. The two original report
workbooks are retained unchanged and are no longer updated by this shortcut.
Existing calculations, formatting rules, and write ranges are unchanged;
the combined refresh still skips CMMR confirmation `X:AA` and `X3:AA3`.

Individual CMMR, Retail, and CMMR confirmation-only entry points also default
to their respective sheets in the source workbook. Explicit separate
destination paths remain supported for legacy use and default to the
original `Supportability` sheet. Individual entry points accept
`-DestinationSheetName` to select a different sheet.
