# HK report refresh shortcut

When the user says `refresh HK all surface reports` (case-insensitive),
execute the complete workflow below. Do not ask the user to paste the
individual commands again.

Requests to document or rename this shortcut are documentation-only requests;
do not run the refresh unless the user requests execution.

## Execution order

1. Switch to the repository and inspect the Git working tree. Preserve all
   existing changes; do not automatically stash, reset, discard, commit, or
   overwrite them. If they prevent switching branches or pulling, stop and
   report the blocker.

   ```powershell
   cd "C:\Users\v-capan\HK_Supportability_report_CLI"
   git status --short
   git switch main
   git pull --ff-only origin main
   ```

   Check each command's result before continuing. If switching or pulling
   fails, stop; do not run the report scripts against an unsynchronized branch.
   Read the repository instructions again after pulling.

2. Ensure all Microsoft Excel windows are closed before automation starts.
   Close Excel gracefully when safe. If an open workbook has unsaved changes,
   stop and ask how to handle them; never discard them or force-terminate Excel.

3. Run the CMMR-RT inventory and system confirmation update first:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrRt.ps1 `
     -WorkbookPath "C:\Users\v-capan\OneDrive - Microsoft\disk C\HK ACCN COM online\3_CLI\1_HK CMMR+RT Order and Supply data CLI.xlsx"
   ```

4. Only after the preceding script and its final validation succeed, run
   the CMMR and Retail supportability refresh:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\scripts\Update-SupportabilityReports.ps1 `
     -SourceWorkbookPath "C:\Users\v-capan\OneDrive - Microsoft\disk C\HK ACCN COM online\3_CLI\1_HK CMMR+RT Order and Supply data CLI.xlsx"
   ```

   This workbook is also the destination. Update `CMMR Supportability` and
   `Retail supportability` in place. The original separate report workbooks
   are retained as archives and must not be refreshed by this shortcut.
   `CMMR PL` is the copied `PL 0922` price-list dependency; CMMR report
   formulas use this local sheet. Do not refresh or recopy the price list as
   part of this shortcut.

## Workflow boundaries

- Use these existing scripts as the authoritative implementations. Do not
  introduce replacement logic, extra refresh scripts, or file-copy/overwrite
  workarounds without user approval.
- Follow the detailed rules in `.github\instructions\cmmr-rt.instructions.md`,
  `.github\instructions\cmmr-supportability.instructions.md`, and
  `.github\instructions\retail-supportability.instructions.md`.
- Preserve the combined supportability entry point's existing scope: it
  skips `CMMR Supportability!X:AA` and `X3:AA3`. Do not add the separate
  confirmation-only script to this shortcut.
- The user manually replaces `HK all OOR` and refreshes `HK OOR Pivot`
  before the combined supportability refresh. Do not automate those steps.
- Preserve existing report formatting and row 1 formulas, except for changes
  explicitly required by the detailed workflow instructions.
- Restore Automatic calculation before saving workbooks.
- Reuse the established Excel COM workflow. Open CMMR allregion OOR directly
  with Excel COM; do not try openpyxl first.
- Retry graceful Excel Quit as specified by the existing scripts. Report any
  remaining script-owned Excel PID without force-terminating it.
- Stop on a processing or validation failure and report the exact failed
  step. Report full success only after both scripts and their final
  validations succeed; clearly identify any already-saved partial results.
