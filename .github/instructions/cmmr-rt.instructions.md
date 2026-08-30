# CMMR-RT Standard Update

When the user says “Run the CMMR-RT Standard Update”:

1. Use the workbook path provided by the user.
2. Run only:

   `powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrRt.ps1 -WorkbookPath "<provided-path>"`

3. Do not create an alternative implementation.
4. Do not perform unnecessary previews, repeated reads, or intermediate checks.
5. Report whether processing and final validation succeeded.
6. The PowerShell script is the authoritative implementation of this workflow.
