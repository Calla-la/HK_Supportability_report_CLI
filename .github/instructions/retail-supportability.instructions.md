# Retail Supportability Report Refresh

When the user says
"Run CMMR supportability report refresh and Retail supportability report refresh":

1. Use `scripts/Update-SupportabilityReports.ps1` as the authoritative
   combined implementation.
2. Run only:

   `powershell -ExecutionPolicy Bypass -File .\scripts\Update-SupportabilityReports.ps1 -SourceWorkbookPath "<source>"`

3. Do not create an alternative implementation.
4. Report success only when both refreshes and their validations succeed.
5. The combined entry point skips CMMR confirmation columns `X:AA` and
   `X3:AA3`; refresh those separately with
   `scripts/Update-CmmrConfirmation.ps1`.
6. The user manually replaces `HK all OOR` and refreshes `HK OOR Pivot`
   before running the combined entry point. The combined entry point must not
   perform either operation.

The Retail implementation is `scripts/Update-RetailSupportability.ps1`.

## Workbooks

- The source and default destination are the same workbook:
  `1_HK CMMR+RT Order and Supply data CLI.xlsx`.
- CMMR destination worksheet: `CMMR Supportability`.
- Retail destination worksheet: `Retail supportability`.
- Both sheets were copied from the original reports; retain the originals
  unchanged and do not refresh them by default.
- Each implementation opens the unified workbook once read/write and saves,
  closes, and releases it without duplicating the COM workbook reference.
- Explicit legacy `-CmmrDestinationWorkbookPath` and
  `-RetailDestinationWorkbookPath` arguments remain supported. A separate
  destination workbook defaults to its original `Supportability` sheet.
- The individual Retail entry point accepts `-DestinationSheetName` for an
  explicitly selected worksheet; omitting `-DestinationWorkbookPath` uses
  the unified workbook and `Retail supportability`.
- The unified `Retail supportability` sheet uses the compact layout below
  after column deletions.
- An explicitly selected separate legacy destination retains its original
  layout: Product ID `P`, monthly actuals `U:W`, open orders `AS:BA`, future
  headers `AU3:BA3`, and weekly actuals `BT`.
- DC inventory refresh logic has been removed for all Retail destinations,
  including the legacy `CF`, `CH`, and `CJ` columns. Do not retain a legacy
  inventory branch or recreate the removed columns.
- Before any destination write, confirm that row 3 of the expected Product
  ID column contains `Product ID`; stop on a layout mismatch.

## Retail product matching

- Match source SKU or Material to destination `Retail supportability` Product ID in
  column `N`.
- Destination product data begins at row 4.
- Write numeric `0` for unmatched Product IDs.

## Retail Task 1 - DS Open Order

- Source: `HK OOR pivot!N:V`.
- Sum Open Order Qty for SoldTo Nbr `5395802` by SKU.
- Write values to destination column `AC`.

## Retail Task 2 - Weekly Open Order

- Source: `HK OOR pivot!Y:AP`.
- Use Requested Del. Dte and Open Order Qty for SoldTo Nbr `5395802`.
- `AD` contains quantities dated on or before the refresh week's Friday,
  including overdue quantities. Do not change `AD3`.
- Group future quantities into calendar-week buckets ending Friday.
- Skip future buckets whose total quantity is zero.
- Write up to seven non-empty future buckets consecutively to `AE:AK`.
- Write each retained Friday to `AE3:AK3` with number format `mm/dd`.
- Clear unused `AE3:AK3` headers and write numeric zero to unused data cells.
- Preserve the existing number format in `AD:AK`.

## Retail Task 3 - Monthly Actual SI

- Source: `DDR pivot!A:AC`.
- Use displayed month order and Delivery Qty for Sold-to Party `5395802`.
- Write the first three months to destination columns `P:R`.
- Write numeric zero for missing months.

## Retail Task 4 - DC Available Stock (removed)

This task is no longer used. Do not aggregate or copy inventory values into
the former inventory columns for any destination. Inventory sheet generation
in `Update-CmmrRt.ps1` is unchanged. The new unified-workbook BB lookup below
is separate from this removed task.

## Retail Task 5 - Dynamic Weekly Actual SI

- Source: `DDR pivot!AI:CZ`.
- Sum Delivery Qty for Sold-to Party `5395802`.
- Write the result to destination column `BA`.
- Format `BA` as a zero-decimal number and display numeric zero as blank.
- Monday through Thursday: use the previous Monday through Sunday.
- Friday through Sunday: use the current Monday through Friday.

## Protection requirements

- For the unified workbook, only write destination columns `P:R`, `AC:AK`,
  `BA`, and the BB product-row formulas below from row 4 onward, plus headers
  `AE3:AK3`.
- Do not change row 1 or `AD3`.
- Preserve all existing cell formatting and number formats, except set
  populated `AE3:AK3` headers to `mm/dd` and format `BA` as a zero-decimal
  number with numeric zero displayed as blank.
- Do not change any other values, formulas, conditional formatting, widths,
  heights, grouping, filters, sorting, outlines, freeze panes, PivotTables,
  charts, shapes, links, names, or workbook structure.
- Open a separate source workbook read-only, or open the unified source and
  destination once read/write. Save the destination once and validate every
  written value after saving.
- Restore Automatic calculation before saving and validate that it remains
  Automatic afterward.
- Retry graceful Excel Quit twice; report any remaining script-owned Excel
  PID after COM release without force-terminating it.
- Retry transient rejected Excel COM operations for at most 10 seconds, then
  report the exact operation that failed.

## Retail BB - Live inventory lookup

- On every unified-workbook refresh, write XLOOKUP formulas in `BB` only for
  rows with a nonblank Product ID in `N`, beginning at row 4.
- Match `N` against the same workbook's `Retail INV!A:A` (Material), returning
  `Retail INV!J:J` (DTV Supply without open DO & Open Order).
- Use `=XLOOKUP(N4,INDIRECT("'Retail INV'!A:A"),INDIRECT("'Retail INV'!J:J"),0)`
  for row 4, with the corresponding row number for each product.
- Keep INDIRECT: the preceding inventory refresh deletes and recreates
  `Retail INV`; direct sheet references would become `#REF!`.
- Use the first match, case-insensitively; unmatched products return zero.
- Preserve BB formatting, blank-product rows, BB3, and row 1. Existing `#`
  number formats display numeric zero as blank.
- Validate the source Material header, each written formula and its result,
  and unchanged BB number formats before and after the single save.
- Do not add BB formulas to explicitly selected separate legacy reports.
