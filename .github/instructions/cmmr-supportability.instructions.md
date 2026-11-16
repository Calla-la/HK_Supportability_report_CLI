# CMMR Supportability Report Refresh

When the user says "Run CMMR supportability report refresh":

1. Use `scripts/Update-CmmrSupportability.ps1` as the authoritative
   implementation.
2. Run only:

   `powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrSupportability.ps1 -SourceWorkbookPath "<source-path>"`

3. Do not create an alternative implementation.
4. Do not perform unnecessary previews, repeated reads, or intermediate checks.
5. Report whether processing and final validation succeeded.
6. If processing or validation fails, report the exact failure and do not
   report success.

CMMR `X:AA` refresh functionality, including its separate entry point and
IB PO confirmation calculations, has been removed. Neither the individual
nor the combined supportability refresh updates `X:AA` or `X3:AA3`.
Preserve all existing values, formulas, headers, and formatting there.

## Workbooks

For an explicitly requested monthly-only refresh, run the authoritative
script with `-MonthlyOnly`. This mode reads only `DDR pivot!A:AC` and the
destination Product IDs, writes and validates only the monthly data range
(`N:P` in the unified workbook), and does not refresh open orders, inventory,
weekly actuals, or Retail.

- Source: `1_HK CMMR+RT Order and Supply data CLI.xlsx`
- Destination: the source workbook itself
- Destination worksheet: `CMMR Supportability`
- Price-list dependency: original `PL 0922` was copied as `CMMR PL` in the
  unified workbook. Related CMMR formulas reference `CMMR PL` locally; do not
  recopy or refresh the price list during report refreshes.
- The original `CMMR Supportability Report_CLI+++++.xlsx` is retained
  unchanged, not used as the default destination.
- The unified workbook uses monthly actuals `N:P`, open orders `R:U`,
  inventory `AK`, and weekly actuals `AL`. Do not write to `V` or `AM`
  and any columns to its right.
- Explicit separate legacy destinations retain monthly actuals `O:Q`,
  open orders `S:V`, inventory `AM`, and weekly actuals `AN`.
- For an explicitly requested legacy separate-workbook refresh,
  `-DestinationWorkbookPath "<destination-path>"` remains supported and
  defaults to worksheet `Supportability`. `-DestinationSheetName` can
  explicitly select a worksheet.
- When source and destination are the same file, open it once read/write.
  Never open a second read-only copy or close/release the same COM workbook
  twice. For separate files, keep the source read-only as before.

## Product matching

- Match source `SKU` or `Material` to destination `Product ID` in column `L`.
- Destination product data begins on row 4.
- A source product that is absent from a required PivotTable contributes
  numeric `0`.

## Task 1 - Open orders

For an explicitly requested open-orders-only update, use
`-OpenOrdersOnly` with the authoritative CMMR script. This mode writes and
validates only `R:U` in the unified destination, preserving formatting and
skipping all other report updates and Retail.
Do not combine it with another `Only` mode.

Write values to unified destination columns `R:U`:

- `R`: From `HK OOR Pivot!A:K`, sum values for SoldTo Nbr `5395800`
  and `5395801` by SKU.
- `S`: From `HK OOR Pivot!A:K`, sum values for SoldTo Nbr `5395799`
  by SKU.
- `T`: From `HK OOR Pivot!N:V`, sum values for SoldTo Nbr `5395800`,
  `5395799`, and `5395801` by SKU.
- `U`: From `HK OOR Pivot!A:K`, sum values for SoldTo Nbr `1001089`
  by SKU.

## Task 2 - Monthly Actual SI

- Use this same monthly logic for the full CMMR refresh, the combined
  supportability refresh (including `refresh HK all surface reports`), and
  `-MonthlyOnly`. It is not a one-off monthly-only correction.
- Source: `DDR pivot!A:AC`.
- Use the calendar months currently displayed by the PivotTable, in display
  order.
- For each SKU/Material and month, sum Delivery Qty for Sold-to Party
  `5395800`, `5395801`, and `5395799`, matching destination Product ID in `L`.
- Write the first three displayed months to unified destination columns
  `N`, `O`, and `P` (October, November, December for the current quarter).
- If fewer than three months are displayed, write numeric `0` for each
  missing month, including the second and third months when absent.
- In the unified destination, format `N:P` product data from row 4 onward
  as `#,##0;-#,##0;;@` so numeric zero displays as blank. Do not write empty
  strings instead of numeric zero or change row 1.

## Task 3 - DC Available Stock

- For an explicitly requested inventory-only update, use `-InventoryOnly`
  with the authoritative CMMR script. Read `CMMR INV` and destination Product
  IDs, write and validate only unified destination `AK` from row 4 onward,
  preserve formatting, and skip all other refresh tasks. Do not combine
  this switch with another `Only` mode.
- Match `CMMR INV` column `A` Material to destination column `L` Product ID.
- Copy the corresponding numeric value from `CMMR INV` column `J` to
  unified destination column `AK`.

## Task 4 - Dynamic Weekly Actual SI

- For an explicitly requested weekly-actual-only update, use `-WeeklyOnly`
  with the authoritative CMMR script. Write and validate only unified
  destination `AL` from row 4 onward, preserving formatting and skipping
  all other refresh tasks. Do not combine it with another `Only` mode.
- Source: `DDR pivot!AI:CZ`.
- Match Material to destination column `L` Product ID.
- Sum Delivery Qty for Sold-to Party `5395800`, `5395801`, and `5395799`.
- Write the result to unified destination column `AL`.
- Calendar weeks run Monday through Sunday.
- On Monday through Thursday, use the previous complete calendar week.
- On Friday, Saturday, or Sunday, use Monday through Thursday of the current
  calendar week.

## Protection requirements

- Do not change row 1.
- Only update unified destination cells in columns `N:P`, `R:U`,
  and `AK:AL`, beginning on row 4.
- Do not modify `X:AA`, including headers `X3:AA3`.
- Preserve `V` and `AM` and all columns to its right; do not clear their
  existing contents or delete columns.
- Do not modify destination columns `AB:AG`.
- Write values only and preserve existing formatting.
- The only formatting exception is the unified destination `N:P` number
  format required to display underlying numeric zero values as blank.
- Do not change any other formulas, values, cell formatting, number formats,
  conditional formatting, column widths, row heights, hidden/grouped state,
  filters, sorting, outlines, freeze panes, PivotTables, charts, shapes,
  links, names, or workbook structure.
- Open a separate source workbook read-only; open the unified source and
  destination workbook once read/write.
- Save the destination workbook once.
- Validate all written values after saving.
- Restore Automatic calculation before each workbook save and validate that
  it remains Automatic afterward.
- Retry graceful Excel Quit twice; report any remaining script-owned Excel
  PID after COM release without force-terminating it.
- Retry rejected Excel COM calls for at most 10 seconds per call.
