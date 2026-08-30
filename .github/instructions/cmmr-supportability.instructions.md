# CMMR Supportability Report Refresh

When the user says "Run CMMR supportability report refresh":

1. Use `scripts/Update-CmmrSupportability.ps1` as the authoritative
   implementation.
2. Run only:

   `powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrSupportability.ps1 -SourceWorkbookPath "<source-path>" -DestinationWorkbookPath "<destination-path>"`

3. Do not create an alternative implementation.
4. Do not perform unnecessary previews, repeated reads, or intermediate checks.
5. Report whether processing and final validation succeeded.
6. If processing or validation fails, report the exact failure and do not
   report success.

## Workbooks

- Source: `1_HK CMMR+RT Order and Supply data CLI.xlsx`
- Destination: `CMMR Supportability Report_CLI+++++.xlsx`
- Destination worksheet: `Supportability`

## Product matching

- Match source `SKU` or `Material` to destination `Product ID` in column `L`.
- Destination product data begins on row 4.
- A source product that is absent from a required PivotTable contributes
  numeric `0`.

## Task 1 - Open orders

Write values to destination columns `S:V`:

- `S`: From `HK OOR Pivot!A:K`, sum values for SoldTo Nbr `5395800`
  and `5395801` by SKU.
- `T`: From `HK OOR Pivot!A:K`, sum values for SoldTo Nbr `5395799`
  by SKU.
- `U`: From `HK OOR Pivot!N:V`, sum values for SoldTo Nbr `5395800`,
  `5395799`, and `5395801` by SKU.
- `V`: From `HK OOR Pivot!N:V`, sum values for SoldTo Nbr `1001089`
  by SKU.

## Task 2 - Monthly Actual SI

- Source: `DDR pivot!A:AC`.
- Use the calendar months currently displayed by the PivotTable, in display
  order.
- For each Material and month, sum Delivery Qty for Sold-to Party `5395800`,
  `5395801`, and `5395799`.
- Write the first three displayed months to destination columns `O`, `P`,
  and `Q`.
- If fewer than three months are displayed, write numeric `0` for each
  missing month.

## Task 3 - DC Available Stock

- Match `CMMR INV` column `A` Material to destination column `L` Product ID.
- Copy the corresponding numeric value from `CMMR INV` column `J` to
  destination column `AN`.

## Task 4 - Dynamic Weekly Actual SI

- Source: `DDR pivot!AI:CZ`.
- Match Material to destination column `L` Product ID.
- Sum Delivery Qty for Sold-to Party `5395800`, `5395801`, and `5395799`.
- Write the result to destination column `AO`.
- Calendar weeks run Monday through Sunday.
- On Monday through Thursday, use the previous complete calendar week.
- On Friday, Saturday, or Sunday, use Monday through Thursday of the current
  calendar week.

## Protection requirements

- Do not change row 1.
- Only update destination cells in columns `O:Q`, `S:V`, `AN`, and `AO`,
  beginning on row 4.
- Write values only and preserve existing formatting.
- Do not change any other formulas, values, cell formatting, number formats,
  conditional formatting, column widths, row heights, hidden/grouped state,
  filters, sorting, outlines, freeze panes, PivotTables, charts, shapes,
  links, names, or workbook structure.
- Open the source workbook read-only.
- Save the destination workbook once.
- Validate all written values after saving.
