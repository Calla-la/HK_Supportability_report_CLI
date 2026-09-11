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

When the user asks to refresh only CMMR confirmation columns `X:AB`, run:

`powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrConfirmation.ps1 -SourceWorkbookPath "<source-path>" -DestinationWorkbookPath "<destination-path>"`

This entry point uses the authoritative CMMR implementation in confirmation-only
mode. It refreshes the IB PO source data, then updates and validates only
`Supportability!X:AB` and `Y3:AB3`.

The combined `Update-SupportabilityReports.ps1` entry point invokes the CMMR
implementation with `-SkipConfirmation`, so it does not read, calculate,
format, write, or validate `Supportability!X:AB` or `Y3:AB3`.

## Workbooks

- Source: `1_HK CMMR+RT Order and Supply data CLI.xlsx`
- IB PO source: `HK ALL IB PO.XLSX`, in the same directory as the source
  workbook
- Destination: `CMMR Supportability Report_CLI+++++.xlsx`
- Destination worksheet: `Supportability`

## Task 0 - Refresh IB PO source data

Before calculating the supportability report:

1. Open `HK ALL IB PO.XLSX` and use its first worksheet.
2. Set column `W` header to `IB PO RDD`.
3. For every populated date in column `P`, write the Friday of that date's
   Monday-through-Sunday calendar week to column `W`.
4. Replace the `IB PO` worksheet in the source workbook with the processed
   worksheet.
5. Refresh the `IB PO pvt` PivotTable and save the source workbook.
6. Reopen the source workbook read-only before running the remaining tasks.

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
- `V`: From `HK OOR Pivot!A:K`, sum values for SoldTo Nbr `1001089`
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
  destination column `AL`.

## Task 4 - Dynamic Weekly Actual SI

- Source: `DDR pivot!AI:CZ`.
- Match Material to destination column `L` Product ID.
- Sum Delivery Qty for Sold-to Party `5395800`, `5395801`, and `5395799`.
- Write the result to destination column `AM`.
- Calendar weeks run Monday through Sunday.
- On Monday through Thursday, use the previous complete calendar week.
- On Friday, Saturday, or Sunday, use Monday through Thursday of the current
  calendar week.

## Task 5 - CMMR Confirmation Weekly Buckets

Update destination worksheet `Supportability`, columns `X:AB`, beginning on
row 4.

### Source and matching

- Source worksheet: `CMMR confirmation`.
- Match source `SKU` in column `X` to destination `Product ID` in column `L`.
- Only include rows where:
  - Plant in source column `AS` is `4815`.
  - Sold-to Nbr in source column `C`, ignoring leading zeros, is `5395801`,
    `5395800`, or `5395799`.
- Use source column `BX` as `Date to release`.
- Use source column `AD` as the quantity.
- For each SKU and BX date, sum AD across all three eligible Sold-to Nbr
  values.

### Week mapping

Calendar weeks run Monday through Sunday based on the refresh date:

- `X`: current calendar week.
- `Y:AB`: the earliest four future calendar weeks that contain eligible AD
  data, in chronological order.
- Skip future calendar weeks with no eligible AD data and shift later
  non-empty weeks left.
- Write each selected future week's Friday to `Y3:AB3`.
- Clear unused `Y3:AB3` cells and write numeric zero to the corresponding
  unused data columns.
- Store missing AD results in `Y:AB` as numeric zero, but apply a number
  format that displays zero values as blank.

### IB PO confirmation

Use source worksheet `IB PO pvt` for confirmation:

- Use the PivotTable filtered for Plant `4815`.
- Match by SKU/Material.
- Match the BX date to the corresponding `Stat.-Rel. Del. Date`.
- Compare the PivotTable quantity at the SKU/date intersection with the
  aggregated AD quantity.
- Ignore leading zeros when matching numeric identifiers.

For destination column `X` only:

1. Use only valid `Date to release` dates in the current Monday-through-Sunday
   calendar week. Exclude blank, `TBD`, and `Delivered`.
2. For each SKU, sum column `AD` across the entire current week and across
   eligible Sold-to Nbr `5395800`, `5395799`, and `5395801`.
3. Compare that weekly AD total with:
   - the same SKU's `CMMR INV` column `I` value; and
   - the same SKU's `IB PO pvt` `Still to be delivered (qty)` for the
     current week's Friday `IB PO RDD`.
4. Write the weekly AD total to `X` when either comparison value is greater
   than or equal to the AD total.
5. When both conditions pass, count the weekly AD total only once.

For destination columns `Y:AB`:

1. Exclude the current calendar week.
2. For each future calendar week and SKU, sum column `AD` across eligible
   Sold-to Nbr `5395800`, `5395799`, and `5395801`.
3. Compare the weekly AD total with the same SKU's `IB PO pvt`
   `Still to be delivered (qty)` for that week's Friday `IB PO RDD`.
4. Include the weekly AD total when it is less than or equal to the IB PO
   quantity; otherwise write numeric zero for that SKU.
5. Select the earliest four future weeks containing eligible AD data.

Sum all qualifying SKU/date quantities within each week and write numeric
values to `X:AB`. Write numeric `0` when no qualifying quantity exists.

## Protection requirements

- Do not change row 1.
- Only update destination cells in columns `O:Q`, `S:V`, `X:AB`, `AL`, and
  `AM`, beginning on row 4, plus headers `Y3:AB3`.
- Do not modify destination columns `AC:AH`.
- Write values only and preserve existing formatting.
- The only formatting exception is the `Y:AB` number format required to
  display underlying numeric zero values as blank.
- Do not change any other formulas, values, cell formatting, number formats,
  conditional formatting, column widths, row heights, hidden/grouped state,
  filters, sorting, outlines, freeze panes, PivotTables, charts, shapes,
  links, names, or workbook structure.
- Open the source workbook read-only.
- Save the destination workbook once.
- Validate all written values after saving.
- Retry rejected Excel COM calls for at most 10 seconds per call.
