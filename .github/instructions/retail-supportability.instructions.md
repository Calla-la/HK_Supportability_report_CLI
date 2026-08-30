# Retail Supportability Report Refresh

When the user says
"Run CMMR supportability report refresh and Retail supportability report refresh":

1. Use `scripts/Update-SupportabilityReports.ps1` as the authoritative
   combined implementation.
2. Run only:

   `powershell -ExecutionPolicy Bypass -File .\scripts\Update-SupportabilityReports.ps1 -SourceWorkbookPath "<source>" -CmmrDestinationWorkbookPath "<cmmr-destination>" -RetailDestinationWorkbookPath "<retail-destination>"`

3. Do not create an alternative implementation.
4. Report success only when both refreshes and their validations succeed.

The Retail implementation is `scripts/Update-RetailSupportability.ps1`.

## Retail product matching

- Match source SKU or Material to destination `Supportability` Product ID in
  column `P`.
- Destination product data begins at row 4.
- Write numeric `0` for unmatched Product IDs.

## Retail Task 1 - DS Open Order

- Source: `HK OOR pivot!N:V`.
- Sum Open Order Qty for SoldTo Nbr `5395802` by SKU.
- Write values to destination column `AS`.

## Retail Task 2 - Weekly Open Order

- Source: `HK OOR pivot!Y:AP`.
- Use Requested Del. Dte and Open Order Qty for SoldTo Nbr `5395802`.
- `AT` contains quantities dated on or before the refresh week's Friday,
  including overdue quantities. Do not change `AT3`.
- Group future quantities into calendar-week buckets ending Friday.
- Skip future buckets whose total quantity is zero.
- Write up to seven non-empty future buckets consecutively to `AU:BA`.
- Write each retained Friday to `AU3:BA3` with number format `mm/dd`.
- Clear unused `AU3:BA3` headers and write numeric zero to unused data cells.
- Preserve the existing number format in `AT:BA`.

## Retail Task 3 - Monthly Actual SI

- Source: `DDR pivot!A:AC`.
- Use displayed month order and Delivery Qty for Sold-to Party `5395802`.
- Write the first three months to destination columns `U:W`.
- Write numeric zero for missing months.

## Retail Task 4 - DC Available Stock

Match `Retail INV` Material to destination Product ID:

- Source `H` to destination `CJ`.
- Source `I` to destination `CL`.
- Source `J` to destination `CN`.

## Retail Task 5 - Dynamic Weekly Actual SI

- Source: `DDR pivot!AI:CZ`.
- Sum Delivery Qty for Sold-to Party `5395802`.
- Write the result to destination column `BT`.
- Monday through Thursday: use the previous Monday through Sunday.
- Friday through Sunday: use the current Monday through Friday.

## Protection requirements

- Only write destination columns `U:W`, `AS:BA`, `BT`, `CJ`, `CL`, and `CN`
  from row 4 onward, plus headers `AU3:BA3`.
- Do not change `AT3`.
- Preserve all existing cell formatting and number formats, except set
  populated `AU3:BA3` headers to `mm/dd`.
- Do not change any other values, formulas, conditional formatting, widths,
  heights, grouping, filters, sorting, outlines, freeze panes, PivotTables,
  charts, shapes, links, names, or workbook structure.
- Open the source workbook read-only, save the destination once, and validate
  every written value after saving.
