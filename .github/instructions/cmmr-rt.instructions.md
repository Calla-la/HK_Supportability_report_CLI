# CMMR-RT Standard Update

When the user says "Run the CMMR-RT Standard Update":

1. Use the workbook path provided by the user.
2. Use `scripts/Update-CmmrRt.ps1` as the authoritative implementation.
3. The two source files are located in the same directory as the target workbook:
   - `CMMR_INV_RAW_DATA.XLS`
   - `RETAIL_INV_RAW_DATA.xls`
4. Run only:

   `powershell -ExecutionPolicy Bypass -File .\scripts\Update-CmmrRt.ps1 -WorkbookPath "<provided-path>"`

5. Do not create an alternative implementation.
6. Do not perform unnecessary previews, repeated reads, or intermediate checks.
7. Report whether processing and final validation succeeded.
8. If processing or validation fails, report the exact failure and do not report success.

## Performance requirements

- Use a hash table keyed by Material and aggregate each source file in one pass.
- Do not use `Group-Object`.
- Treat both `.XLS` source files as tab-delimited text files.
- Normalize SAP-style numbers such as `- 184.000`.
- Start Excel COM only once.
- Write data using 2D arrays in bulk.
- Save the workbook once.
- Perform one final validation after saving.
- Avoid unnecessary previews, repeated reads, and intermediate checks.

## Process CMMR_INV_RAW_DATA.XLS

1. Delete the existing `CMMR INV` sheet.
2. Exclude rows where `SpecStkInd` is populated.
3. Aggregate by `Material`.
4. Sum `QtyOpenDeliveries` across all eligible rows.
5. Sum `QtyOpenOrders` across all eligible rows.
6. Sum `QtyUnrestr` only for rows where `SLoc = 0010`.
7. Keep materials that have at least one row where `SLoc = 0010`.
8. Keep materials without `SLoc = 0010` only when `QtyOpenDeliveries` or `QtyOpenOrders` is non-zero.
9. Output `SLoc` as the numeric value `10`.
10. For materials without `SLoc = 0010`, output `QtyUnrestr` as `0`.

## Output columns

Output the following columns in this exact order:

1. Material
2. Description
3. Plant
4. SLoc
5. QtyUnrestr
6. QtyOpenDeliveries
7. QtyOpenOrders
8. DTV Supply
9. DTV Supply without open DO
10. DTV Supply without open DO & Open Order

## Calculations

Calculate the output columns as follows:

- `DTV Supply = QtyUnrestr`
- `DTV Supply without open DO = QtyUnrestr - QtyOpenDeliveries`
- `DTV Supply without open DO & Open Order = QtyUnrestr + QtyOpenOrders`

## Create CMMR INV sheet

1. Create the replacement sheet as `CMMR INV`.
2. Group and collapse columns `C:G`.
3. Format numeric values with zero decimal places.
4. Use a light-green, bold, black header.
5. Set every column width to `23`.
6. Set row 1 height to `30`.
7. Wrap the header text.
8. Freeze row 1.
9. Add filters to row 1.

## Process RETAIL_INV_RAW_DATA.xls

Apply the same filtering, aggregation, output columns, calculations, and formatting rules used for `CMMR_INV_RAW_DATA.XLS`, with these differences:

1. Delete the existing `Retail INV` sheet.
2. Create the replacement sheet as `Retail INV`.
3. Use a light-blue, bold, black header.

## Saving

Save the completed workbook to the original path supplied through the `-WorkbookPath` parameter.

Do not create another output workbook.

## Final validation

Confirm all of the following:

- `CMMR INV` exists exactly once.
- `Retail INV` exists exactly once.
- Every output `SLoc` value is numeric `10`.
- All calculated columns contain the expected values.
- Filters are enabled on row 1.
- Row 1 is frozen.
- Columns `C:G` are grouped and collapsed.
- Every output column width is `23`.
- Row 1 height is `30`.
- Numeric columns use a zero-decimal number format.
- `CMMR INV` has a light-green, bold, black header.
- `Retail INV` has a light-blue, bold, black header.
- Header text wrapping is enabled.
