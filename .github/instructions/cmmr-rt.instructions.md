# CMMR-RT Inventory and System Confirmation Update

When the user says "Run the CMMR-RT inventory and system confirmation update":

1. Use the workbook path provided by the user.
2. Use `scripts/Update-CmmrRt.ps1` as the authoritative implementation.
3. The two source files are located in the same directory as the target workbook:
   - `CMMR_INV_RAW_DATA.XLS`
   - `RETAIL_INV_RAW_DATA.xls`
   The two system-confirmation workbooks are also located in that directory:
   - `CMMR allregion OOR.xlsx`
   - `RT ALLRG OOR.xlsx`
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
- Open all three target workbooks in the same Excel COM instance.
- Complete one successful save for each target workbook.
- Retry only Excel COM save failures, up to three attempts per workbook, and
  report the exact workbook name if all attempts fail.

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

## System confirmation update

Apply the following rules to `CMMR allregion OOR.xlsx` and
`RT ALLRG OOR.xlsx`:

1. Update the first worksheet.
2. Name column `BX` as `Date to release`.
3. In the `Plant` column, fill each blank cell with the value from the cell
   above.
4. Trim Plant values and convert the entire Plant column to numeric values so
   PivotTables do not retain separate text and numeric items for the same
   Plant.
5. Group rows by the same `Order Nbr` and `SKU`.
6. If a group contains only `Sched Line # = 0`, retain those rows for the
   Date-to-release calculation.
7. If a group contains both zero and non-zero schedule lines, clear
   `Date to release` for rows where `Sched Line # = 0` and retain non-zero
   rows.
8. For retained rows:
   - If `Delivered Qty` is not zero, enter `Delivered`.
   - If `Delivered Qty` is zero and `Material Avail.Date` is blank or zero,
     enter `TBD`.
   - Otherwise, use `=BH2+MOD(5-WEEKDAY(BH2,2),7)`, adjusted for each row.
9. For rows where `Sched Line # = 0`, highlight columns `A:BX` light green
   when `ShipTo Ctry` is `HK`, and light blue otherwise.
10. Refresh PivotTable caches after normalizing Plant values.

## System confirmation validation

Confirm all of the following after saving:

- Both OOR workbooks have `Date to release` in `BX1`.
- Every populated Plant value is numeric and blank Plant cells were filled
  from above.
- Mixed schedule-line groups have blank Date-to-release values on their zero
  schedule rows.
- Retained rows contain `Delivered`, `TBD`, or the expected row formula.
- Zero schedule rows use the expected HK/non-HK highlight color.
