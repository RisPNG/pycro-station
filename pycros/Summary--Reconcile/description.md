> Build a shipment reconciliation using confirmed forecast, actual-shipment, duplicate, and source-precedence rules.
> [!info]
> [Version 1.3.5](#bae1ffff)
>
> [Author](#bae1ffff)
> SIG4475
>
> [Last updated date](#bae1ffff)
> 2026/11/04

Creates an Excel reconciliation from the Order Control workbook, Shipment Forecast, Weekly Export Local, and Weekly Export VN files.

Weekly actuals include descriptive sample labels as separate Salesman Sample rows. An amount-only DN entry appears with zero quantity and the DN remark. Actual shipment amounts take precedence over planned amounts.

BDS columns are detected by header names in both single-row and stacked-header exports. Missing or ambiguous required columns stop processing with an error instead of silently producing zero totals. Where Job Number is repeated, the rightmost Job Number column is used, matching the original base-job export field.

Includes the previous month's BDS records for quantity-based movement checks and reports through fiscal April. Later-month remarks remain blank except for confirmed shipment movements. Include the previous month's GAC records in the same Order Control export; no additional input file is required.

Delayed movements carry the source month's balancing quantity and amount into the later Ann sheet. Partial early movements split the later BDS row into the moved balance and an unmarked remainder. Full-quantity early movements keep one marked row with the original BDS amount, without a separate amount-only remainder. BDS totals are preserved.

The reporting month comes from the Order Control workbook's Date Prepared in pcp2012!A1, using DD/MM/YYYY. A BDS prepared on 3 November 2026 starts the reconciliation in November and uses November and December for Ann, even when the Shipment Forecast retains October shipments. October BDS records and weekly actuals supply movement context only. The report continues through fiscal April. A missing or invalid prepared date stops processing with an error. Unmatched quantities remain Missing for manual review; differing amounts alone do not prevent a movement match.
