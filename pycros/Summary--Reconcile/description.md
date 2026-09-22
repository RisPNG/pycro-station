> Build a shipment reconciliation using confirmed forecast, actual-shipment, duplicate, and source-precedence rules.
> [!info]
> [Version 1.3.1](#bae1ffff)
>
> [Author](#bae1ffff)
> OpenAI
>
> [Last updated date](#bae1ffff)
> 2026/09/22

Creates an Excel reconciliation from the Order Control workbook, Shipment Forecast, Weekly Export Local, and Weekly Export VN files.

Includes the previous month's BDS records for quantity-based movement checks and reports through fiscal April. Later-month remarks remain blank except for confirmed shipment movements. Include the previous month's GAC records in the same Order Control export; no additional input file is required.

Delayed movements carry the source month's balancing quantity and amount into the later Ann sheet. Early movements split the later BDS row into the moved balance and an unmarked remainder, preserving the original totals.

The starting month is still the earliest usable PLAN EX-FTY month in the forecast, not the export date or filename. A forecast containing August shipments therefore starts the reconciliation in August, even if exported in September. Unmatched quantities remain Missing for manual review; differing amounts alone do not prevent a movement match.
