> Build a BSD-versus-Ann shipment reconciliation across the full month range detected from the Ann Forecast.
> [!info]
> [Version 1.1.4](#bae1ffff)
>
> [Author](#bae1ffff)
> OpenAI
>
> [Last updated date](#bae1ffff)
> 2026/07/17

Creates an Excel reconciliation from the BSD Order Control workbook, Shipment Forecast, Weekly Export Local, and Weekly Export VN files.

The Pycro automatically detects the earliest and latest usable PLAN EX-FTY month in the Ann Forecast SHIPMENTS sheet and reconciles every calendar month in that range. It validates the four layouts, aggregates BDS and Ann values by normalised job number, replaces the first detected month's forecasts with weekly actual shipments, detects adjacent-month early or delayed shipments, and creates a fiscal summary containing only the detected reconciliation months plus one reconciliation sheet per detected month.

There is no report-date or FX input in the Pycro interface. Every generated workbook is named `Summary_Reconcile_yyyymmdd_hhmmss.xlsx` using its creation timestamp.

The fiscal summary always reserves a highlighted `Fx Adjustment` row directly below `Price Discrepancy`. Every detected month starts at `0.00`. The row is summed normally with every other variance reason. Users should enter the signed amount shown in the approved reconciliation; for example, a negative FX value increases Ann and reduces BDS minus Ann. The row is included in the monthly variance totals, Ann Report totals, BDS-minus-Ann totals, and grand totals.
