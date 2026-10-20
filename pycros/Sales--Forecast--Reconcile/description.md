> Reconcile weekly Sales Forecast changes and update the variance report.
> [!info]
> [Version 0.1.0](#bae1ffff)
>
> [Last updated date](#bae1ffff)
> 2026/10/20

Select the Sales Report, FOTP by season, Shipment Forecast from Ann, and previous-week BDS workbooks. The program compares each month's New and Old entries, creates the Recon sheets, and updates the variance sheet in a new workbook.

The Shipment Forecast supplies shortship/TBA matches, and the previous-week BDS identifies new forecast POs. Recon columns B:C are New and D:E are Old.

Matching FOTP job and total PO quantity determines the GAC code for delay shipments. When the quantity does not match or matching rows have different GAC codes, the code remains blank. SIG Finance Comment cells use formulas against the Sales Report's Reason code sheet so added mappings can resolve when the workbook is recalculated.
