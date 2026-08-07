> Builds a diverted purchase order size reconciliation workbook from Nike PO search results and original PO PDFs or Excel files.
> [!info]
> [Version 0.2.0](#bae1ffff)
>
> [Author](#bae1ffff)
> Nik Faris Aiman bin Nik Rahiman
>
> [Requested by](#bae1ffff)
> Internal automation request
>
> [Latest maintained by](#bae1ffff)
> Nik Faris Aiman bin Nik Rahiman
>
> [Last updated date](#bae1ffff)
> 2026/08/07

Reads every worksheet in a PO search-results Excel file that contains the required PO search-results headers, reads original Nike purchase order PDF or Excel files, detects diverted PO line items, reconciles NEW / ORI / NOW size quantities, and outputs an Excel workbook in the same layout style as the divert construct examples.

The processor avoids OCR for normal Nike PO PDFs. It extracts selectable PDF text, validates item totals where possible, or reads the purchase order number, PO line, size, and quantity columns from an original PO Excel file. It maps target line-item suffixes to actual target size rows, supports target-side Diverted From text when source rows are not present in the search-results workbook, caps fallback allocations against both target size quantities and source ORI-minus-NOW quantities to ignore stale Item Text entries, and strips unused size columns from the final workbook so only utilized sizes remain.
