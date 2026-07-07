> Builds a diverted purchase order size reconciliation workbook from Nike PO search results and original PO PDFs.
> [!info]
> [Version 0.1.1](#bae1ffff)
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
> 2026/07/07

Reads a PO search-results Excel file and original Nike purchase order PDF files, detects diverted PO line items, reconciles NEW / ORI / NOW size quantities, and outputs an Excel workbook in the same layout style as the divert construct examples.

The processor avoids OCR for normal Nike PO PDFs. It extracts selectable PDF text, validates item totals where possible, maps target line-item suffixes to actual target size rows, caps allocations against both target size quantities and source ORI-minus-NOW quantities to ignore stale Item Text entries, and strips unused size columns from the final workbook so only utilized sizes remain.
