> Combine VTEC payment summary workbooks into monthly worksheets with S Rate matching.
> [!info]
> [Version 1.0.0](#bae1ffff)
>
> [Author](#bae1ffff)
> OpenAI Codex
>
> [Requested by](#bae1ffff)
> Workspace user
>
> [Latest maintained by](#bae1ffff)
> OpenAI Codex
>
> [Last updated date](#bae1ffff)
> 2026/03/27

Add one or more input folders that contain `.xlsx` VTEC payment summaries, choose one or more S Rate workbooks, and optionally choose an output folder. Use `Add Input Folder` again to include more folders. The selected folders are scanned recursively, so files inside subfolders are included automatically.

The pycro will:
- combine all valid rows into one workbook
- split the output into sheets such as `JAN'25`, `FEB'25`, and `DEC'25` based on `Payment to Supplier`
- insert `Standard Rate` and `Status` after `Currency Rate`
- read `Material Delivery -> From/To` date ranges plus `S Rate` from the selected S Rate workbook(s)
- match `S Rate` using `VAT Invoice Date`, then mark `Status` against `Currency Rate`
- write an audit `.txt` log beside the workbook for skipped multi-sheet files, missing required headers, missing mandatory dates, invalid S Rate rows, and dates without matching S Rate ranges
