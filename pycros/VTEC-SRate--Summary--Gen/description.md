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
- create yearly summary sheets such as `SUM'25` with formula-based monthly status breakdowns
- insert `Standard Rate` and `Status` after `Currency Rate`
- read `Material Delivery -> From/To` date ranges plus `S Rate` from the selected S Rate workbook(s)
- match `S Rate` using `VAT Invoice Date`, then mark `Status` against `Currency Rate`
- accept source files as long as they have exactly 1 sheet plus the headers `VAT Invoice Date` and `Currency Rate`
- reject a whole source file when any data row is missing `VAT Invoice Date` or `Currency Rate`, and log the file plus row numbers in the audit `.txt`
- write `Status` as an Excel formula so it updates if `Currency Rate` or `Standard Rate` is edited later
