> Combine VTEC payment summary workbooks into monthly worksheets with VCB USD rate lookups.
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
> 2026/03/26

Choose one or more input folders that contain `.xlsx` VTEC payment summaries, choose one or more VCB exchange-rate workbooks, and optionally choose an output folder.

The pycro will:
- combine all valid rows into one workbook
- split the output into sheets such as `JAN'25`, `FEB'25`, and `DEC'25` based on `Payment to Supplier`
- add forex calculation columns using VCB USD Telegraphic Buying rates
- write an audit `.txt` log beside the workbook for skipped files, missing headers, missing mandatory row dates, and other issues
