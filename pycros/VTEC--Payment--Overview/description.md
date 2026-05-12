> Generate VTEC Payment Overview and duplicate records from VTEC payment Excel files.
> [!info]
> [Version 1.0.0](#bae1ffff)
>
> [Author](#bae1ffff)
> Nik Faris Aiman bin Nik Rahiman (faris@sig.com.my)
>
> [Requested by](#bae1ffff)
> Converted from VTEC Payment Overview & Duplicates VBA Macro
>
> [Latest maintained by](#bae1ffff)
> Nik Faris Aiman bin Nik Rahiman (faris@sig.com.my)
>
> [Last updated date](#bae1ffff)
> 2026/05/12

Converts the legacy VTEC payment overview VBA workflow into a Pycro.

This Pycro reads one or more VTEC payment workbooks, extracts rows from worksheets whose names contain `payment`, creates a fresh timestamped `.xlsx` workbook for the current session, records processed file/sheet pairs in a `Processing Log` sheet, and moves later duplicate records into `VTEC Payment Duplicates` while keeping the earliest `Payment To VTEC (LSKhor)` record in `VTEC Payment Overview`.

Each run creates a new workbook beside the first selected payment file. The output workbook always contains `VTEC Payment Overview`, `VTEC Payment Duplicates`, and `Processing Log`.
