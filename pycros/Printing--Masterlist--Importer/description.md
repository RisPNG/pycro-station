> Import BDD strike off request rows into the print strike off tracker.
> [!info]
> [Version 1.0.0](#bae1ffff)
>
> [Author](#bae1ffff)
> Nik Faris Aiman bin Nik Rahiman (faris@sig.com.my)
>
> [Requested by](#bae1ffff)
> Faris
>
> [Latest maintained by](#bae1ffff)
> Nik Faris Aiman bin Nik Rahiman (faris@sig.com.my)
>
> [Last updated date](#bae1ffff)
> 2026/05/08

Select one or more BDD Strike Off Request workbooks and one Print Strike Off Tracker workbook.

The pycro imports rows where `STYLE NO.` is filled, maps matching headers from A into B, and starts writing at the first B row where `STYLE NO.` is empty. In-cell `LOGO ARTWORK` images are transferred without using openpyxl save, so Excel richData images stay in the output workbook.
