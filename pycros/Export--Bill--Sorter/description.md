> Populates and updates Export Bill from weekly charts + Trade Card
> [!info]
> [Version 0.1.0](#bae1ffff)
>
> [Author](#bae1ffff)
> Nik Faris Aiman bin Nik Rahiman (faris@sig.com.my)
>
> [Last updated date](#bae1ffff)
> 2026/01/27

Takes **VN Weekly Export Chart**, **Local Weekly Export Chart**, **Export Bill**, **Trade Card**, and **Foreign Exchange Administrative Control Chart** Excel files, plus **Year/Month/Week** input, then:
- Inserts new invoices into `Export Bill.xlsx` (NK / Patagonia / NK Local Export)
- Updates value dates, payment refs, and lead-time formulas from `Trade Card.xlsx`
- Regroups matched invoices by Value Date (col E) and sorts by FEAC Ref. No. order (col J)
