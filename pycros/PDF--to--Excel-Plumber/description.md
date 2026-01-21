> Convert PDF documents to structured Excel files
> [!info]
> [Version 1.0.0](#bae1ffff)
>
> [Author](#bae1ffff)
> NIK FARIS AIMAN BIN NIK RAHIMAN (faris@sig.com.my)
>
> [Last updated date](#bae1ffff)
> 2026/01/21

Extracts tables and structured data from text-based PDF documents and converts them to Excel format.

Optimized for systems with integrated graphics (Intel Iris Xe, AMD APU) or limited RAM.

**Primary method:** Uses pdfplumber for fast, direct text/table extraction from text-based PDFs. Works instantly on most PDFs.

**Fallback method:** marker-pdf on CPU for complex layouts or image-heavy PDFs. This is slower but handles edge cases.

Specifically designed for payment settlement PDFs with reference numbers, transaction details, and fee schedules.
