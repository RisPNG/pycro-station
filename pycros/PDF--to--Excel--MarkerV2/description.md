> PDF to Excel with LLM-powered context understanding
> [!info]
> [Version 2.0.0](#bae1ffff)
>
> [Author](#bae1ffff)
> NIK FARIS AIMAN BIN NIK RAHIMAN (faris@sig.com.my)
>
> [Last updated date](#bae1ffff)
> 2026/01/21

Converts ANY PDF to intelligently structured Excel using a 3-step pipeline:

**Pipeline:**
1. **marker-pdf**: Extract content from any PDF (text-based or scanned)
2. **LLM (Ministral-3B or similar)**: Understand document context and decide structure
3. **Output**: Properly organized Excel with meaningful sheets/columns

**LLM Context Understanding:**
- Identifies document type (invoice, payment settlement, report, etc.)
- Determines appropriate sheet structure
- Maps data to meaningful column names
- Handles complex multi-section documents

**Requirements:**
- marker-pdf, torch, pandas, openpyxl (required)
- llama-cpp-python + GGUF model (optional but recommended)

**Recommended model:** Ministral-3B-Instruct-Q4_K_M.gguf (~2GB, runs on CPU)

**Performance (i5 + 16GB RAM):**
- marker-pdf extraction: 2-5 min/page
- LLM analysis: ~1-2 min
- Total for 4-page PDF: ~10 minutes
