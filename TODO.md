# Open Questions / TODOs

Items are listed in priority order (highest first).

---

## 1. Output file naming convention

**Current pattern:** `ЗВІТ_втрачені_БпЛА_по_DD_MM_YYYY.xlsx`
The date is derived from the **latest** loss date found across all parsed records.

**Questions:**
- Should the date reflect the latest loss date, the run date, or a date range (e.g. `з_DD_MM_YYYY_по_DD_MM_YYYY`)?
- The example output file on disk is named `ЗВІТ_втрачені_БпЛА_по_21_03_2026.xlsx` — does this match the expected convention?
- The code comment mentions a possible brigade identifier in the name (e.g. `3_БрОП_НГУ`). Should this be added once reports are standardized?

---

## 2. Input file discovery / naming convention

**Current behavior:** the script reads **all** `.docx` files from the folder where the executable is placed (skipping `~$` temp files). There is no filename filter.

**Questions:**
- Is a standard filename pattern expected for input reports (e.g. `Позатермінове_бойове_донесення_*.docx`)? If yes, the glob filter should be tightened so unrelated `.docx` files are ignored.
- Should all input `.docx` files follow a single, unified template/format? Currently the parser handles minor variations between documents, but a enforced standard format would make parsing more reliable and remove the need for fallback logic.
- Should the script also recurse into sub-folders, or stay flat?

---

## 3. Correct parsing of all document types, including frequency tables

**Current behavior:** frequencies are **not** parsed from the `.docx` content. They are assigned from a hard-coded lookup table:
- Vampire / HeavyShot models → control 2.4 GHz / video 2.4 GHz
- All other models → control 2.4 GHz / video 5.8 GHz

**Questions:**
- Do the `.docx` reports actually contain a frequency table? If so, should the parser extract real values from it instead of using the lookup table?
- Are there drone models that fall outside the two categories above (different frequency pairs)?
- Are there other table-based fields in the documents that are currently missed by the paragraph-only parser?

---

## 4. Sorting of the output file

**Current behavior:** records are sorted by **loss date, then loss time** (earliest first).

**Questions:**
- Is date+time the correct sort key?
- Should a secondary sort by unit (`Військова частина`) or drone model be applied?
- Should newest records appear first instead?

---

## 5. Output file styling (font / size / family)

**Current styling:**
- Header row: `Calibri 11 Bold`, fill `#BDD7EE`, wrap + center aligned, row height 60pt
- Data rows: `Calibri` (default size, no explicit override), wrap + center vertical, thin borders
- No explicit font/size set on data cells — relies on workbook default

**Questions:**
- Is `Calibri 11` the correct font for both headers and data cells?
- Should data rows have an explicit font size set (currently inherited from workbook default)?
- Is the header row height of 60pt correct for the expected column header text?
- Is the blue fill (`#BDD7EE`) the right header color?

---

## 6. Column names and widths

**Current columns (in order):**

| # | Header | Width (chars) |
|---|--------|---------------|
| 1 | № | 5 |
| 2 | Назва БпЛА | 22 |
| 3 | Час | 8 |
| 4 | Дата | 13 |
| 5 | Місце втрати (КООРДИНАТИ) | 30 |
| 6 | Частота керування, МГц | 12 |
| 7 | Частота відеоканалу, МГц | 12 |
| 8 | Що відбувалось з БпЛА під час подавлення | 38 |
| 9 | НАЯВНІСТЬ ЗАЯВКИ НА ПРОЛЬОТ (перед заповненням поля розібратись) | 18 |
| 10 | Причина втрати (Тех. Проблеми / Погодні умови, Збито, Ворожий РЕБ, Дружній РЕБ) | 48 |
| 11 | Серійний номер | 24 |
| 12 | Військова частина / підрозділ | 32 |

**Questions:**
- Are all 12 columns present and in the correct order relative to the reference report?
- Are any columns missing or extra?
- Are the column widths wide enough for typical content, or do any need adjustment?
- Should frequency columns (6, 7) show `МГц` or `ГГц`? (Values are currently stored as GHz floats — e.g. `2.4` — but headers say `МГц`.)