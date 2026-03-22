# Drone Loss Parser

A single-file Python script that parses Ukrainian combat drone loss reports (`.docx`) and consolidates all records into a single `.xlsx` output file.

## What it does

1. Scans the current folder for all `.docx` files
2. Parses each file as a Ukrainian "ПОЗАТЕРМІНОВЕ БОЙОВЕ ДОНЕСЕННЯ" (Extraordinary Combat Report)
3. Writes all parsed records into a single `.xlsx` file sorted by date and time

## Usage (end users)

1. Place `parse_drone_losses.exe` into the folder containing the `.docx` report files
2. Double-click `parse_drone_losses.exe` (or run it from the command prompt)
3. The console window shows progress and results
4. The output `.xlsx` file appears in the same folder

The output file is named `ЗВІТ_втрачені_БпЛА_по_DD_MM_YYYY.xlsx`, where the date is the latest loss date found across all parsed files.

> The console window stays open after completion — press **Enter** to close it.

## Output XLSX structure

| # | Column | Description |
|---|---|---|
| 1 | № | Row number |
| 2 | Назва БпЛА | Drone model |
| 3 | Час | Time of loss |
| 4 | Дата | Date of loss |
| 5 | Місце втрати (КООРДИНАТИ) | MGRS/UTM coordinates |
| 6 | Частота керування, МГц | Control frequency (GHz) |
| 7 | Частота відеоканалу, МГц | Video frequency (GHz) |
| 8 | Що відбувалось з БпЛА під час подавлення | What happened during suppression |
| 9 | НАЯВНІСТЬ ЗАЯВКИ НА ПРОЛЬОТ | Flight clearance filed |
| 10 | Причина втрати | Loss reason (from report) |
| 11 | Серійний номер | Serial number |
| 12 | Військова частина / підрозділ | Military unit |

Columns 11 and 12 are additions not present in the standard template.

### Frequency lookup

Frequency data is not in the `.docx` files — it is determined by drone model:

| Drone type | Control | Video |
|---|---|---|
| Вампір / HeavyShot | 2.4 GHz | 2.4 GHz |
| All other models | 2.4 GHz | 5.8 GHz |

## Console output example

```
============================================================
   Парсер втрат БпЛА — старт
============================================================
[INFO] Знайдено файлів .docx: 4

[INFO] Обробка: Донесення_2АДн_Mavic_4Pro_Сафарі_487E.docx
[OK]   DJI Mavic 4 Pro | 21.03.2026 09:37 | 37U CP 50517 46538

[INFO] Обробка: Донесення_2АДн_Autel_4_T_Сафарі_3657.docx
[OK]   Autel Evo Max 4Т | 21.03.2026 10:06 | 37U CP 45225 47782

============================================================
   РЕЗУЛЬТАТ
============================================================
   Всього файлів:            4
   Успішно оброблено:        4
   Не вдалося розпізнати:    0

   Файл збережено: ЗВІТ_втрачені_БпЛА_по_21_03_2026.xlsx
============================================================
```

## Development

### Prerequisites

```
python-docx   # reading .docx files
openpyxl      # writing .xlsx files
pyinstaller   # building the standalone executable
```

Install with [Poetry](https://python-poetry.org/):

```bash
poetry install
```

### Run locally

```bash
make run
# or
poetry run python parse_drone_losses.py
```

### Build standalone executable

**macOS:**
```bash
make build-mac
# output: dist/parse_drone_losses
```

**Windows** (run on a Windows machine):
```bash
make build-win
# output: dist/parse_drone_losses.exe
```

Or manually:
```bash
poetry run pyinstaller --onefile --console parse_drone_losses.py
```

The resulting `dist/parse_drone_losses.exe` is the single file to distribute — no Python installation required on the end user's machine.

### Clean build artifacts

```bash
make clean
```

## Error handling

- **Invalid `.docx`** — logged as a warning, file is skipped
- **Missing mandatory fields** (drone model or date) — logged with the filename and missing field, record is skipped
- **Output file open in Excel** — `PermissionError` is caught with a clear message to close the file
- **No `.docx` files found** — prints a clear message and exits gracefully