# xlsx-to-csv-benchmark

Benchmark comparing **python-calamine**, **xlsx2csv**, and **LibreOffice headless** for XLSX → CSV conversion in Python.

## Results Summary

### Speed (mean over 3 runs)

| Tool | 1K rows | 50K rows | 200K rows | 5×200K rows |
|---|---|---|---|---|
| **python-calamine** | **0.04s** 🥇 | **1.61s** 🥇 | **6.33s** 🥇 | **13.9s** 🥇 |
| xlsx2csv | 0.12s | 5.15s | 20.1s | 38.9s |
| libreoffice | 3.61s | 10.87s | 35.3s | 81.6s ⚠️ |

### Memory (peak RSS)

| Tool | 1K rows | 50K rows | 200K rows | 5×200K rows |
|---|---|---|---|---|
| **python-calamine** | 16 MB | 98 MB | 345 MB | 245 MB |
| **xlsx2csv** | 22 MB | **21 MB** 🥇 | **21 MB** 🥇 | **21 MB** 🥇 |
| libreoffice | 288 MB | 431 MB | 860 MB | 1,264 MB |

> **calamine** memory scales with file size (~4× of xlsx size).  
> **xlsx2csv** uses streaming SAX parse → ~21 MB constant regardless of file size.

### Formula Handling

All three tools read **cached values only** — none evaluate formulas at runtime.

| Scenario | calamine | xlsx2csv | openpyxl `data_only=True` |
|---|---|---|---|
| Formula + cached value (Excel / xlsxwriter) | ✅ | ✅ | ✅ |
| Formula without cached value (openpyxl save) | ❌ empty | ❌ empty | ❌ None |
| Live formula evaluation | ❌ | ❌ | ❌ |

For live evaluation: use **LibreOffice** or the [`formulas`](https://github.com/vinci1it2000/formulas) library.

---

## Decision Guide

| Scenario | Recommended |
|---|---|
| Small–medium files, speed priority | **python-calamine** |
| Lambda ≤ 256 MB or very large files | **xlsx2csv** |
| 1M rows × 5 sheets on constrained RAM | **xlsx2csv** (calamine needs ~1.7 GB) |
| Multi-sheet export | **Do NOT use LibreOffice** (exports first sheet only) |
| Files with uncached formulas | **LibreOffice** or `formulas` library |

---

## Setup

```bash
pip install -r requirements.txt

# LibreOffice (Ubuntu/Debian)
sudo apt install libreoffice-calc libreoffice-core

# LibreOffice (macOS)
brew install --cask libreoffice
```

## Quick Start

Run all benchmarks at once:

```bash
# Local (requires LibreOffice installed)
./run-all.sh

# Docker (includes LibreOffice)
docker build -t xlsx-benchmark .
docker run --rm xlsx-benchmark
```

## Individual Benchmarks

```bash
# 1. Generate test data
python scripts/01_generate_data.py

# 2. (Optional) Generate 1M-row per-sheet files (~70s each)
python scripts/02_generate_large.py

# 3. Speed benchmark
python scripts/03_bench_speed.py

# 4. Memory benchmark (Linux only — uses /proc/<pid>/status)
python scripts/04_bench_memory.py

# 5. Multi-sheet benchmark
python scripts/05_bench_multisheet.py

# 6. Formula handling test
python scripts/06_test_formulas.py
```

## Notes

- Memory benchmark uses `psutil` for cross-platform RSS measurement (Linux, macOS, Windows).
- LibreOffice `--convert-to csv` exports **first sheet only**. Full multi-sheet export requires `python-uno` or per-sheet invocation.
- The 1M × 5 sheets scenario was extrapolated from single-sheet measurements; generating 5 × 1M rows takes ~6 minutes on a standard machine.
- `python-calamine` loads the full sheet into memory before streaming — memory scales linearly with file size.
- `xlsx2csv` parses XML via SAX — memory stays constant (~21 MB) regardless of file size.

## Environment

Tested on Ubuntu 24.04, Python 3.12, LibreOffice 24.2.7.
