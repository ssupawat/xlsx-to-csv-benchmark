# xlsx-to-csv-benchmark

Benchmark comparing **python-calamine**, **xlsx2csv**, and **LibreOffice headless** for XLSX → CSV conversion in Python.

## Results Summary

### Speed (mean over 3 runs)

| Tool | 1K rows | 50K rows | 200K rows | 5×200K rows |
|---|---|---|---|---|
| **python-calamine** | **0.016s** 🥇 | **0.870s** 🥇 | **3.009s** 🥇 | **8.47s** 🥇 |
| libreoffice | 0.883s | 1.973s | 6.892s | 49.25s ⚠️ |
| xlsx2csv | 0.053s | 2.557s | 9.424s | 26.19s |

### Memory (peak RSS)

| Tool | 1K rows | 50K rows | 200K rows | 5×200K rows |
|---|---|---|---|---|
| **python-calamine** | 0.1 MB | 90.9 MB | 335.0 MB | 196 MB |
| **xlsx2csv** | **14.2 MB** 🥇 | **14.2 MB** 🥇 | **14.2 MB** 🥇 | **14 MB** 🥇 |
| libreoffice | 154.1 MB | 295.2 MB | 701.1 MB | 1,153 MB |

> **calamine** memory scales with file size (~9× of xlsx size for large files).
> **xlsx2csv** uses streaming SAX parse → ~14 MB constant regardless of file size.
> **libreoffice** has high startup overhead and memory; multi-sheet exports first sheet only.

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
| 5×200K sheets on constrained RAM | **xlsx2csv** (14 MB constant) |
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
- LibreOffice `--convert-to csv` exports **first sheet only**. Multi-sheet results are projected from single-sheet measurements.
- `python-calamine` loads the full sheet into memory before streaming — memory scales linearly with file size.
- `xlsx2csv` parses XML via SAX — memory stays constant (~14 MB) regardless of file size.

## Environment

**Tested in Podman container:**
- Base: `python:3.12-slim` (Debian Trixie)
- Python: 3.12
- LibreOffice: (Debian package, latest via apt)
- Architecture: linux/arm64

**To reproduce:**
```bash
podman build -t xlsx-benchmark .
podman run --rm xlsx-benchmark
```
