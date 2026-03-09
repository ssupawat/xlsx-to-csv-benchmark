"""
03_bench_speed.py
Benchmark XLSX → CSV conversion speed across:
  - python-calamine
  - xlsx2csv
  - libreoffice (headless)

Metrics: mean time (s), stdev, MB/s  over N runs per file.
"""

import gc
import os
import shutil
import statistics
import subprocess
import tempfile
import time
from pathlib import Path

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")


# ── converters ────────────────────────────────────────────────────────────────

def conv_calamine(xlsx_path: str, out_path: str) -> int:
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(xlsx_path)
    sheet = wb.get_sheet_by_index(0)
    rows = 0
    with open(out_path, "w", encoding="utf-8") as f:
        for row in sheet.to_python(skip_empty_area=False):
            f.write(",".join("" if v is None else str(v) for v in row) + "\n")
            rows += 1
    return rows


def conv_xlsx2csv(xlsx_path: str, out_path: str) -> int:
    from xlsx2csv import Xlsx2csv
    Xlsx2csv(xlsx_path, outputencoding="utf-8").convert(out_path)
    with open(out_path) as f:
        return sum(1 for _ in f)


def conv_libreoffice(xlsx_path: str, out_path: str) -> int:
    out_dir = str(Path(out_path).parent)
    subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "csv", "--outdir", out_dir, xlsx_path],
        capture_output=True, text=True, timeout=300,
    )
    expected = Path(out_dir) / (Path(xlsx_path).stem + ".csv")
    if expected.exists() and str(expected) != out_path:
        shutil.move(str(expected), out_path)
    with open(out_path) as f:
        return sum(1 for _ in f)


# ── helpers ───────────────────────────────────────────────────────────────────

def measure(fn, runs: int = 3):
    """Return (mean_s, stdev_s, last_row_count)"""
    times, rows = [], 0
    for _ in range(runs):
        gc.collect()
        t0 = time.perf_counter()
        rows = fn()
        times.append(time.perf_counter() - t0)
    return statistics.mean(times), (statistics.stdev(times) if runs > 1 else 0.0), rows


# ── config ────────────────────────────────────────────────────────────────────

FILES = {
    "small  (1 K rows, ~0.2 MB)":  "small.xlsx",
    "medium (50 K rows, ~9.5 MB)": "medium.xlsx",
    "large  (200 K rows, ~38 MB)": "large.xlsx",
}

TOOLS = [
    ("python-calamine", conv_calamine),
    ("xlsx2csv",        conv_xlsx2csv),
    ("libreoffice",     conv_libreoffice),
]

RUNS = 3


# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 72)
    print(f"  XLSX → CSV  Speed Benchmark  (×{RUNS} runs, mean ± stdev)")
    print("=" * 72)

    summary: dict[str, list[tuple[str, float | None]]] = {}

    for label, fname in FILES.items():
        xlsx_path = os.path.join(DATA_DIR, fname)
        if not os.path.exists(xlsx_path):
            print(f"\n  SKIP {label}: {xlsx_path} not found")
            continue

        file_mb = os.path.getsize(xlsx_path) / 1024 / 1024
        print(f"\n📄 {label}")
        print(f"  {'Tool':<20} {'Mean (s)':>10}  {'Stdev':>8}  {'Rows':>8}  {'MB/s':>8}")
        print(f"  {'-'*20} {'-'*10}  {'-'*8}  {'-'*8}  {'-'*8}")

        for tool_name, fn in TOOLS:
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tf:
                out = tf.name
            try:
                mean, std, rows = measure(lambda p=xlsx_path, o=out: fn(p, o), RUNS)
                mbps = file_mb / mean if mean > 0 else 0
                print(f"  {tool_name:<20} {mean:>10.3f}  {std:>8.3f}  {rows:>8,}  {mbps:>8.2f}")
                summary.setdefault(tool_name, []).append((label.strip(), mean))
            except Exception as e:
                print(f"  {tool_name:<20}  ERROR: {e}")
                summary.setdefault(tool_name, []).append((label.strip(), None))
            finally:
                try:
                    os.unlink(out)
                except OSError:
                    pass

    # speedup table
    print("\n" + "=" * 72)
    print("  SUMMARY — speedup vs slowest (per file)")
    print("=" * 72)
    for label, fname in FILES.items():
        short = label.strip().split()[0]
        times = {}
        for tool_name, _ in TOOLS:
            for lbl, t in summary.get(tool_name, []):
                if lbl.startswith(short) and t is not None:
                    times[tool_name] = t
        if not times:
            continue
        slowest = max(times.values())
        print(f"\n  {label}")
        for tool, t in sorted(times.items(), key=lambda x: x[1]):
            print(f"    {tool:<20} {t:>7.3f}s  {slowest / t:>5.1f}×")
