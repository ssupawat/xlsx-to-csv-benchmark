"""
05_bench_multisheet.py
Benchmark multi-sheet XLSX → CSV conversion (all sheets exported).

Tests the scenario: 5 sheets × N rows each.

Default uses data/multisheet.xlsx (5 sheets × 200K rows = 1M total).
For the 1M-rows-per-sheet scenario, first run 02_generate_large.py then
point FILES_LARGE at the per-sheet files.

Key finding: LibreOffice --convert-to csv exports ONLY the first/active
sheet. Multi-sheet export requires iterating per sheet (5× spawns).
"""

import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path
import psutil

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")


# ── RSS sampler ───────────────────────────────────────────────────────────────

def sample_rss_until_done(proc: subprocess.Popen, interval: float = 0.05) -> float:
    """Sample RSS memory of process and all children until completion."""
    peaks = []
    while proc.poll() is None:
        try:
            parent = psutil.Process(proc.pid)
            rss = parent.memory_info().rss / 1024 / 1024
            for child in parent.children(recursive=True):
                rss += child.memory_info().rss / 1024 / 1024
        except psutil.NoSuchProcess:
            rss = 0.0
        peaks.append(rss)
        time.sleep(interval)
    return max(peaks) if peaks else 0.0


# ── converter scripts ─────────────────────────────────────────────────────────

CALAMINE_SCRIPT = """
import sys, time
from python_calamine import CalamineWorkbook
xlsx = sys.argv[1]
wb = CalamineWorkbook.from_path(xlsx)
t0 = time.perf_counter()
total = 0
for i in range(len(wb.sheet_names)):
    sheet = wb.get_sheet_by_index(i)
    with open(f"/tmp/cal_sheet_{i}.csv", "w") as f:
        for row in sheet.to_python(skip_empty_area=False):
            f.write(",".join("" if v is None else str(v) for v in row) + "\\n")
            total += 1
print(f"{time.perf_counter()-t0:.3f}s | {total} rows")
"""

XLSX2CSV_SCRIPT = """
import sys, time
from xlsx2csv import Xlsx2csv
xlsx = sys.argv[1]
t0 = time.perf_counter()
x = Xlsx2csv(xlsx, outputencoding="utf-8")
total = 0
for i, name in enumerate(x.workbook.sheets):
    out = f"/tmp/x2c_sheet_{i}.csv"
    x.convert(out, sheetid=i + 1)
    with open(out) as f:
        total += sum(1 for _ in f)
print(f"{time.perf_counter()-t0:.3f}s | {total} rows")
"""


def run_tool(script: str, xlsx_path: str) -> tuple[float, int, float]:
    """Returns (elapsed_s, total_rows, peak_rss_mb)"""
    proc = subprocess.Popen(
        ["python3", "-c", script, xlsx_path],
        stdout=subprocess.PIPE, stderr=subprocess.DEVNULL,
    )
    peak = sample_rss_until_done(proc)
    stdout = proc.stdout.read().decode().strip()
    try:
        parts = stdout.split("|")
        elapsed = float(parts[0].strip().rstrip("s"))
        rows = int(parts[1].strip().split()[0])
    except (IndexError, ValueError):
        elapsed, rows = 0.0, 0
    return elapsed, rows, peak


def run_libreoffice_all_sheets(xlsx_path: str) -> tuple[float, int, float]:
    """
    LO can only export one sheet per invocation.
    We iterate sheet names and spawn once per sheet.
    Returns (total_elapsed_s, total_rows, peak_rss_mb_of_worst_spawn)
    """
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(xlsx_path)
    sheet_names = wb.sheet_names

    total_elapsed = 0.0
    total_rows = 0
    worst_rss = 0.0
    out_dir = tempfile.mkdtemp()

    for i, name in enumerate(sheet_names):
        # LO doesn't support choosing sheet via CLI easily;
        # common workaround: copy sheet to a temp single-sheet file
        # Here we just note the limitation and run once for timing purposes
        t0 = time.time()
        proc = subprocess.Popen(
            ["libreoffice", "--headless", "--convert-to", "csv",
             "--outdir", out_dir, xlsx_path],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        rss = sample_rss_until_done(proc)
        elapsed = time.time() - t0

        worst_rss = max(worst_rss, rss)
        total_elapsed += elapsed

        csvs = list(Path(out_dir).glob("*.csv"))
        for c in csvs:
            with open(c) as f:
                total_rows += sum(1 for _ in f)
            c.unlink()

        # Only run first sheet for timing (subsequent would be same file)
        # In real usage you'd split the workbook first
        print(f"    LO sheet {i+1}/{len(sheet_names)}: {elapsed:.1f}s  RSS {rss:.0f}MB")
        break  # NOTE: remove this break to measure all sheets (very slow)

    shutil.rmtree(out_dir, ignore_errors=True)
    projected_total = total_elapsed * len(sheet_names)
    # Note: LibreOffice processes sheets sequentially, not simultaneously.
    # Peak memory is the worst single sheet's RSS, not the sum of all sheets.
    return projected_total, total_rows, worst_rss


# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    xlsx_path = os.path.join(DATA_DIR, "multisheet.xlsx")
    if not os.path.exists(xlsx_path):
        print(f"ERROR: {xlsx_path} not found. Run 01_generate_data.py first.")
        raise SystemExit(1)

    file_mb = os.path.getsize(xlsx_path) / 1024 / 1024

    print("=" * 68)
    print("  Multi-sheet Benchmark  (5 sheets × 200K rows = 1M rows total)")
    print(f"  File: {os.path.basename(xlsx_path)}  ({file_mb:.1f} MB)")
    print("=" * 68)
    print(f"\n  {'Tool':<20} {'Time (s)':>10}  {'Peak RSS':>10}  {'Total Rows':>12}")
    print(f"  {'-'*20} {'-'*10}  {'-'*10}  {'-'*12}")

    for tool_name, script in [("python-calamine", CALAMINE_SCRIPT),
                                ("xlsx2csv",        XLSX2CSV_SCRIPT)]:
        elapsed, rows, rss = run_tool(script, xlsx_path)
        print(f"  {tool_name:<20} {elapsed:>10.2f}  {rss:>9.0f}MB  {rows:>12,}")

    print("\n  libreoffice — NOTE: exports first sheet only per invocation.")
    print("  Timing below is (1 sheet × 5) projected.")
    elapsed, rows, rss = run_libreoffice_all_sheets(xlsx_path)
    print(f"  {'libreoffice (×5)':<20} {elapsed:>10.2f}  {rss:>9.0f}MB  {rows:>12,} (projected)")
