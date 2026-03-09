"""
04_bench_memory.py
Benchmark peak RSS memory usage during XLSX → CSV conversion.

Uses /proc/<pid>/status (Linux only) to sample RSS of child processes,
giving accurate per-tool memory profiles including native/Rust extensions.

Outputs:
  - Peak RSS (MB)
  - Peak RSS / file size ratio
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


# ── converter scripts (run as subprocesses) ───────────────────────────────────

CALAMINE_SCRIPT = """
import sys
from python_calamine import CalamineWorkbook
xlsx, out = sys.argv[1], sys.argv[2]
wb = CalamineWorkbook.from_path(xlsx)
sheet = wb.get_sheet_by_index(0)
with open(out, "w") as f:
    for row in sheet.to_python(skip_empty_area=False):
        f.write(",".join("" if v is None else str(v) for v in row) + "\\n")
"""

XLSX2CSV_SCRIPT = """
import sys
from xlsx2csv import Xlsx2csv
xlsx, out = sys.argv[1], sys.argv[2]
Xlsx2csv(xlsx, outputencoding="utf-8").convert(out)
"""


def run_python_tool(script: str, xlsx_path: str, out_path: str) -> float:
    proc = subprocess.Popen(
        ["python3", "-c", script, xlsx_path, out_path],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    return sample_rss_until_done(proc)


def run_libreoffice(xlsx_path: str, out_path: str) -> float:
    out_dir = str(Path(out_path).parent)
    proc = subprocess.Popen(
        ["libreoffice", "--headless", "--convert-to", "csv", "--outdir", out_dir, xlsx_path],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    peak = sample_rss_until_done(proc)
    expected = Path(out_dir) / (Path(xlsx_path).stem + ".csv")
    if expected.exists() and str(expected) != out_path:
        shutil.move(str(expected), out_path)
    return peak


# ── config ────────────────────────────────────────────────────────────────────

FILES = {
    "small  (1 K rows, ~0.2 MB)":  "small.xlsx",
    "medium (50 K rows, ~9.5 MB)": "medium.xlsx",
    "large  (200 K rows, ~38 MB)": "large.xlsx",
}


# ── main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("=" * 62)
    print("  XLSX → CSV  RSS Memory Benchmark  (actual process memory)")
    print("=" * 62)

    for label, fname in FILES.items():
        xlsx_path = os.path.join(DATA_DIR, fname)
        if not os.path.exists(xlsx_path):
            print(f"\n  SKIP {label}: file not found")
            continue

        file_mb = os.path.getsize(xlsx_path) / 1024 / 1024
        print(f"\n📄 {label}")
        print(f"  {'Tool':<20} {'Peak RSS MB':>12}  {'Peak/File':>10}")
        print(f"  {'-'*20} {'-'*12}  {'-'*10}")

        for tool_name, script in [("python-calamine", CALAMINE_SCRIPT),
                                    ("xlsx2csv",        XLSX2CSV_SCRIPT)]:
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tf:
                out = tf.name
            try:
                peak = run_python_tool(script, xlsx_path, out)
                print(f"  {tool_name:<20} {peak:>12.1f}  {peak / file_mb:>9.1f}×")
            finally:
                try:
                    os.unlink(out)
                except OSError:
                    pass

        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tf:
            out = tf.name
        try:
            peak = run_libreoffice(xlsx_path, out)
            print(f"  {'libreoffice':<20} {peak:>12.1f}  {peak / file_mb:>9.1f}×")
        finally:
            try:
                os.unlink(out)
            except OSError:
                pass
