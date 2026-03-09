"""
02_generate_large.py
Generate one xlsx file per sheet (1M rows each) for the large-scale benchmark.

Output:
  data/s1_1m.xlsx  ... data/s5_1m.xlsx   (~73 MB each, ~365 MB total)

Each file takes ~70s on a typical machine.
Run this once; files are skipped if they already exist.
"""

import os
import random
import string
import time
import xlsxwriter

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
os.makedirs(DATA_DIR, exist_ok=True)

ROWS = 1_000_000
COLS = 10


def generate_sheet(sheet_num: int):
    path = os.path.join(DATA_DIR, f"s{sheet_num}_1m.xlsx")
    if os.path.exists(path):
        mb = os.path.getsize(path) / 1024 / 1024
        print(f"  s{sheet_num}_1m.xlsx already exists ({mb:.1f} MB), skipping")
        return

    t0 = time.time()
    wb = xlsxwriter.Workbook(path, {"constant_memory": True})
    ws = wb.add_worksheet(f"Sheet{sheet_num}")

    for i in range(COLS):
        ws.write(0, i, f"col_{i}")

    for r in range(1, ROWS + 1):
        ws.write_row(r, 0, [
            random.randint(0, 999_999),
            round(random.uniform(0, 9_999), 2),
            "".join(random.choices(string.ascii_letters, k=12)),
            r % 2,
            random.randint(0, 100),
            round(random.uniform(0, 500), 3),
            str(random.randint(10_000_000, 99_999_999)),
            random.randint(0, 9_999),
            round(random.uniform(0, 1), 5),
            r % 3,
        ])

    wb.close()
    mb = os.path.getsize(path) / 1024 / 1024
    print(f"  s{sheet_num}_1m.xlsx: {ROWS:,} rows => {mb:.1f} MB  ({time.time() - t0:.0f}s)")


if __name__ == "__main__":
    print(f"Generating {5} x {ROWS:,}-row xlsx files...\n")
    for i in range(1, 6):
        generate_sheet(i)
    print("\nDone.")
