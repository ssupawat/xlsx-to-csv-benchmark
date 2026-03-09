#!/bin/bash
# Run all XLSX to CSV benchmarks

set -e

echo "========================================================================"
echo "  XLSX → CSV Benchmark Suite"
echo "========================================================================"

# Detect LibreOffice command (macOS uses 'soffice', Linux uses 'libreoffice')
if command -v libreoffice &> /dev/null; then
    export LIBREOFFICE_CMD=libreoffice
elif command -v soffice &> /dev/null; then
    export LIBREOFFICE_CMD=soffice
else
    echo "WARNING: LibreOffice not found. Skipping LibreOffice benchmarks."
    export LIBREOFFICE_CMD=""
fi

echo ""
echo "Step 1: Generating test data..."
python3 scripts/01_generate_data.py

echo ""
echo "========================================================================"
echo "Step 2: Speed Benchmark"
echo "========================================================================"
python3 scripts/03_bench_speed.py

echo ""
echo "========================================================================"
echo "Step 3: Memory Benchmark"
echo "========================================================================"
python3 scripts/04_bench_memory.py

echo ""
echo "========================================================================"
echo "Step 4: Multi-sheet Benchmark"
echo "========================================================================"
python3 scripts/05_bench_multisheet.py

echo ""
echo "========================================================================"
echo "Step 5: Formula Handling Test"
echo "========================================================================"
python3 scripts/06_test_formulas.py

echo ""
echo "========================================================================"
echo "  All benchmarks complete!"
echo "========================================================================"
