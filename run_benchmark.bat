@echo off
chcp 65001 >nul
echo =========================================================================
echo       CHAY BENCHMARK THUC NGHIEM DOC LAP DT005 (5 ITERATIONS)
echo =========================================================================
python -m src.benchmark.run_benchmark --dir realistic_dataset --runs 5 --results-dir results
pause
