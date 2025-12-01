# Benchmarks and CI thresholds

This repository uses pytest-benchmark to measure performance for critical code paths (formula evaluation, generator). The CI workflow runs benchmarks and checks the following thresholds:

- `test_generator_benchmark` p95 must be <= 2.0 seconds
- `test_safe_eval_benchmark` mean must be <= 0.01 seconds

If you intentionally change code that affects performance, update the thresholds in the CI benchmark check script at `scripts/ci/benchmark_check.py` accordingly and confirm the new baseline by running the local benchmarks:

```bash
python -m pytest tests/benchmark_formula_eval.py -q --benchmark-json=bench.json
python scripts/ci/benchmark_check.py bench.json --gen-threshold 2.0 --eval-threshold 0.01
```

To update thresholds permanently, edit `scripts/ci/benchmark_check.py` and adjust the `--gen-threshold` and `--eval-threshold` CLI defaults or keep them as CLI options managed as arguments in the GitHub Actions step.
