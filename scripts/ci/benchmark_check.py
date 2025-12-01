#!/usr/bin/env python3
import json
import sys
from pathlib import Path

def find_benchmark(bm, name):
    for item in bm.get('benchmarks', []):
        if item.get('name') == name:
            return item
    return None

def get_p95(item):
    p95 = item['stats'].get('p95')
    if p95 is None:
        data = item['stats'].get('data', [])
        if not data:
            return None
        data_sorted = sorted(data)
        idx = int(len(data_sorted) * 0.95) - 1
        idx = max(0, min(idx, len(data_sorted) - 1))
        return data_sorted[idx]
    return p95

def get_mean(item):
    return item['stats'].get('mean')

def main():
    if len(sys.argv) < 2:
        print('Usage: benchmark_check.py <benchmark-json> [--gen-threshold N] [--eval-threshold N]')
        sys.exit(2)
    path = Path(sys.argv[1])
    if not path.exists():
        print(f'Missing file {path}')
        sys.exit(2)

    gen_threshold = 2.0
    eval_threshold = 0.01
    # parse additional args
    for i, a in enumerate(sys.argv[2:], start=2):
        if a == '--gen-threshold':
            gen_threshold = float(sys.argv[i+1])
        if a == '--eval-threshold':
            eval_threshold = float(sys.argv[i+1])

    with open(path, 'r') as f:
        data = json.load(f)

    gen = find_benchmark(data, 'test_generator_benchmark')
    evalb = find_benchmark(data, 'test_safe_eval_benchmark')
    exit_code = 0
    if gen is None:
        print('Generator benchmark not found')
    else:
        g_p95 = get_p95(gen)
        print(f'Generator benchmark p95 = {g_p95}s (threshold {gen_threshold}s)')
        if g_p95 is None or g_p95 > gen_threshold:
            print('Generator p95 too high!')
            exit_code = 1

    if evalb is None:
        print('Eval benchmark not found')
    else:
        e_mean = get_mean(evalb)
        print(f'Eval benchmark mean = {e_mean}s (threshold {eval_threshold}s)')
        if e_mean is None or e_mean > eval_threshold:
            print('Safe Eval mean too high!')
            exit_code = 1

    sys.exit(exit_code)

if __name__ == '__main__':
    main()
