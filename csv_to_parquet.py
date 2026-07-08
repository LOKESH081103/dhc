#!/usr/bin/env python3
"""
Convert a CSV file to Parquet.

Usage:
    python csv_to_parquet.py path/to/input.csv
    python csv_to_parquet.py path/to/input.csv path/to/output.parquet

If no output path is given, the Parquet file is written next to the CSV
with a .parquet extension.

Requires: pandas, pyarrow
    pip install pandas pyarrow
"""

import argparse
import sys
from pathlib import Path


def csv_to_parquet(csv_path: str, parquet_path: str | None = None) -> str:
    """Convert a CSV file to Parquet and return the output path."""
    import pandas as pd

    csv_file = Path(csv_path)
    if not csv_file.is_file():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    if parquet_path is None:
        parquet_path = str(csv_file.with_suffix(".parquet"))

    df = pd.read_csv(csv_path)
    df.to_parquet(parquet_path, engine="pyarrow", index=False)
    return parquet_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert a CSV file to Parquet.")
    parser.add_argument("csv_path", help="Path to the source CSV file")
    parser.add_argument(
        "parquet_path",
        nargs="?",
        default=None,
        help="Destination Parquet path (default: same name with .parquet)",
    )
    args = parser.parse_args()

    try:
        out = csv_to_parquet(args.csv_path, args.parquet_path)
    except Exception as exc:  # noqa: BLE001 - surface the error to the user
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
