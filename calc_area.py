"""Simple script to sum the values in column F of 0615.xlsx."""
import sys

try:
    import pandas as pd
except Exception as e:
    sys.stderr.write("pandas is required to run this script\n")
    raise


def main(filename="0615.xlsx", column_index=5):
    try:
        df = pd.read_excel(filename)
    except FileNotFoundError:
        sys.stderr.write(f"File '{filename}' not found\n")
        return 1
    except Exception as e:
        sys.stderr.write(f"Failed to read '{filename}': {e}\n")
        return 1

    if column_index >= len(df.columns):
        sys.stderr.write(f"File only has {len(df.columns)} columns; cannot read column {column_index+1}\n")
        return 1

    numeric = pd.to_numeric(df.iloc[:, column_index], errors="coerce")
    total = numeric.sum(skipna=True)
    print(total)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

