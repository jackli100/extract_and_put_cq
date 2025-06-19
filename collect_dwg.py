import argparse
import os
import shutil


def collect_dwg(src_dir: str, dest_dir: str) -> None:
    """Copy all DWG files from ``src_dir`` to ``dest_dir`` without duplicates."""
    os.makedirs(dest_dir, exist_ok=True)
    seen = set()
    for root, _, files in os.walk(src_dir):
        for name in files:
            if not name.lower().endswith('.dwg'):
                continue
            if name in seen:
                continue
            seen.add(name)
            src_path = os.path.join(root, name)
            dest_path = os.path.join(dest_dir, name)
            try:
                shutil.copy2(src_path, dest_path)
                print(f"Copied {src_path} -> {dest_path}")
            except Exception as e:
                print(f"Failed to copy {src_path}: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Collect unique DWG files")
    parser.add_argument("src", help="Source directory to scan")
    parser.add_argument("dest", help="Directory to copy unique DWG files to")
    args = parser.parse_args()
    collect_dwg(args.src, args.dest)
