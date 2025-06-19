import argparse
import os
from pathlib import Path
from ezdxf.addons import odafc


def convert_directory(src: str, dest: str, version: str = "R2013") -> None:
    """Convert all DWG files in ``src`` to DXF in ``dest``.

    Requires the ODA File Converter to be installed and accessible.
    """
    os.makedirs(dest, exist_ok=True)
    for name in os.listdir(src):
        if not name.lower().endswith('.dwg'):
            continue
        src_path = Path(src) / name
        dest_path = Path(dest) / (src_path.stem + '.dxf')
        try:
            odafc.convert(str(src_path), str(dest_path), version=version, replace=True)
            print(f"Converted {src_path} -> {dest_path}")
        except Exception as e:
            print(f"Failed to convert {src_path}: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert DWG files to DXF")
    parser.add_argument("src", help="Directory containing DWG files")
    parser.add_argument("dest", help="Destination directory for DXF files")
    parser.add_argument(
        "--version", default="R2013", help="DXF version for output (default: R2013)"
    )
    args = parser.parse_args()
    convert_directory(args.src, args.dest, args.version)
