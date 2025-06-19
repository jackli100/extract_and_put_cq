import argparse
import ezdxf
from ezdxf.addons import Importer
from pathlib import Path


def merge(files, output: str) -> None:
    """Merge all entities from ``files`` into ``output`` DXF."""
    if not files:
        print("No DXF files supplied")
        return
    merged = ezdxf.new()
    msp = merged.modelspace()
    for f in files:
        try:
            doc = ezdxf.readfile(f)
        except Exception as e:
            print(f"Failed to read {f}: {e}")
            continue
        importer = Importer(doc, merged)
        importer.import_modelspace(msp)
        importer.finalize()
        print(f"Merged {f}")
    merged.saveas(output)
    print(f"Written merged file to {output}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Merge multiple DXF files")
    parser.add_argument("output", help="Path of merged DXF file")
    parser.add_argument("files", nargs="+", help="DXF files to merge")
    args = parser.parse_args()
    merge(args.files, args.output)
