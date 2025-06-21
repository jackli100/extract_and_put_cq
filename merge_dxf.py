import argparse
import ezdxf
from ezdxf.addons import Importer
from ezdxf.document import Drawing
from pathlib import Path


def _reset_insbase(doc: Drawing) -> None:
    """Set INSBASE of *doc* to (0, 0, 0) to avoid automatic offsets."""
    try:
        doc.header["$INSBASE"] = (0, 0, 0)
    except Exception:
        pass


def merge_from_folder(folder_path: str, output: str) -> None:
    """Merge all DXF files from a folder into a single output DXF."""
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"Folder {folder_path} does not exist")
        return
    
    if not folder.is_dir():
        print(f"{folder_path} is not a directory")
        return
    
    # Find all DXF files in the folder
    dxf_files = list(folder.glob("*.dxf"))
    
    if not dxf_files:
        print(f"No DXF files found in {folder_path}")
        return
    
    print(f"Found {len(dxf_files)} DXF files to merge")
    
    merged = ezdxf.new()
    msp = merged.modelspace()
    
    merged_count = 0
    for dxf_file in dxf_files:
        try:
            doc = ezdxf.readfile(dxf_file)
        except Exception as e:
            print(f"Failed to read {dxf_file}: {e}")
            continue

        _reset_insbase(doc)
        importer = Importer(doc, merged)
        importer.import_modelspace(msp)
        importer.finalize()
        print(f"Merged {dxf_file.name}")
        merged_count += 1
    
    if merged_count > 0:
        merged.saveas(output)
        print(f"Successfully merged {merged_count} files into {output}")
    else:
        print("No files were successfully merged")


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
        _reset_insbase(doc)
        importer = Importer(doc, merged)
        importer.import_modelspace(msp)
        importer.finalize()
        print(f"Merged {f}")
    merged.saveas(output)
    print(f"Written merged file to {output}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Merge multiple DXF files")
    parser.add_argument("output", help="Path of merged DXF file")
    
    # 添加互斥参数组
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument("-f", "--folder", help="Folder containing DXF files to merge")
    input_group.add_argument("files", nargs="*", help="Individual DXF files to merge")
    
    args = parser.parse_args()
    
    if args.folder:
        merge_from_folder(args.folder, args.output)
    else:
        merge(args.files, args.output)
