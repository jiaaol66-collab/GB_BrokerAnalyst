from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from apply_backfill import MAIN, sheet_member, snapshot


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--cells", required=True, help="Comma-separated cell references")
    args = parser.parse_args()

    targets = {item.strip().upper() for item in args.cells.split(",") if item.strip()}
    with ZipFile(args.workbook, "r") as source:
        member = sheet_member(source, args.sheet)
        original = source.read(member)
        root = ET.fromstring(original)
        before = snapshot(root)
        cleared: list[str] = []
        for cell in root.findall(f".//{{{MAIN}}}c"):
            ref = cell.get("r", "").upper()
            if ref not in targets:
                continue
            for tag in ("v", "f", "is"):
                child = cell.find(f"{{{MAIN}}}{tag}")
                if child is not None:
                    cell.remove(child)
            cell.attrib.pop("t", None)
            cleared.append(ref)
        patched = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        after = snapshot(ET.fromstring(patched))
        invariant_keys = (
            "formula_count",
            "conditional_format_count",
            "merge_count",
            "auto_filter",
            "pane",
        )
        before_invariants = {key: before[key] for key in invariant_keys}
        after_invariants = {key: after[key] for key in invariant_keys}
        if before_invariants != after_invariants:
            raise RuntimeError(
                f"Workbook invariant changed: before={before_invariants!r}, after={after_invariants!r}"
            )
        with tempfile.NamedTemporaryFile(delete=False, suffix=args.workbook.suffix) as temp:
            temp_path = Path(temp.name)
        try:
            with ZipFile(temp_path, "w", ZIP_DEFLATED) as output:
                for info in source.infolist():
                    output.writestr(info, patched if info.filename == member else source.read(info.filename))
            shutil.copyfile(temp_path, args.workbook)
        finally:
            temp_path.unlink(missing_ok=True)

    missing = sorted(targets - set(cleared))
    print({"cleared": sorted(cleared), "missing": missing, "invariants": before_invariants})


if __name__ == "__main__":
    main()
