#!/usr/bin/env python3
"""List HK rows with blank Wind backfill targets in an XLSX worksheet."""

from __future__ import annotations

import argparse
import json
import re
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

from apply_backfill import MAIN, sheet_member


def cell_value(cell: ET.Element, shared: list[str]):
    value = cell.find(f"{{{MAIN}}}v")
    inline = cell.find(f"{{{MAIN}}}is")
    if inline is not None:
        return "".join(inline.itertext())
    if value is None or value.text in (None, ""):
        return None
    if cell.get("t") == "s":
        return shared[int(value.text)]
    return value.text


def excel_date(value):
    if value in (None, ""):
        return None
    try:
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date().isoformat()
    except (TypeError, ValueError):
        return str(value)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--market-suffix", default=".HK")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    with zipfile.ZipFile(args.workbook, "r") as archive:
        shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{{{MAIN}}}si"):
                shared.append("".join(item.itertext()))

        root = ET.fromstring(archive.read(sheet_member(archive, args.sheet)))
        missing_rows = []
        for row in root.findall(f".//{{{MAIN}}}sheetData/{{{MAIN}}}row"):
            row_num = int(row.get("r"))
            values = {cell.get("r"): cell_value(cell, shared) for cell in row.findall(f"{{{MAIN}}}c")}
            code = values.get(f"A{row_num}")
            if not isinstance(code, str) or not code.upper().endswith(args.market_suffix.upper()):
                continue
            missing = [column for column in ("F", "G", "I", "J") if values.get(f"{column}{row_num}") in (None, "")]
            if not missing:
                continue
            missing_rows.append({
                "row": row_num,
                "code": code,
                "name": values.get(f"B{row_num}"),
                "previous_period": excel_date(values.get(f"C{row_num}")),
                "later_period": excel_date(values.get(f"D{row_num}")),
                "missing": missing,
                "existing": {column: values.get(f"{column}{row_num}") for column in ("F", "G", "I", "J")},
            })

    groups = []
    for item in missing_rows:
        if groups and groups[-1]["code"] == item["code"]:
            groups[-1]["rows"].append(item)
        else:
            groups.append({"code": item["code"], "name": item["name"], "rows": [item]})
    payload = {
        "workbook": str(args.workbook.resolve()),
        "sheet": args.sheet,
        "market_suffix": args.market_suffix,
        "missing_row_count": len(missing_rows),
        "groups": groups,
    }
    output = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output + "\n", encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
