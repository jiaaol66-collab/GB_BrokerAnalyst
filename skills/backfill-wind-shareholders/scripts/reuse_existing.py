#!/usr/bin/env python3
"""Find blank targets that can reuse an existing same-code/date/metric value."""

from __future__ import annotations

import argparse
import json
import zipfile
from collections import defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

from apply_backfill import MAIN, sheet_member
from list_missing import cell_value, excel_date


FIELD_MAP = {
    "F": ("C", "institutional_ratio"),
    "G": ("D", "institutional_ratio"),
    "I": ("C", "shareholder_count"),
    "J": ("D", "shareholder_count"),
}


def numeric(value):
    if value in (None, ""):
        return None
    number = float(value)
    return int(number) if number.is_integer() else number


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

    rows = []
    for row in root.findall(f".//{{{MAIN}}}sheetData/{{{MAIN}}}row"):
        row_num = int(row.get("r"))
        values = {cell.get("r"): cell_value(cell, shared) for cell in row.findall(f"{{{MAIN}}}c")}
        code = values.get(f"A{row_num}")
        if isinstance(code, str) and code.upper().endswith(args.market_suffix.upper()):
            rows.append((row_num, code, values))

    observed = defaultdict(set)
    sources = defaultdict(list)
    for row_num, code, values in rows:
        for target_col, (date_col, metric) in FIELD_MAP.items():
            date = excel_date(values.get(f"{date_col}{row_num}"))
            value = numeric(values.get(f"{target_col}{row_num}"))
            if date and value is not None:
                key = (code, date, metric)
                observed[key].add(value)
                sources[key].append(f"{target_col}{row_num}")

    reusable = {}
    reused_from = {}
    conflicts = []
    for row_num, code, values in rows:
        for target_col, (date_col, metric) in FIELD_MAP.items():
            target = f"{target_col}{row_num}"
            if values.get(target) not in (None, ""):
                continue
            date = excel_date(values.get(f"{date_col}{row_num}"))
            key = (code, date, metric)
            candidates = observed.get(key, set())
            if len(candidates) == 1:
                reusable[target] = next(iter(candidates))
                reused_from[target] = sources[key]
            elif len(candidates) > 1:
                conflicts.append({
                    "target": target,
                    "code": code,
                    "date": date,
                    "metric": metric,
                    "values": sorted(candidates),
                    "sources": sources[key],
                })

    payload = {
        "workbook": str(args.workbook.resolve()),
        "sheet": args.sheet,
        "values": reusable,
        "reused_from": reused_from,
        "conflicts": conflicts,
    }
    output = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(output + "\n", encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
