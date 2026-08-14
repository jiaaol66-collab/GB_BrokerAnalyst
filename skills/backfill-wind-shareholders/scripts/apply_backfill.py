#!/usr/bin/env python3
"""Fill only blank numeric cells in one XLSX worksheet and verify invariants."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
DOCREL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKGREL = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("", MAIN)
ET.register_namespace("r", DOCREL)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sheet_member(archive: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rel_id = None
    for sheet in workbook.findall(f"{{{MAIN}}}sheets/{{{MAIN}}}sheet"):
        if sheet.get("name") == sheet_name:
            rel_id = sheet.get(f"{{{DOCREL}}}id")
            break
    if not rel_id:
        raise ValueError(f"Worksheet not found: {sheet_name}")

    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = None
    for rel in rels.findall(f"{{{PKGREL}}}Relationship"):
        if rel.get("Id") == rel_id:
            target = rel.get("Target")
            break
    if not target:
        raise ValueError(f"Worksheet relationship missing: {sheet_name}")
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def cell_col_number(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref.upper())
    if not letters:
        raise ValueError(f"Invalid cell reference: {ref}")
    value = 0
    for char in letters.group(0):
        value = value * 26 + ord(char) - 64
    return value


def cell_row_number(ref: str) -> int:
    match = re.search(r"\d+$", ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    return int(match.group(0))


def has_content(cell: ET.Element) -> bool:
    if cell.find(f"{{{MAIN}}}f") is not None:
        return True
    value = cell.find(f"{{{MAIN}}}v")
    if value is not None and value.text not in (None, ""):
        return True
    inline = cell.find(f"{{{MAIN}}}is")
    if inline is not None and "".join(inline.itertext()).strip():
        return True
    return False


def numeric_text(value: object) -> str:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise TypeError(f"Backfill value must be numeric, got {value!r}")
    if isinstance(value, float):
        return format(value, ".15g")
    return str(value)


def snapshot(root: ET.Element) -> dict[str, object]:
    cells = {}
    for cell in root.findall(f".//{{{MAIN}}}c"):
        ref = cell.get("r")
        if ref:
            cells[ref] = ET.tostring(cell, encoding="unicode")
    pane = root.find(f".//{{{MAIN}}}pane")
    auto_filter = root.find(f"{{{MAIN}}}autoFilter")
    return {
        "cells": cells,
        "formula_count": len(root.findall(f".//{{{MAIN}}}f")),
        "conditional_format_count": len(root.findall(f"{{{MAIN}}}conditionalFormatting")),
        "merge_count": len(root.findall(f"{{{MAIN}}}mergeCells/{{{MAIN}}}mergeCell")),
        "auto_filter": auto_filter.get("ref") if auto_filter is not None else None,
        "pane": dict(pane.attrib) if pane is not None else None,
    }


def patch_sheet(
    xml_bytes: bytes,
    values: dict[str, object],
    overwrite_existing: bool = False,
) -> tuple[bytes, dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    before = snapshot(root)
    sheet_data = root.find(f"{{{MAIN}}}sheetData")
    if sheet_data is None:
        raise ValueError("Worksheet has no sheetData")

    rows = {int(row.get("r")): row for row in sheet_data.findall(f"{{{MAIN}}}row")}
    changed = []
    skipped_existing = []

    for ref, value in sorted(values.items(), key=lambda item: (cell_row_number(item[0]), cell_col_number(item[0]))):
        ref = ref.upper()
        row_num = cell_row_number(ref)
        row = rows.get(row_num)
        if row is None:
            row = ET.Element(f"{{{MAIN}}}row", {"r": str(row_num)})
            inserted = False
            for index, existing_row in enumerate(sheet_data.findall(f"{{{MAIN}}}row")):
                if int(existing_row.get("r")) > row_num:
                    sheet_data.insert(index, row)
                    inserted = True
                    break
            if not inserted:
                sheet_data.append(row)
            rows[row_num] = row

        cell = next((item for item in row.findall(f"{{{MAIN}}}c") if item.get("r") == ref), None)
        if cell is not None and has_content(cell) and not overwrite_existing:
            skipped_existing.append(ref)
            continue
        if cell is None:
            cell = ET.Element(f"{{{MAIN}}}c", {"r": ref})
            inserted = False
            for index, existing_cell in enumerate(row.findall(f"{{{MAIN}}}c")):
                if cell_col_number(existing_cell.get("r")) > cell_col_number(ref):
                    row.insert(index, cell)
                    inserted = True
                    break
            if not inserted:
                row.append(cell)

        for child in list(cell):
            if child.tag in {f"{{{MAIN}}}v", f"{{{MAIN}}}is", f"{{{MAIN}}}f"}:
                cell.remove(child)
        cell.attrib.pop("t", None)
        value_node = ET.SubElement(cell, f"{{{MAIN}}}v")
        value_node.text = numeric_text(value)
        changed.append(ref)

    after = snapshot(root)
    targets = set(values)
    non_target_changes = sorted(
        ref for ref in set(before["cells"]) | set(after["cells"])
        if ref not in targets and before["cells"].get(ref) != after["cells"].get(ref)
    )
    invariant_keys = ["formula_count", "conditional_format_count", "merge_count", "auto_filter", "pane"]
    invariant_changes = {
        key: {"before": before[key], "after": after[key]}
        for key in invariant_keys if before[key] != after[key]
    }
    audit = {
        "changed": changed,
        "skipped_existing": skipped_existing,
        "non_target_cell_changes": non_target_changes,
        "invariant_changes": invariant_changes,
        "before": {key: before[key] for key in invariant_keys},
        "after": {key: after[key] for key in invariant_keys},
    }
    if non_target_changes or invariant_changes:
        raise RuntimeError(f"Workbook invariant verification failed: {json.dumps(audit, ensure_ascii=False)}")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), audit


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--sheet", required=True)
    parser.add_argument("--values", type=Path, required=True, help="JSON file with either a values object or cell-value mapping")
    parser.add_argument("--audit", type=Path)
    parser.add_argument(
        "--overwrite-existing",
        action="store_true",
        help="Overwrite only the explicitly listed cells; intended for correcting a prior backfill.",
    )
    args = parser.parse_args()

    workbook = args.workbook.resolve()
    payload = json.loads(args.values.read_text(encoding="utf-8-sig"))
    values = payload.get("values", payload)
    if not isinstance(values, dict):
        raise TypeError("Values JSON must contain an object mapping cell references to numbers")

    original_hash = sha256(workbook)
    with zipfile.ZipFile(workbook, "r") as source:
        member = sheet_member(source, args.sheet)
        original_members = source.namelist()
        patched_xml, audit = patch_sheet(
            source.read(member),
            values,
            overwrite_existing=args.overwrite_existing,
        )

        fd, temp_name = tempfile.mkstemp(prefix=f".{workbook.stem}-", suffix=".xlsx", dir=workbook.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            with zipfile.ZipFile(temp_path, "w") as target:
                for info in source.infolist():
                    data = patched_xml if info.filename == member else source.read(info.filename)
                    target.writestr(info, data)
            with zipfile.ZipFile(temp_path, "r") as check:
                if check.namelist() != original_members:
                    raise RuntimeError("ZIP member list changed")
                check.testzip()
                verified_member = sheet_member(check, args.sheet)
                verified = snapshot(ET.fromstring(check.read(verified_member)))
                for ref in audit["changed"]:
                    if ref not in verified["cells"]:
                        raise RuntimeError(f"Patched cell missing after write: {ref}")
            shutil.move(str(temp_path), str(workbook))
        finally:
            if temp_path.exists():
                temp_path.unlink()

    result = {
        "workbook": str(workbook),
        "sheet": args.sheet,
        "sheet_member": member,
        "original_sha256": original_hash,
        "updated_sha256": sha256(workbook),
        **audit,
    }
    output = json.dumps(result, ensure_ascii=False, indent=2)
    if args.audit:
        args.audit.parent.mkdir(parents=True, exist_ok=True)
        args.audit.write_text(output + "\n", encoding="utf-8")
    print(output)


if __name__ == "__main__":
    main()
