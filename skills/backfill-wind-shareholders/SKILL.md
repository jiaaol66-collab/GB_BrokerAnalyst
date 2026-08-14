---
name: backfill-wind-shareholders
description: Read Wind Terminal Global Shareholder Holdings Summary through ToDesk and backfill missing institutional ownership percentages and shareholder totals into an existing Excel workbook, with explicit visual confirmation of the institutional-investor filter before every extraction. Use for GB Analyst or GB BrokerAnalyst workbooks whose rows contain stock code, stock name, previous/later report dates, and blank target columns that must be filled without overwriting existing values.
---

# Backfill Wind Shareholders

Use the Computer Use skill for ToDesk/Wind interaction and the Spreadsheets skill for the workbook. Read [references/wind-workflow.md](references/wind-workflow.md) before operating Wind.

## Workflow

1. Run `scripts/reuse_existing.py` first. For the same stock code, report date, and metric, copy a unique existing Excel value into duplicate blank cells; do not query Wind for those cells.
2. Run `scripts/list_missing.py` to produce the remaining HK-only work queue. Treat existing F/G/I/J values as immutable.
3. Group consecutive rows by stock code and all required report dates. Search each stock once, then switch dates.
4. For every required date:
   - Select the exact report date from Excel column C or D. Before reading anything, visually confirm that the date displayed in Wind exactly matches the target `yyyy-mm-dd`.
   - If the exact report date is not available in Wind, leave all still-blank target cells for that stock and report date blank and add them to the manual-processing list. Never substitute `最新`, a nearby date, or the currently displayed date.
   - Before reading an institutional percentage, visually confirm that `只查看机构投资者` is **checked**. If it is not checked, click the checkbox, wait for the table to finish updating, reacquire the screen, and confirm the checked state. Only then read the `合计` row under `占总股本比例(%)`.
   - Before reading a shareholder count, visually confirm that `只查看机构投资者` is **unchecked**. If it is checked, click the checkbox, wait for the table to finish updating, reacquire the screen, and confirm the unchecked state. Only then read the `合计` row under `股东数量`.
   - On the maximized 1277x752 ToDesk capture, `(534,126)` is `合并机构股东` and must not be used for this workflow. The `只查看机构投资者` checkbox center is near `(634,126)`. Always confirm the label before clicking.
   - On the same maximized capture, use the lower-right stock search launcher near `(1077,713)`, its opened-panel input near `(1035,455)`, and the upper-left report-date arrow near `(318,104)`. Do not use the upper-right `F5` box or the left navigation search. Select stock candidates and report dates by exact visible text, not by coordinates alone.
   - Never infer the date or filter state from a previous selection, click, metric, or expected workflow. Confirm both visible states again before every extraction.
5. Cache results as a JSON object keyed by cell reference. Store evidence for each newly observed value, including stock code, target report date, visibly confirmed Wind report date, metric, and the visibly confirmed filter state. Do not keep Excel open during collection.
6. Run `scripts/apply_backfill.py` after Excel is closed. It writes numeric values only into blank target cells and verifies worksheet invariants.
7. Re-run `list_missing.py` and independently inspect the completed rows. Report changed cells and any skipped existing cells.

## Mandatory extraction gate

Treat report-date and filter-state confirmation as hard preconditions, not best-effort checks.

- A value is valid only when the report date visibly displayed in Wind exactly equals the Excel report date for its destination cell.
- `institutional_ratio` is valid only when the checkbox itself is visibly checked immediately before reading.
- `shareholder_count` is valid only when the checkbox itself is visibly unchecked immediately before reading.
- After changing the date or clicking the checkbox, wait for the table to finish refreshing and use a new screen observation to confirm both states before extraction.
- If the exact date is absent from Wind, preserve existing cells and leave the remaining target cells blank for manual processing.
- If the date or checkbox is obscured, either state is visually ambiguous, the screenshot may be stale, or the table is still loading, do not read, cache, copy, or write the value. Move the pointer away, wait, reacquire the screen, and verify again. If either state remains uncertain, skip that metric and report it as unverified.
- Any value collected without a recorded confirmed state is unverified. Re-read it from Wind with the correct state before writeback; do not accept it merely because the number looks plausible.

## Commands

```powershell
python scripts/reuse_existing.py workbook.xlsx --sheet 导出计数_标识 --output reusable.json
python scripts/apply_backfill.py workbook.xlsx --sheet 导出计数_标识 --values reusable.json --audit reuse-audit.json
python scripts/list_missing.py workbook.xlsx --sheet 导出计数_标识 --output missing.json
python scripts/apply_backfill.py workbook.xlsx --sheet 导出计数_标识 --values collected.json --audit audit.json
```

Use the bundled workspace Python when available. Keep support JSON/audit files outside the source workbook directory unless the user requests otherwise.

## Stop rules

- Skip non-HK codes unless the user explicitly expands scope.
- Never overwrite populated target cells.
- Treat the workbook's existing stock code as authoritative. Select an exact code match even when the Wind candidate name differs by naming convention; do not rewrite the workbook's existing code or name.
- Do not extract a value unless the visible Wind report date exactly matches Excel and the visible `只查看机构投资者` state matches the metric's required state.
- When Wind has no exact matching report date, leave the corresponding missing cells blank and hand them off for manual processing; do not approximate or substitute another date.
- Do not modify unrelated sheets, formulas, styles, filters, freeze panes, conditional formatting, or workbook structure.
- If the workbook is locked, continue collecting to JSON and defer writeback until it is closed.
