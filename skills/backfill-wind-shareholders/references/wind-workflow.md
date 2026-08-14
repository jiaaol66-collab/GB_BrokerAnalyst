# Wind workflow

## Fast path

1. Click the lower-right `代码/名称/简拼/功能` launcher near `(1077,713)`. In the opened panel, right-click the input near `(1035,455)`, then choose `粘贴` near `(1074,514)` to enter the numeric code prefix. Never use the upper-right `F5` box or the left navigation search.
2. Verify the stock code. The workbook's existing code is authoritative. If the first candidate code is an exact match, press Enter; otherwise continue typing or click the exact code match. A different short name is acceptable and must not trigger a workbook name correction.
3. Work by stock, then by unique report date. Do not re-enter the code for consecutive rows with the same stock.
4. Select the exact report date shown in Excel column C or D. Confirm that the Wind date field visibly displays the same `yyyy-mm-dd` before continuing. If that exact date is absent, leave the corresponding missing cells blank and add them to the manual-processing list; never use another date.
5. Set and verify `只查看机构投资者` for every metric:
   - checked: after visibly confirming the checked state, read `合计 → 占总股本比例(%)` for F (column C date) or G (column D date);
   - unchecked: after visibly confirming the unchecked state, read `合计 → 股东数量` for I (column C date) or J (column D date).
6. Record all required values from the current date before switching dates.

## Mandatory date and filter-state verification

The report date and filter state must both be visibly confirmed immediately before every extraction. Never infer either state from a previous selection, click, or expected workflow.

1. Compare the target Excel report date with the date visibly displayed in Wind. They must match exactly in `yyyy-mm-dd` form.
2. If the target date does not appear in the Wind date list, stop processing that report date. Preserve populated cells, leave its remaining target cells blank, and record the stock code, report date, and destination cells for manual processing. Do not select `最新`, the nearest available date, or any substitute date.
3. Locate the checkbox square immediately to the left of `只查看机构投资者`.
4. Compare its visible state with the metric requirement:
   - institutional percentage: blue filled box with a white mark;
   - shareholder count: empty light-gray outline.
5. If the state is wrong, click the checkbox square once, wait for the table to finish updating, move the pointer away, and reacquire the screen.
6. In the new observation, reconfirm both the exact report date and the checkbox state.
7. Only after both confirmations read the required `合计` value.

Do not read, cache, copy, or write a value when the date or checkbox is obscured, either state is ambiguous, the observation may be stale, or the table is still loading. Reacquire and verify; if certainty cannot be restored, skip the metric and mark it unverified.

Record evidence for each newly observed value: stock code, target Excel report date, confirmed visible Wind report date, metric, confirmed filter state (`checked` or `unchecked`), destination cells, and observation time or screenshot reference. A value without both confirmation records is unverified and must be re-read before writeback.

## Strict control locations

Always identify controls by label and relative position first. Use coordinates only as a fallback when the ToDesk window is maximized at 1280×752.

| Control | Exact visual location | Fallback coordinate | State check |
| --- | --- | --- | --- |
| Stock-code search launcher | Lower-right status area labeled `代码/名称/简拼/功能` | Near `(1077, 713)` | The lower-right search panel must open |
| Stock-code input | Top input line inside the opened lower-right search panel | Right-click near `(1035, 455)`; choose paste near `(1074, 514)` | The pasted numeric prefix must be visible; click only the exact `.HK` candidate |
| Report date | First filter row at the upper-left of the content area: label `日期`, followed immediately by the current date value and its drop-down arrow | Click the arrow near `(318, 104)` | The selected `yyyy-mm-dd` must appear inside the date field |
| Merge institutional holders (do not toggle) | Second filter row: checkbox immediately to the left of `合并机构股东` | Near `(534, 126)` | This is not the institutional-only filter; leave it untouched |
| Institutional filter | Second filter row: the checkbox immediately to the left of the text `只查看机构投资者`; it is to the right of `合并机构股东` | Click near `(634, 126)` | Checked is a blue filled box with a white mark; unchecked is an empty light-gray outline |

Do not confuse the institutional filter with:

- the `股东类型` drop-down on the first filter row;
- the `合并机构股东` checkbox immediately to its left;
- the `股东类型统计` tab above the chart.

These fallback coordinates were calibrated with both Wind and ToDesk maximized in a 1277x752 ToDesk capture. For stock changes, use only the lower-right search launcher and its panel input. When switching dates, click only the upper-left date field arrow near `(318,104)`, then click the exact visible date text; never choose a date solely from its vertical position. When toggling the institutional filter, click the checkbox square rather than the text or adjacent checkbox. Never use `(534,126)` for `只查看机构投资者`. If the window is resized, ignore the fallback coordinates and relocate by labels.

The date list is descending and may begin with `最新`. Select by the visible date text from Excel, never by list position alone. Opening another stock normally returns the date to `最新`; therefore reselect the required report date after every stock change.

An unavailable report date is a valid manual-processing outcome. Leave the related missing cells blank. Do not treat another displayed date as equivalent, and do not copy a value from another report date.

Do not add user-facing confirmation between clicks. Continue directly while the page state is clear. Refresh/reacquire the window only after misclicks, focus loss, user interruption, stale screenshot errors, or uncertain state.

## Recovery after leaving the page

Navigate in Wind through:

`股票 → 深度资料 → 全球股东持仓汇总`

If the current stock is absent after opening the page, input the same stock code again and select the verified matching candidate.

## Mapping

| Excel field | Wind date | Wind state | Wind value |
| --- | --- | --- | --- |
| F | C 前一期 | 只查看机构投资者 checked | 合计/占总股本比例(%) |
| G | D 后一期 | 只查看机构投资者 checked | 合计/占总股本比例(%) |
| I | C 前一期 | 只查看机构投资者 unchecked | 合计/股东数量 |
| J | D 后一期 | 只查看机构投资者 unchecked | 合计/股东数量 |

## Performance rules

- Build the full missing-cell queue before opening ToDesk.
- Before Wind lookup, reuse a unique existing Excel value for every duplicate `(stock code, report date, metric)` combination. Existing workbook values take precedence; do not replace them with a newly observed Wind value.
- For one stock, order dates to minimize dropdown travel and checkbox toggles.
- Skip a metric when its destination cell already has data.
- Cache each unique `(stock code, report date, filter state)` result and reuse it across duplicate rows.
- Reuse only cache entries whose confirmed Wind report date exactly matches the Excel report date and whose recorded filter state matches the metric requirement. Never reuse an unverified observation.
- Save the cache after every completed stock so interruptions do not lose work.
