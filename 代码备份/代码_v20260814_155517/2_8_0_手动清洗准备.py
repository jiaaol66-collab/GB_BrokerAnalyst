import sys, os, time
from copy import copy, deepcopy
from datetime import timedelta
from pathlib import Path
os.chdir(os.path.dirname(__file__))

def find_and_add_package_path(
        package_name: str,
        search_depth: int = 5,
        root_dir: Path = None
) -> bool:
    """
    在文件系统中搜索指定包并添加到 Python 路径

    :param package_name: 要搜索的包名称 (如 'lib')
    :param search_depth: 最大搜索深度 (默认5级)
    :param root_dir: 搜索起始目录 (默认当前目录)
    :return: 是否成功找到并添加包路径
    """
    # 设置起始搜索目录
    if root_dir is None:
        root_dir = Path(os.getcwd())

    # 检查是否已经存在该包
    try:
        __import__(package_name)
        print(f"{package_name} 已在 Python 路径中")
        return True
    except ImportError:
        pass

    # 准备包可能的标识文件名
    package_markers = [
        '__init__.py',
        '__init__.pyi',
        f'{package_name}.py'
    ]

    # 开始搜索
    found_paths = []
    print(f"正在搜索包: {package_name} ...")

    for depth in range(search_depth + 1):
        if depth == 0:
            search_dirs = [root_dir]
        else:
            # 生成搜索模式: 当前目录上溯几层目录
            up_parents = Path(os.path.join(*(['..'] * depth)))
            search_dirs = [
                (root_dir / up_parents).resolve(),
                (root_dir / up_parents / 'src').resolve(),
                (root_dir / up_parents / 'lib').resolve()
            ]

        for search_dir in search_dirs:
            # 检查目录存在性
            if not search_dir.exists() or not search_dir.is_dir():
                continue

            # 检查包目录是否存在
            package_dir = search_dir / package_name
            if package_dir.exists() and package_dir.is_dir():
                # 验证是有效Python包
                is_valid = any(
                    (package_dir / marker).exists()
                    for marker in package_markers
                )
                if is_valid:
                    found_paths.append(str(search_dir))
                    print(f"在 {search_dir} 发现包 {package_name}")

    # 添加到系统路径
    added = False
    for path in found_paths:
        if path not in sys.path:
            sys.path.insert(0, path)
            print(f"添加路径到 sys.path: {path}")
            added = True

    if added:
        # 重新尝试导入
        try:
            __import__(package_name)
            print(f"成功导入 {package_name}")
            return True
        except ImportError as e:
            print(f"导入失败: {str(e)}")

    if not found_paths:
        print(f"未能在搜索深度 {search_depth} 内找到 {package_name}")

    return added

if __name__ == '__main__':
    target_package = "userlib"

    try:
        from userlib.tools import *

        print(f"{target_package} 已存在")
    except ImportError:
        print(f"尝试搜索并添加 {target_package}...")
        # 执行搜索
        found = find_and_add_package_path(
            package_name=target_package,
            search_depth=5,
            root_dir=Path(__file__).resolve().parent
        )

        if found:
            # 尝试使用导入
            exec(f"from {target_package}.tools import *")

            print("成功调用工具包函数")
        else:
            print("无法找到包，请检查目录结构")
            time.sleep(2)
            sys.exit(0)

import warnings
warnings.filterwarnings('ignore')

from openpyxl import Workbook, load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries

rel = find_rel(__file__, '参评信息')
start = insert_bookmark(__file__,0, rel)
impt = find_rel(__file__, '重要文件')
mtr = find_rel(__file__, '资料')
processed = find_rel(__file__, '已处理')
org = find_rel(__file__, '原始数据集')
yr = get_year(rel)
make_dir('输出文件')

BASE_COLUMNS = [
    '分析师姓名', '唯一标识', '公司名称', '研究领域', '研究领域清洗版',
    '标的股票代码(原)', '标的股票代码(清洗)', '标的股票代码',
    '标的股票名称(原)', '标的股票名称', '标的股票名称(对比)', '同行业指数',
    '推荐日期', '推荐日期清洗版', '推荐强度', '推荐强度清洗版',
    '看涨/看跌', '权重', '机构编号', '研报发布时标的价格(原)',
    '研报推荐的目标价格(原)', '研报推荐的目标价格',
    '研报推荐的目标价格下限', '上限', '研报发布时标的价格'
]
# 正式模板以 H 列“标的股票代码”为最终代码，I 列仅用于 Wind 公式回填核对。
# 上游数据仍沿用“标的股票代码(清洗)”和“标的股票代码”两个旧字段名，
# 因此在写模板时显式映射，避免改动前面各清洗步骤的接口。
OUTPUT_COLUMNS = [
    '序号', '分析师姓名', '唯一标识', '公司名称', '研究领域',
    '研究领域清洗版', '标的股票代码(原)', '标的股票代码',
    '标的股票代码(回填)', '标的股票名称(原)', '标的股票名称',
    '标的股票名称(对比)', '同行业指数', '推荐日期', '推荐日期清洗版',
    '推荐强度', '推荐强度清洗版', '看涨/看跌', '权重', '机构编号',
    '研报发布时标的价格(原)', '研报推荐的目标价格(原)',
    '研报推荐的目标价格', '研报推荐的目标价格下限', '上限',
    '研报发布时标的价格'
]
OUTPUT_SOURCE_COLUMN_MAP = {column: column for column in OUTPUT_COLUMNS}
OUTPUT_SOURCE_COLUMN_MAP.update({
    '标的股票代码': '标的股票代码(清洗)',
    '标的股票代码(回填)': '标的股票代码',
})
OUTPUT_COLUMN_LETTERS = {
    column_name: get_column_letter(column_index)
    for column_index, column_name in enumerate(OUTPUT_COLUMNS, 1)
}
NON_STOCK_DOMAINS = ['金融工程', '固定收益', '宏观', '策略']
STOCK_CONTENT_MARKER = '_stock_row_has_any_value'
STOCK_DETAIL_COLUMNS = [
    '标的股票代码(原)',
    '标的股票名称(原)',
    '推荐日期',
    '推荐强度',
    '研报发布时标的价格(原)',
    '研报推荐的目标价格(原)',
    '研报标题',
]
STOCK_PLACEHOLDERS = {'', '-', 'nan', 'none', 'nat'}
EXCEL_ERROR_TEXTS = {
    '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'
}
SKIP_OPEN_FILES = os.environ.get('GB_SKIP_OPEN_FILES', '').strip() == '1'


def open_output(path):
    """测试时可用 GB_SKIP_OPEN_FILES=1 禁止拉起 Office，正常运行保持原行为。"""
    if not SKIP_OPEN_FILES:
        open_file(str(path))


def to_output_layout(dataframe, extra_columns=None):
    """把上游旧字段名映射成当前模板 A:Z 列名与顺序。"""
    extra_columns = extra_columns or []
    source_columns = [
        OUTPUT_SOURCE_COLUMN_MAP[column]
        for column in OUTPUT_COLUMNS
    ]
    output = dataframe[source_columns].copy()
    output.columns = OUTPUT_COLUMNS
    for column in extra_columns:
        output[column] = dataframe[column]
    return output


def resolve_template_path():
    """优先使用当前通用模板；年度已处理文件和资料副本仅作兼容回退。"""
    candidates = [
        Path(org) / '通用' / '模板复制.xlsx',
        Path(processed) / str(yr) / '模板复制.xlsx',
        Path(processed) / str(yr) / '模板' / '模板复制.xlsx',
        Path(mtr) / '代码' / '模板' / str(yr) / '模板复制.xlsx',
    ]

    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()
    candidate_text = '\n'.join(f'  - {path}' for path in candidates)
    raise FileNotFoundError(f'未找到{yr}年模板复制.xlsx，已检查：\n{candidate_text}')


def last_business_day_on_or_before(day):
    """若6月30日是周末，则按说明回退到当周最后一个工作日。"""
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day


def append_reason(reason_series, mask, reason):
    """同一条记录可累计多个建议剔除原因。"""
    mask = mask.fillna(False)
    empty = mask & reason_series.eq('')
    non_empty = mask & reason_series.ne('')
    reason_series.loc[empty] = reason
    reason_series.loc[non_empty] = reason_series.loc[non_empty] + '；' + reason


def stock_row_has_any_value(dataframe):
    """股票相关字段任意一个有值即视为有内容，后续不得按空槽位删除。"""
    if STOCK_CONTENT_MARKER in dataframe.columns:
        return dataframe[STOCK_CONTENT_MARKER].fillna(False).astype(bool)
    available_columns = [
        column for column in STOCK_DETAIL_COLUMNS
        if column in dataframe.columns
    ]
    if not available_columns:
        return pd.Series(False, index=dataframe.index)
    value_masks = [
        ~dataframe[column]
        .astype(str)
        .str.strip()
        .str.lower()
        .isin(STOCK_PLACEHOLDERS)
        for column in available_columns
    ]
    return pd.concat(value_masks, axis=1).any(axis=1)


def stock_identity_blank_mask(dataframe):
    """全部股票填报字段都为空，才视为一个空股票槽位。"""
    return ~stock_row_has_any_value(dataframe)


def remove_redundant_blank_stock_rows(stock_data):
    """
    空股票槽位一直保留到本步骤再处理：

    - 同一分析师同一领域已有任意填报内容：纯空槽位静默剔除，不进取消资格；
    - 同组全部为空：只保留第一条空占位，供取消资格清单追踪。
    """
    data = stock_data.copy()
    group_columns = ['唯一标识', '研究领域清洗版']
    blank_mask = stock_identity_blank_mask(data)
    group_keys = [data[column] for column in group_columns]
    group_has_stock = (~blank_mask).groupby(group_keys, dropna=False).transform('any')

    redundant_with_stock = blank_mask & group_has_stock
    data = data.loc[~redundant_with_stock].copy()

    remaining_blank = stock_identity_blank_mask(data)
    repeated_all_blank = remaining_blank & data.duplicated(
        subset=group_columns,
        keep='first'
    )
    data = data.loc[~repeated_all_blank].copy().reset_index(drop=True)
    return data, int(redundant_with_stock.sum()), int(repeated_all_blank.sum())


def split_suggested_exclusions(stock_data):
    """先保留固定序号，再按说明形成建议剔除清单；保留数据绝不重新编号。"""
    data = stock_data.copy()
    reasons = pd.Series('', index=data.index, dtype='object')

    # 只有全部股票字段均空的占位才按“未填报股票”处理；任一字段有值均先保留。
    append_reason(
        reasons,
        stock_identity_blank_mask(data),
        '未填报标的股票代码及名称'
    )

    cleaned_dates = pd.to_datetime(data['推荐日期清洗版'], errors='coerce')
    start_boundary = pd.Timestamp(int(yr) - 1, 7, 1)
    end_boundary = pd.Timestamp(
        last_business_day_on_or_before(datetime(int(yr), 6, 30))
    )
    append_reason(reasons, cleaned_dates.isna(), '推荐日期清洗版缺失或无法识别')
    append_reason(
        reasons,
        cleaned_dates.le(start_boundary),
        f'推荐日期不晚于{start_boundary:%Y-%m-%d}'
    )
    append_reason(
        reasons,
        cleaned_dates.gt(end_boundary),
        f'推荐日期晚于{end_boundary:%Y-%m-%d}'
    )

    suggested = data.loc[reasons.ne('')].copy()
    suggested['建议剔除原因'] = reasons.loc[reasons.ne('')]
    kept = data.loc[reasons.eq('')].copy()
    suggested.drop(columns=[STOCK_CONTENT_MARKER], errors='ignore', inplace=True)
    kept.drop(columns=[STOCK_CONTENT_MARKER], errors='ignore', inplace=True)
    return kept, suggested, start_boundary, end_boundary


def excel_value(value):
    """把 pandas/numpy 标量转换为 openpyxl 可安全写入的值，并清除错误字面量。"""
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, str) and value.strip().upper() in EXCEL_ERROR_TEXTS:
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if hasattr(value, 'item'):
        try:
            return value.item()
        except (TypeError, ValueError):
            pass
    return value


def translated_formula(formula, source_coordinate, target_coordinate):
    """按照模板第2行公式生成目标行公式。"""
    return Translator(formula, origin=source_coordinate).translate_formula(target_coordinate)


def style_list_sheet(worksheet, dataframe):
    """为建议剔除/取消资格清单设置筛选、冻结窗格和基础可读格式。"""
    yellow = PatternFill('solid', fgColor='FFF2CC')
    green = PatternFill('solid', fgColor='E2F0D9')
    for col_index, column_name in enumerate(dataframe.columns, 1):
        cell = worksheet.cell(1, col_index, column_name)
        cell.font = Font(bold=True)
        cell.fill = yellow
        cell.alignment = Alignment(horizontal='center', vertical='center')
        values = [str(column_name)] + [str(value) for value in dataframe[column_name].dropna().head(300)]
        width = min(max(max(map(len, values)) + 2, 10), 42)
        worksheet.column_dimensions[get_column_letter(col_index)].width = width
        if column_name == '建议剔除原因':
            for row_index in range(2, len(dataframe) + 2):
                worksheet.cell(row_index, col_index).fill = green
    last_column = get_column_letter(max(1, len(dataframe.columns)))
    last_row = max(1, len(dataframe) + 1)
    # 只冻结标题行，不锁住A列，方便横向查看右侧模板。
    worksheet.freeze_panes = 'A2'
    worksheet.auto_filter.ref = f'A1:{last_column}{last_row}'


def write_dataframe_sheet(worksheet, dataframe):
    for col_index, column_name in enumerate(dataframe.columns, 1):
        worksheet.cell(1, col_index, column_name)
    for row_index, row_values in enumerate(dataframe.itertuples(index=False, name=None), 2):
        for col_index, value in enumerate(row_values, 1):
            worksheet.cell(row_index, col_index, excel_value(value))
    style_list_sheet(worksheet, dataframe)


def write_cancel_workbook(dataframe, target_path):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = '建议剔除'
    write_dataframe_sheet(
        worksheet,
        to_output_layout(dataframe, ['建议剔除原因'])
    )
    workbook.save(target_path)


def add_manual_cleaning_conditional_formats(worksheet, last_row, last_column):
    """复刻说明文档中的人工核对配色，并把范围延伸到所有数据行。"""
    if last_row < 2:
        return
    stock_code = OUTPUT_COLUMN_LETTERS['标的股票代码']
    backfilled_code = OUTPUT_COLUMN_LETTERS['标的股票代码(回填)']
    original_name = OUTPUT_COLUMN_LETTERS['标的股票名称(原)']
    cleaned_name = OUTPUT_COLUMN_LETTERS['标的股票名称']
    compared_name = OUTPUT_COLUMN_LETTERS['标的股票名称(对比)']
    # 差异样式需把同一实色同时写入前景和背景；Excel 实际从背景色槽取色。
    # 8 位颜色格式为 ARGB，显式使用 FF，避免被写成透明色 00。
    periwinkle = PatternFill(
        'solid', fgColor='FFE4DFEC', bgColor='FFE4DFEC'
    )
    cornflower = PatternFill(
        'solid', fgColor='FFDAEEF3', bgColor='FFDAEEF3'
    )
    chocolate = PatternFill(
        'solid', fgColor='FFFDE9D9', bgColor='FFFDE9D9'
    )
    scarlet = PatternFill(
        'solid', fgColor='FFF9DBDF', bgColor='FFF9DBDF'
    )

    worksheet.conditional_formatting.add(
        f'{backfilled_code}2:{backfilled_code}{last_row}',
        FormulaRule(
            formula=[f'LEFT({backfilled_code}2,7)<>LEFT({stock_code}2,7)'],
            fill=periwinkle
        )
    )
    worksheet.conditional_formatting.add(
        f'{stock_code}2:{stock_code}{last_row}',
        FormulaRule(
            formula=[f'{stock_code}2<>{backfilled_code}2'],
            fill=cornflower
        )
    )
    worksheet.conditional_formatting.add(
        f'{original_name}2:{original_name}{last_row}',
        FormulaRule(
            formula=[
                f'NOT(ISNUMBER(SEARCH({cleaned_name}2,{compared_name}2)))'
            ],
            fill=chocolate
        )
    )
    worksheet.conditional_formatting.add(
        f'A2:{last_column}{last_row}',
        FormulaRule(formula=['CELL("row")=ROW()'], fill=scarlet)
    )


def capture_template_conditional_formats(worksheet):
    """
    在清空模板示例行前，完整保存模板自带的条件格式规则。

    正式模板的旧式条件格式把实色写在 bgColor，fgColor 则是透明黑。
    Excel 能直接显示原模板，但 openpyxl 另存后可能显示为无填充；复制时把
    同一个颜色同步到 solid 填充真正使用的 fgColor，颜色主题及 tint 均不变。
    """
    captured = []
    for conditional_range in worksheet.conditional_formatting:
        rules = []
        for source_rule in worksheet.conditional_formatting[conditional_range]:
            rule = deepcopy(source_rule)
            if rule.dxf is not None and rule.dxf.fill is not None:
                fill = rule.dxf.fill
                foreground = fill.fgColor
                background = fill.bgColor
                if (
                    fill.fill_type == 'solid'
                    and foreground is not None
                    and foreground.type == 'rgb'
                    and foreground.rgb == '00000000'
                    and background is not None
                    and background.type in {'rgb', 'theme', 'indexed'}
                ):
                    fill.fgColor = copy(background)
            rules.append(rule)
        captured.append((str(conditional_range.sqref), rules))
    return captured


def extend_row2_range_to_last_row(source_range, last_row):
    """把模板中仅作用于第2行的区域按原列范围延伸到主表末行。"""
    extended_parts = []
    for range_part in source_range.split():
        min_col, min_row, max_col, max_row = range_boundaries(range_part)
        if min_row == 2 and max_row == 2:
            start = f'{get_column_letter(min_col)}2'
            end = f'{get_column_letter(max_col)}{last_row}'
            extended_parts.append(f'{start}:{end}')
        else:
            extended_parts.append(range_part)
    return ' '.join(extended_parts)


def restore_and_extend_template_conditional_formats(
        worksheet,
        captured_rules,
        last_row
):
    """原样恢复模板条件格式，并将其范围扩展至全部主表数据行。"""
    worksheet.conditional_formatting = ConditionalFormattingList()
    for source_range, rules in captured_rules:
        target_range = extend_row2_range_to_last_row(source_range, last_row)
        for rule in rules:
            worksheet.conditional_formatting.add(target_range, rule)


def build_filled_template_workbook(dataframe, suggested, template_path, target_path):
    """以正式模板为底稿，主表保留全部 A:Z 记录并按模板填满公式。"""
    workbook = load_workbook(template_path, data_only=False)
    worksheet = workbook[workbook.sheetnames[0]]
    template_last_column = worksheet.max_column
    template_last_letter = get_column_letter(template_last_column)
    data_last_column = len(OUTPUT_COLUMNS)
    data_last_letter = get_column_letter(data_last_column)
    formula_first_column = data_last_column + 1
    template_headers = [worksheet.cell(1, column).value for column in range(1, template_last_column + 1)]

    if template_last_letter != 'BM':
        raise ValueError(f'模板末列应为BM，实际为{template_last_letter}')
    if template_headers[:len(OUTPUT_COLUMNS)] != OUTPUT_COLUMNS:
        raise ValueError(
            f'正式模板A:{data_last_letter}列名与脚本输出列名不一致，请先核对模板'
        )

    row2_values = {column: worksheet.cell(2, column).value for column in range(1, template_last_column + 1)}
    template_conditional_formats = capture_template_conditional_formats(worksheet)
    row2_styles = {column: copy(worksheet.cell(2, column)._style) for column in range(1, template_last_column + 1)}
    row2_number_formats = {column: worksheet.cell(2, column).number_format for column in range(1, template_last_column + 1)}
    row2_alignments = {column: copy(worksheet.cell(2, column).alignment) for column in range(1, template_last_column + 1)}
    row2_protections = {column: copy(worksheet.cell(2, column).protection) for column in range(1, template_last_column + 1)}

    # 保留原清洗版的主表名称；建议剔除仅作为第二张表，不从主表预先删除。
    worksheet.title = 'Sheet1'
    for extra_sheet in list(workbook.worksheets)[1:]:
        workbook.remove(extra_sheet)
    if worksheet.max_row > 1:
        worksheet.delete_rows(2, worksheet.max_row - 1)

    output_dataframe = to_output_layout(dataframe)
    for row_index, row_values in enumerate(
            output_dataframe.itertuples(index=False, name=None),
            2
    ):
        for column in range(1, template_last_column + 1):
            target_cell = worksheet.cell(row_index, column)
            target_cell._style = copy(row2_styles[column])
            target_cell.number_format = row2_number_formats[column]
            target_cell.alignment = copy(row2_alignments[column])
            target_cell.protection = copy(row2_protections[column])

        for column, value in enumerate(row_values, 1):
            worksheet.cell(row_index, column, excel_value(value))

        # A:Z 中的核对公式和 AA:BM 中的计算公式均以模板第2行为准；
        # 脚本不再重写具体公式，也不额外包裹 IFERROR。
        for column in range(1, template_last_column + 1):
            column_letter = get_column_letter(column)
            target_cell = worksheet.cell(row_index, column)
            source_value = row2_values[column]

            if isinstance(source_value, str) and source_value.startswith('='):
                formula = translated_formula(
                    source_value,
                    f'{column_letter}2',
                    f'{column_letter}{row_index}'
                )
                target_cell.value = formula
            elif column >= formula_first_column:
                # BJ、BK、BM为人工/后续步骤填入列，按正式模板保持空白。
                target_cell.value = None

    last_row = max(1, len(dataframe) + 1)
    # 只冻结标题行，不锁住A列，方便横向查看右侧模板。
    worksheet.freeze_panes = 'A2'
    worksheet.auto_filter.ref = f'A1:{template_last_letter}{last_row}'
    worksheet.sheet_view.showGridLines = True
    restore_and_extend_template_conditional_formats(
        worksheet,
        template_conditional_formats,
        len(dataframe) + 1
    )
    # H/I/J 及整行提示均直接使用模板自带规则；不重复追加脚本规则，
    # 避免覆盖用户刚调整的颜色和引用。

    suggested_sheet = workbook.create_sheet('建议剔除')
    write_dataframe_sheet(
        suggested_sheet,
        to_output_layout(suggested, ['建议剔除原因'])
    )

    workbook.calculation.calcMode = 'auto'
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(target_path)

#%% 去重、固定编号、生成建议剔除清单，并将正式模板完整填充到最后一行
source_file = Path(impt) / '股票信息' / '7_已初步清洗股票代码.xlsx'
df_source = pd.read_excel(source_file)
missing_columns = [column for column in BASE_COLUMNS if column not in df_source.columns]
if missing_columns:
    raise KeyError(f'股票信息文件缺少必要列：{missing_columns}')

# 姓名清洗步骤已执行同一规则；这里保留一次防御性兜底，并携带内部内容标记。
df_source[STOCK_CONTENT_MARKER] = stock_row_has_any_value(df_source)
df_source, silent_blank_count, collapsed_all_blank_count = remove_redundant_blank_stock_rows(
    df_source
)

# 四大非股票领域另存，不参与股票序号。
df3 = df_source[BASE_COLUMNS + [STOCK_CONTENT_MARKER]].sort_values(
    by=['机构编号', '分析师姓名'],
    kind='stable'
)
duplicate_count = int(df3.duplicated(subset=BASE_COLUMNS).sum())
df3 = df3.drop_duplicates(subset=BASE_COLUMNS, keep='first').reset_index(drop=True)

df5 = df3[df3['研究领域清洗版'].isin(NON_STOCK_DOMAINS)].copy()
df5 = df5[['分析师姓名', '研究领域清洗版', '公司名称']]
df5.drop_duplicates(subset=['分析师姓名', '公司名称'], inplace=True)
path_sub = Path(impt) / '名单'
path_sub.mkdir(parents=True, exist_ok=True)
df5.to_excel(path_sub / '四大领域分析师.xlsx', index=False)

# 空股票槽位原则上已在清洗姓名步骤处理；此处使用防御性兜底后的结果。
df_stock = df3[~df3['研究领域清洗版'].isin(NON_STOCK_DOMAINS)].copy().reset_index(drop=True)

# 实际股票记录与全空单占位先一次性编号，再进行任何建议剔除。
df_stock.insert(0, '序号', range(1, len(df_stock) + 1))

df_kept, df_suggested, date_start, date_end = split_suggested_exclusions(
    df_stock
)

path3 = Path('输出文件') / '1_数据准备' / '1_股票数据清洗'
path3.mkdir(parents=True, exist_ok=True)
template_path = resolve_template_path()
copy_file(str(template_path), '模板复制.xlsx')

file_path8 = path3 / '1_股票信息部分清洗版.xlsx'
cancel_path = path3 / '取消资格.xlsx'
check_with_retry(str(file_path8))
check_with_retry(str(cancel_path))

build_filled_template_workbook(
    df_stock,
    df_suggested,
    template_path,
    file_path8
)
write_cancel_workbook(df_suggested, cancel_path)

print(f'原始记录：{len(df_source)} 条')
print(f'去除完全重复记录：{duplicate_count} 条')
print(f'四大非股票领域：{len(df5)} 位分析师/公司组合')
print(f'已有内容组内纯空槽位静默剔除：{silent_blank_count} 条（不进取消资格）')
print(f'全空组重复占位合并：{collapsed_all_blank_count} 条（每组只留1条）')
print(f'股票记录编号范围：1-{len(df_stock)}（建议剔除前编号）')
print(f'主表保留全部编号记录：{len(df_stock)} 条')
print(f'建议保留：{len(df_kept)} 条；建议剔除：{len(df_suggested)} 条（未从主表删除）')
print(f'推荐日期建议有效区间：({date_start:%Y-%m-%d}, {date_end:%Y-%m-%d}]')
print(f'模板已填充：A:BM，共 {len(df_stock)} 条主表数据')

open_output(file_path8)
open_output(cancel_path)
open_output('2_手动清洗说明.docx')
insert_bookmark(__file__, 1, rel, start)
