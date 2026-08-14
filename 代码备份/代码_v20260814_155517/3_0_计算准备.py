import sys, os, re, time
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

from copy import deepcopy

from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)

rel = find_rel(__file__, '参评信息')
start = insert_bookmark(__file__,0, rel)
impt = find_rel(__file__, '重要文件')
mtr = find_rel(__file__, '资料')
processed = find_rel(__file__, '已处理')
org = find_rel(__file__, '原始数据集')
yr = get_year(rel)

INITIAL_DIR = Path('2_数据已提取') / '1_股票代码清洗与数据提取'
INITIAL_TEMPLATE_SOURCE = INITIAL_DIR / '分析师定量数据.xlsx'
INITIAL_EXTRACTED_SOURCE = INITIAL_DIR / '分析师定量数据_已提取.xlsx'

UPDATE_DIR = Path('3_万得提取')
UPDATE_TEMPLATE_SOURCE = UPDATE_DIR / '分析师定量数据.xlsx'
EXTRACTED_FILE_PATTERN = re.compile(
    r'^分析师已提取(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'
)
UPDATE_EXTRACTED_CANDIDATES = [
    path
    for path in UPDATE_DIR.glob('分析师已提取*.xlsx')
    if EXTRACTED_FILE_PATTERN.fullmatch(path.name)
]

PROCESS_INITIAL_EXTRACTION = custom_dialog(
    '选择处理类型',
    '请选择本次处理类型：\n提取：处理 2_数据已提取\n更新：处理 3_万得提取',
    'confirm',
    '提取',
    '更新',
)
if PROCESS_INITIAL_EXTRACTION is None:
    print('未选择处理类型，程序已退出。')
    sys.exit(0)

if PROCESS_INITIAL_EXTRACTION:
    EXTRACTION_STAGE = '首次提取'
    TEMPLATE_SOURCE = INITIAL_TEMPLATE_SOURCE
    EXTRACTED_SOURCE = INITIAL_EXTRACTED_SOURCE
    FINAL_DATE_TAG = time.strftime(
        '%m%d',
        time.localtime(EXTRACTED_SOURCE.stat().st_mtime),
    )
    WIND_EXTRACTION_LAST_HEADER = '股东比例变化'
    INDUSTRY_RETURN_LAST_HEADER = '超额回报率考察期'
    SOURCE_ALIGNMENT_LAST_HEADER = '超额回报率考察期'
    INCLUDE_PUBLIC_FUND = True
    EXPECTED_HEADERS = {
        'A': '序号',
        'B': '分析师姓名',
        'F': '研究领域清洗版',
        'H': '标的股票代码',
        'AU': '超额回报率六个月',
        'BB': '超额回报率考察期',
    }
else:
    if not UPDATE_EXTRACTED_CANDIDATES:
        raise FileNotFoundError(
            '选择了“更新”，但 3_万得提取 中未找到分析师已提取日期.xlsx'
        )
    EXTRACTION_STAGE = '后续更新'
    TEMPLATE_SOURCE = UPDATE_TEMPLATE_SOURCE
    EXTRACTED_SOURCE = max(
        UPDATE_EXTRACTED_CANDIDATES,
        key=lambda path: (
            EXTRACTED_FILE_PATTERN.fullmatch(path.name).group('date'),
            path.stat().st_mtime,
        ),
    )
    FINAL_DATE_TAG = EXTRACTED_FILE_PATTERN.fullmatch(
        EXTRACTED_SOURCE.name
    ).group('date')
    WIND_EXTRACTION_LAST_HEADER = '股东比例变化'
    INDUSTRY_RETURN_LAST_HEADER = '超额回报率考察期'
    SOURCE_ALIGNMENT_LAST_HEADER = '股东比例变化'
    INCLUDE_PUBLIC_FUND = False
    EXPECTED_HEADERS = {
        'A': '序号',
        'B': '分析师姓名',
        'F': '研究领域清洗版',
        'H': '标的股票代码',
        'BB': '超额回报率考察期',
        'BC': '前一观察期',
        'BM': '股东比例变化',
    }

FORMULA_TEMPLATE_DIR = Path(org) / '通用'
BUYER_REACTION_TEMPLATE = FORMULA_TEMPLATE_DIR / '分析师买方反应公式版.xlsx'
PUBLIC_FUND_TEMPLATE = FORMULA_TEMPLATE_DIR / '分析师公募基金公式版.xlsx'

# Wind 已提取版本使用了更明确的“进入数目”表头；公式模板和最终分裂文件
# 继续保留原表头。这里只兼容名称，不改变列位置、数据或公式。
SOURCE_HEADER_ALIASES = {
    '头部公募基金数量_前一期': '头部公募基金进入数目_前一期',
    '头部公募基金数量_后一期': '头部公募基金进入数目_后一期',
    '头部公募基金变化': '头部公募基金进入数目变化',
}
HEADER_CANONICAL_NAMES = {
    source_header: template_header
    for template_header, source_header in SOURCE_HEADER_ALIASES.items()
}

OUTPUT_PATHS = {}
if EXTRACTION_STAGE == '首次提取':
    OUTPUT_PATHS['template'] = (
        Path(impt) / '提取模板' / '分析师回报率提取模版.xlsx'
    )
OUTPUT_PATHS.update({
    'industry': Path(impt) / '同行业回报率' / f'同行业回报率{FINAL_DATE_TAG}.xlsx',
    'buyer_ratio': Path(impt) / '分析师清洗版' / f'买方反应-比例{FINAL_DATE_TAG}.xlsx',
})
if INCLUDE_PUBLIC_FUND:
    OUTPUT_PATHS['buyer_fund'] = (
        Path(impt) / '分析师清洗版' / f'买方反应-公募基金{FINAL_DATE_TAG}.xlsx'
    )


def require_source_file(path):
    if not path.is_file():
        raise FileNotFoundError(f'未找到输入文件：{path.resolve()}')


def canonicalize_header(header):
    """将已提取版本的新表头还原为公式模板使用的标准表头。"""
    return HEADER_CANONICAL_NAMES.get(header, header)


def resolve_source_header(source_headers, output_header):
    """按原表头优先、已提取版本别名次之，定位源数据列。"""
    if output_header in source_headers:
        return output_header
    alias = SOURCE_HEADER_ALIASES.get(output_header)
    if alias in source_headers:
        return alias
    return None


def validate_source_alignment(template_path, extracted_path):
    """确认最新模板与 Wind 已提取版本在提取范围内按列名完全一致。"""
    workbooks = []
    try:
        for path in (template_path, extracted_path):
            workbook = load_workbook(
                path,
                data_only=False,
                read_only=True,
                keep_links=True,
            )
            workbooks.append(workbook)
        template_sheet = workbooks[0][workbooks[0].sheetnames[0]]
        extracted_sheet = workbooks[1][workbooks[1].sheetnames[0]]
        template_headers_all = [
            template_sheet.cell(1, column).value
            for column in range(1, template_sheet.max_column + 1)
        ]
        try:
            last_column = template_headers_all.index(
                SOURCE_ALIGNMENT_LAST_HEADER
            ) + 1
        except ValueError as error:
            raise KeyError(
                f'{template_path} 缺少分解边界列：'
                f'{SOURCE_ALIGNMENT_LAST_HEADER}'
            ) from error
        if extracted_sheet.max_column < last_column:
            raise ValueError(
                f'{extracted_path} 至少应提取到“{SOURCE_ALIGNMENT_LAST_HEADER}”'
                f'（{get_column_letter(last_column)}列），实际末列为 '
                f'{get_column_letter(extracted_sheet.max_column)}列'
            )
        template_headers = [
            template_sheet.cell(1, column).value
            for column in range(1, last_column + 1)
        ]
        extracted_headers = [
            extracted_sheet.cell(1, column).value
            for column in range(1, last_column + 1)
        ]
        canonical_template_headers = [
            canonicalize_header(header) for header in template_headers
        ]
        canonical_extracted_headers = [
            canonicalize_header(header) for header in extracted_headers
        ]
        if canonical_template_headers != canonical_extracted_headers:
            differences = [
                f'{get_column_letter(column)}: '
                f'模板={template_headers[column - 1]!r}, '
                f'已提取={extracted_headers[column - 1]!r}'
                for column in range(1, last_column + 1)
                if canonical_template_headers[column - 1]
                != canonical_extracted_headers[column - 1]
            ]
            raise ValueError(
                '公式模板与已提取版本列名不一致：\n  - '
                + '\n  - '.join(differences)
            )
        if template_sheet.max_row != extracted_sheet.max_row:
            raise ValueError(
                '公式模板与已提取版本主表行数不一致：'
                f'{template_sheet.max_row} != {extracted_sheet.max_row}'
            )
    finally:
        for workbook in workbooks:
            workbook.close()


def validate_main_sheet(worksheet, source_path):
    """按最新模板表头校验当前工作表实际包含的关键列。"""
    errors = []
    for column_letter, expected_header in EXPECTED_HEADERS.items():
        if column_index_from_string(column_letter) > worksheet.max_column:
            continue
        actual_header = worksheet[f'{column_letter}1'].value
        if canonicalize_header(actual_header) != canonicalize_header(expected_header):
            errors.append(
                f'{column_letter}列应为“{expected_header}”，实际为“{actual_header}”'
            )
    if errors:
        details = '\n'.join(f'  - {error}' for error in errors)
        raise ValueError(f'{source_path} 表头与0805模板不一致：\n{details}')


def count_formulas(worksheet, max_column):
    return sum(
        1
        for row in worksheet.iter_rows(
            min_row=1,
            max_row=worksheet.max_row,
            min_col=1,
            max_col=max_column,
        )
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith('=')
    )


def get_header_column(worksheet, header, source_path):
    """按列名定位分解边界，避免模板增列后继续依赖旧列字母。"""
    matches = [
        column
        for column in range(1, worksheet.max_column + 1)
        if canonicalize_header(worksheet.cell(1, column).value)
        == canonicalize_header(header)
    ]
    if not matches:
        raise KeyError(f'{source_path} 缺少列：{header}')
    if len(matches) > 1:
        raise ValueError(f'{source_path} 存在重复列名：{header}')
    return matches[0]


def is_wind_formula(value):
    """Wind 公式在工作簿中保存为带外部名称空间的 [编号]!函数。"""
    return (
        isinstance(value, str)
        and value.startswith('=')
        and re.search(r'\[[^\]]+\]!', value) is not None
    )


def retain_conditional_formats_through(worksheet, max_column):
    """保留并裁剪 A 至指定末列内的原模板条件格式。"""
    retained = []
    for conditional_range in worksheet.conditional_formatting:
        clipped_parts = []
        for range_part in str(conditional_range.sqref).split():
            min_col, min_row, max_col, max_row = range_boundaries(range_part)
            if min_col > max_column:
                continue
            max_col = min(max_col, max_column)
            clipped_parts.append(
                f'{get_column_letter(min_col)}{min_row}:'
                f'{get_column_letter(max_col)}{max_row}'
            )
        if clipped_parts:
            rules = [
                deepcopy(rule)
                for rule in worksheet.conditional_formatting[conditional_range]
            ]
            retained.append((' '.join(clipped_parts), rules))

    worksheet.conditional_formatting = ConditionalFormattingList()
    for target_range, rules in retained:
        for rule in rules:
            worksheet.conditional_formatting.add(target_range, rule)


def keep_main_sheet_columns(source_path, target_path, last_header, require_formulas):
    """从完整工作簿按末列列名保留主表；模板以公式模式读取。"""
    workbook = load_workbook(source_path, data_only=False)
    worksheet = workbook[workbook.sheetnames[0]]
    validate_main_sheet(worksheet, source_path)

    last_column = get_header_column(worksheet, last_header, source_path)
    last_column_letter = get_column_letter(last_column)
    source_formula_count = count_formulas(worksheet, last_column)
    if require_formulas and source_formula_count == 0:
        raise ValueError(f'{source_path} 的 A:{last_column_letter} 未检测到公式')

    for extra_sheet in list(workbook.worksheets)[1:]:
        workbook.remove(extra_sheet)
    if worksheet.max_column > last_column:
        worksheet.delete_cols(
            last_column + 1,
            worksheet.max_column - last_column,
        )
    retain_conditional_formats_through(worksheet, last_column)
    worksheet.freeze_panes = 'A2'
    worksheet.auto_filter.ref = f'A1:{last_column_letter}{worksheet.max_row}'
    workbook.calculation.calcMode = 'auto'
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True

    target_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(target_path)

    saved_workbook = load_workbook(target_path, data_only=False, read_only=True)
    saved_worksheet = saved_workbook[saved_workbook.sheetnames[0]]
    saved_formula_count = count_formulas(saved_worksheet, last_column)
    saved_workbook.close()
    workbook.close()
    if saved_formula_count != source_formula_count:
        raise RuntimeError(
            f'{target_path} 公式数量校验失败：'
            f'{source_formula_count} -> {saved_formula_count}'
        )


def make_header_map(worksheet, source_path):
    """建立唯一列名到列号的映射，避免重名列静默错配。"""
    header_map = {}
    duplicates = []
    for column in range(1, worksheet.max_column + 1):
        header = worksheet.cell(1, column).value
        if header in header_map:
            duplicates.append(header)
        else:
            header_map[header] = column
    if duplicates:
        raise ValueError(f'{source_path} 存在重复列名：{duplicates}')
    return header_map


def build_formula_preserving_workbook(
        layout_path,
        extracted_path,
        fallback_path,
        target_path,
        last_header=None,
):
    """
    以指定版式按同名列分解：Wind 公式替换为已提取值，其他公式原样保留。

    已提取文件没有的非公式列回到最新“分析师定量数据”取值；这主要用于
    头部公募基金分块中尚待后续提取的空白列。
    """
    print(f'    正在载入：{target_path.name}', flush=True)
    layout_workbook = load_workbook(
        layout_path,
        data_only=False,
        keep_links=True,
    )
    extracted_workbook = load_workbook(
        extracted_path,
        data_only=True,
        read_only=True,
        keep_links=True,
    )
    fallback_workbook = load_workbook(
        fallback_path,
        data_only=False,
        read_only=True,
        keep_links=True,
    )
    layout_worksheet = layout_workbook[layout_workbook.sheetnames[0]]
    extracted_worksheet = extracted_workbook[extracted_workbook.sheetnames[0]]
    fallback_worksheet = fallback_workbook[fallback_workbook.sheetnames[0]]
    validate_main_sheet(extracted_worksheet, extracted_path)

    row_counts = {
        layout_path: layout_worksheet.max_row,
        extracted_path: extracted_worksheet.max_row,
        fallback_path: fallback_worksheet.max_row,
    }
    if len(set(row_counts.values())) != 1:
        details = '，'.join(f'{path}={rows}' for path, rows in row_counts.items())
        raise ValueError(f'分解来源主表行数不一致：{details}')

    if last_header is None:
        last_column = layout_worksheet.max_column
    else:
        last_column = get_header_column(
            layout_worksheet,
            last_header,
            layout_path,
        )
    output_headers = [
        layout_worksheet.cell(1, column).value
        for column in range(1, last_column + 1)
    ]
    make_header_map(layout_worksheet, layout_path)
    extracted_headers = make_header_map(extracted_worksheet, extracted_path)
    fallback_headers = make_header_map(fallback_worksheet, fallback_path)

    extracted_header_matches = {
        header: resolve_source_header(extracted_headers, header)
        for header in output_headers
    }
    fallback_header_matches = {
        header: resolve_source_header(fallback_headers, header)
        for header in output_headers
    }

    retained_formulas = {}
    wind_formula_columns = set()
    for row in layout_worksheet.iter_rows(
        min_row=2,
        max_row=layout_worksheet.max_row,
        min_col=1,
        max_col=last_column,
    ):
        for cell in row:
            if not (isinstance(cell.value, str) and cell.value.startswith('=')):
                continue
            if is_wind_formula(cell.value):
                wind_formula_columns.add(cell.column)
            else:
                retained_formulas[cell.coordinate] = cell.value

    missing_wind_headers = [
        output_headers[column - 1]
        for column in sorted(wind_formula_columns)
        if extracted_header_matches[output_headers[column - 1]] is None
    ]
    if missing_wind_headers:
        raise KeyError(
            f'{extracted_path} 缺少 Wind 公式结果列：{missing_wind_headers}'
        )

    total_data_rows = extracted_worksheet.max_row - 1
    extracted_rows = extracted_worksheet.iter_rows(min_row=2, values_only=True)
    fallback_rows = fallback_worksheet.iter_rows(min_row=2, values_only=True)
    replaced_wind_formulas = 0
    for row_index, (extracted_row, fallback_row) in enumerate(
            zip(extracted_rows, fallback_rows),
            start=2,
    ):
        for target_column, header in enumerate(output_headers, 1):
            target_cell = layout_worksheet.cell(row_index, target_column)
            if (
                    isinstance(target_cell.value, str)
                    and target_cell.value.startswith('=')
                    and not is_wind_formula(target_cell.value)
            ):
                continue

            source_header = extracted_header_matches[header]
            if source_header is not None:
                source_column = extracted_headers[source_header]
                if is_wind_formula(target_cell.value):
                    replaced_wind_formulas += 1
                target_cell.value = extracted_row[source_column - 1]
                continue

            fallback_header = fallback_header_matches[header]
            if fallback_header is None:
                raise KeyError(
                    f'{extracted_path} 和 {fallback_path} 均缺少列：{header}'
                )
            fallback_column = fallback_headers[fallback_header]
            target_cell.value = fallback_row[fallback_column - 1]

        completed_rows = row_index - 1
        if completed_rows % 250 == 0 or completed_rows == total_data_rows:
            print(
                f'\r    数据回填：{completed_rows}/{total_data_rows} 行',
                end='\n' if completed_rows == total_data_rows else '',
                flush=True,
            )

    for extra_sheet in list(layout_workbook.worksheets)[1:]:
        layout_workbook.remove(extra_sheet)
    if layout_worksheet.max_column > last_column:
        layout_worksheet.delete_cols(
            last_column + 1,
            layout_worksheet.max_column - last_column,
        )
    retain_conditional_formats_through(layout_worksheet, last_column)
    layout_worksheet.freeze_panes = 'A2'
    last_column_letter = get_column_letter(last_column)
    layout_worksheet.auto_filter.ref = (
        f'A1:{last_column_letter}{layout_worksheet.max_row}'
    )
    layout_workbook.calculation.calcMode = 'auto'
    layout_workbook.calculation.fullCalcOnLoad = True
    layout_workbook.calculation.forceFullCalc = True

    target_path.parent.mkdir(parents=True, exist_ok=True)
    print(
        f'    Wind 公式结果化：{replaced_wind_formulas} 个；'
        f'保留其他公式：{len(retained_formulas)} 个',
        flush=True,
    )
    print('    正在保存并校验公式...', flush=True)
    layout_workbook.save(target_path)
    layout_workbook.close()
    extracted_workbook.close()
    fallback_workbook.close()

    saved_workbook = load_workbook(
        target_path,
        data_only=False,
        read_only=True,
        keep_links=True,
    )
    saved_worksheet = saved_workbook[saved_workbook.sheetnames[0]]
    saved_formulas = {}
    saved_wind_formulas = []
    for row in saved_worksheet.iter_rows(
        min_row=2,
        max_row=saved_worksheet.max_row,
        min_col=1,
        max_col=saved_worksheet.max_column,
    ):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith('='):
                saved_formulas[cell.coordinate] = cell.value
                if is_wind_formula(cell.value):
                    saved_wind_formulas.append(cell.coordinate)
    saved_workbook.close()
    if saved_wind_formulas:
        raise ValueError(
            f'{target_path} 仍存在 Wind 公式：{saved_wind_formulas[:10]}'
        )
    if saved_formulas != retained_formulas:
        missing_or_changed = sorted(
            coordinate
            for coordinate in set(saved_formulas) | set(retained_formulas)
            if saved_formulas.get(coordinate) != retained_formulas.get(coordinate)
        )
        raise ValueError(
            f'{target_path} 非 Wind 公式未原样保留：'
            f'{missing_or_changed[:10]}'
        )


def run_with_progress(step_number, total_steps, description, function, *args, **kwargs):
    """显示每个分块任务的开始、完成状态和耗时。"""
    prefix = f'[{step_number}/{total_steps}]'
    print(f'{prefix} 开始：{description}')
    step_start = time.time()
    function(*args, **kwargs)
    elapsed = time.time() - step_start
    print(f'{prefix} 完成：{description}（{elapsed:.1f} 秒）')


require_source_file(TEMPLATE_SOURCE)
require_source_file(EXTRACTED_SOURCE)
require_source_file(BUYER_REACTION_TEMPLATE)
if INCLUDE_PUBLIC_FUND:
    require_source_file(PUBLIC_FUND_TEMPLATE)
validate_source_alignment(TEMPLATE_SOURCE, EXTRACTED_SOURCE)
total_steps = 4 if INCLUDE_PUBLIC_FUND else 2
print(f'提取阶段：{EXTRACTION_STAGE}')
print(f'本阶段模板：{TEMPLATE_SOURCE.resolve()}')
print(f'本阶段已提取文件：{EXTRACTED_SOURCE.resolve()}')
print(f'日期标签：{FINAL_DATE_TAG}')
print(f'公募基金分解：{"是" if INCLUDE_PUBLIC_FUND else "否"}')

# Wind 提取模板只在首次提取时生成；后续更新直接复用首次生成的模板。
if EXTRACTION_STAGE == '首次提取':
    run_with_progress(
        1,
        total_steps,
        '生成分析师回报率提取模版（至“股东比例变化”，保留全部公式）',
        keep_main_sheet_columns,
        TEMPLATE_SOURCE,
        OUTPUT_PATHS['template'],
        WIND_EXTRACTION_LAST_HEADER,
        require_formulas=True,
    )

# 同行业回报率：截取至“超额回报率考察期”，只结果化 Wind 公式。
if EXTRACTION_STAGE == '首次提取':
    run_with_progress(
        2,
        total_steps,
        f'生成同行业回报率{FINAL_DATE_TAG}（首次提取，保留非 Wind 公式）',
        build_formula_preserving_workbook,
        TEMPLATE_SOURCE,
        EXTRACTED_SOURCE,
        TEMPLATE_SOURCE,
        OUTPUT_PATHS['industry'],
        INDUSTRY_RETURN_LAST_HEADER,
    )
else:
    run_with_progress(
        1,
        total_steps,
        f'生成同行业回报率{FINAL_DATE_TAG}（后续更新，保留非 Wind 公式）',
        build_formula_preserving_workbook,
        TEMPLATE_SOURCE,
        EXTRACTED_SOURCE,
        TEMPLATE_SOURCE,
        OUTPUT_PATHS['industry'],
        INDUSTRY_RETURN_LAST_HEADER,
    )

# 买方反应：按通用公式版 A:R 分解，Wind 公式结果化，其他公式保留。
run_with_progress(
    3 if INCLUDE_PUBLIC_FUND else 2,
    total_steps,
    f'生成买方反应-比例{FINAL_DATE_TAG}（公式版 A:R）',
    build_formula_preserving_workbook,
    BUYER_REACTION_TEMPLATE,
    EXTRACTED_SOURCE,
    TEMPLATE_SOURCE,
    OUTPUT_PATHS['buyer_ratio'],
)

if INCLUDE_PUBLIC_FUND:
    run_with_progress(
        4,
        total_steps,
        f'生成买方反应-公募基金{FINAL_DATE_TAG}（首次提取）',
        build_formula_preserving_workbook,
        PUBLIC_FUND_TEMPLATE,
        EXTRACTED_SOURCE,
        TEMPLATE_SOURCE,
        OUTPUT_PATHS['buyer_fund'],
    )

for output_name, output_path in OUTPUT_PATHS.items():
    print(f'{output_name}: {output_path.resolve()}')

insert_bookmark(__file__, 1, rel, start)

