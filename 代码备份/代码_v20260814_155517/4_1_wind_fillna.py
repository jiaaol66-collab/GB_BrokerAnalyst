import sys, time, os, re, warnings
from copy import copy
from pathlib import Path
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

os.chdir(os.path.dirname(__file__))
def find_and_add_package_path(
        package_name: str,
        search_depth: int = 5,
        root_dir: Path = None
) -> bool:
    """
    在文件系统中搜索指定包并添加到 Python 路径

    :param package_name: 要搜索的包名称 (如 'banklib')
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
                (root_dir / up_parents / 'userlib').resolve()
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
            from userlib.tools import *

            print("成功调用工具包函数")
        else:
            print("无法找到包，请检查目录结构")
            time.sleep(2)
            sys.exit(0)

if os.path.basename(os.getcwd()) in ['代码','更新代码']:
    sys.exit(0)

rel = find_rel(__file__, '参评信息')
start = insert_bookmark(__file__,0, rel)
impt = find_rel(__file__, '重要文件')
mtr = find_rel(__file__, '资料')
processed = find_rel(__file__, '已处理')
org = find_rel(__file__, '原始数据集')
yr = get_year(rel)
## 导入...
WORKSPACE_DIR = Path(__file__).resolve().parents[1]
PROJECT_DIR = WORKSPACE_DIR.parent
DATA_DIR = WORKSPACE_DIR / '重要文件' / '分析师清洗版'
RATIO_FILE_PATTERN = re.compile(
    r'^买方反应-比例(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'
)


SOURCE_FILE, DATE_TAG = select_latest_dated_file(
    DATA_DIR,
    RATIO_FILE_PATTERN,
    '买方反应-比例文件',
    int(yr),
)
ANCHOR_FILE = PROJECT_DIR / '已处理' / '2025' / '买方反应_缺失信息股票.xlsx'
DELIVERY_FILE = DATA_DIR / f'买方反应_缺失信息股票{DATE_TAG}.xlsx'

SOURCE_COLUMNS = {
    '股票代码': '标的股票代码',
    '股票名称': '标的股票名称(原)',
    '前一期': '前一观察期',
    '后一期': '后一观察期',
    '前一期机构持股比例': '前一报告期机构持股比例',
    '后一期机构持股比例': '后一报告期机构持股比例',
    '前一期股东总户数': '前一观察期股东总户数',
    '后一期股东总户数': '后一观察期股东总户数',
}


def require_file(path, label):
    if not path.is_file():
        raise FileNotFoundError(f'未找到{label}：{path}')


def get_header_map(worksheet):
    header_map = {}
    for column_index, cell in enumerate(worksheet[1], start=1):
        if cell.value in header_map:
            raise ValueError(f'表头重复：{cell.value}')
        header_map[cell.value] = column_index
    return header_map


def is_zero(value):
    """仅把可解析的数值 0 视为缺失标记；空值和 Excel 错误不算 0。"""
    if value is None or isinstance(value, bool):
        return False
    try:
        return float(value) == 0
    except (TypeError, ValueError):
        return False


def period_sort_value(value):
    if hasattr(value, 'toordinal'):
        return -value.toordinal()
    return str(value)


def collect_missing_rows(source_file):
    workbook = load_workbook(source_file, data_only=True, read_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    header_map = get_header_map(worksheet)

    required_source_headers = set(SOURCE_COLUMNS.values()) | {'看涨/看跌'}
    missing_headers = sorted(required_source_headers - set(header_map))
    if missing_headers:
        workbook.close()
        raise KeyError(f'处理对象缺少列：{missing_headers}')

    selected_count = 0
    seen = set()
    records = []
    group_counts = {
        '机构及股东均缺失': 0,
        '仅机构缺失': 0,
        '仅股东缺失': 0,
    }

    for values in worksheet.iter_rows(min_row=2, values_only=True):
        row = {header: values[column_index - 1] for header, column_index in header_map.items()}
        ins_missing = (
            is_zero(row[SOURCE_COLUMNS['前一期机构持股比例']])
            and is_zero(row[SOURCE_COLUMNS['后一期机构持股比例']])
        )
        shareholder_missing = (
            is_zero(row[SOURCE_COLUMNS['前一期股东总户数']])
            and is_zero(row[SOURCE_COLUMNS['后一期股东总户数']])
        )
        if not (ins_missing or shareholder_missing):
            continue

        selected_count += 1
        core_values = tuple(row[SOURCE_COLUMNS[column]] for column in SOURCE_COLUMNS)
        if core_values in seen:
            continue
        seen.add(core_values)

        direction = row['看涨/看跌']
        records.append({
            '股票代码': row[SOURCE_COLUMNS['股票代码']],
            '股票名称': row[SOURCE_COLUMNS['股票名称']],
            '前一期': row[SOURCE_COLUMNS['前一期']],
            '后一期': row[SOURCE_COLUMNS['后一期']],
            '是否无数据_ins': '无数据' if ins_missing else direction,
            '前一期机构持股比例': (
                None if ins_missing else row[SOURCE_COLUMNS['前一期机构持股比例']]
            ),
            '后一期机构持股比例': (
                None if ins_missing else row[SOURCE_COLUMNS['后一期机构持股比例']]
            ),
            '是否无数据_shareholder': '无数据' if shareholder_missing else direction,
            '前一期股东总户数': (
                None if shareholder_missing else row[SOURCE_COLUMNS['前一期股东总户数']]
            ),
            '后一期股东总户数': (
                None if shareholder_missing else row[SOURCE_COLUMNS['后一期股东总户数']]
            ),
        })

        if ins_missing and shareholder_missing:
            group_counts['机构及股东均缺失'] += 1
        elif ins_missing:
            group_counts['仅机构缺失'] += 1
        else:
            group_counts['仅股东缺失'] += 1

    workbook.close()
    records.sort(
        key=lambda record: (
            str(record['股票代码']),
            period_sort_value(record['前一期']),
            str(record['股票名称']),
        )
    )
    return records, selected_count, group_counts


def build_output(anchor_file, output_file, records):
    workbook = load_workbook(anchor_file)
    worksheet = workbook[workbook.sheetnames[0]]
    anchor_headers = [cell.value for cell in worksheet[1]]
    if not anchor_headers or anchor_headers[0] != '标识':
        workbook.close()
        raise ValueError(f'锚点首列应为辅助列“标识”，实际为：{anchor_headers[:1]}')

    output_headers = anchor_headers[1:]
    expected_headers = [
        '股票代码', '股票名称', '前一期', '后一期',
        '是否无数据_ins', '前一期机构持股比例', '后一期机构持股比例',
        '是否无数据_shareholder', '前一期股东总户数', '后一期股东总户数',
    ]
    if output_headers != expected_headers:
        workbook.close()
        raise ValueError(
            '锚点列名或顺序与预期不一致：\n'
            f'实际：{output_headers}\n预期：{expected_headers}'
        )

    header_templates = [copy(worksheet.cell(1, column)._style) for column in range(2, 12)]
    row_templates = [copy(worksheet.cell(2, column)._style) for column in range(2, 12)]
    source_widths = [worksheet.column_dimensions[get_column_letter(column)].width for column in range(2, 12)]
    template_row_height = worksheet.row_dimensions[2].height

    if worksheet.max_row > 1:
        worksheet.delete_rows(2, worksheet.max_row - 1)
    worksheet.delete_cols(1)

    for column, header in enumerate(output_headers, start=1):
        cell = worksheet.cell(1, column, header)
        cell._style = header_templates[column - 1]
        worksheet.column_dimensions[get_column_letter(column)].width = source_widths[column - 1]

    for row_number, record in enumerate(records, start=2):
        for column, header in enumerate(output_headers, start=1):
            cell = worksheet.cell(row_number, column, record[header])
            cell._style = copy(row_templates[column - 1])
        if template_row_height is not None:
            worksheet.row_dimensions[row_number].height = template_row_height

    worksheet.auto_filter.ref = f'A1:J{len(records) + 1}'
    workbook.save(output_file)
    workbook.close()


"""""""""
   主体
"""""""""
require_file(SOURCE_FILE, '处理对象')
require_file(ANCHOR_FILE, '2025锚点文件')

missing_rows, raw_selected_count, missing_group_counts = collect_missing_rows(SOURCE_FILE)
DELIVERY_FILE.parent.mkdir(parents=True, exist_ok=True)
build_output(ANCHOR_FILE, DELIVERY_FILE, missing_rows)
"""""""""
   主体
"""""""""

## 导出...
print(f'处理对象：{SOURCE_FILE}')
print(f'原始命中记录：{raw_selected_count}')
print(f'去重后记录：{len(missing_rows)}')
for label, count in missing_group_counts.items():
    print(f'{label}：{count}')
print(f'填充文件：{DELIVERY_FILE}')
print('success')

# 添加时间记录
insert_bookmark(__file__, 1, rel, start)
if os.environ.get('GB_SKIP_FLUSH_BUFFER') != '1':
    open_file('flush_buffer.py')
