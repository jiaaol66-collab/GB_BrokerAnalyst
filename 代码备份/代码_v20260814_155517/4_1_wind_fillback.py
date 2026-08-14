import sys, time, os, warnings
import hashlib
import math
import re
import shutil
from datetime import date, datetime
from pathlib import Path
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.utils.datetime import from_excel

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
DATA_DIR = WORKSPACE_DIR / '重要文件' / '分析师清洗版'
BACKUP_DIR = DATA_DIR / '备份'

RATIO_FILE_PATTERN = re.compile(
    r'^买方反应-比例(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'
)
MISSING_FILE_PATTERN = re.compile(
    r'^买方反应[-_]缺失信息(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'
)

RATIO_HEADERS = {
    '股票代码': '标的股票代码',
    '股票名称': '标的股票名称(原)',
    '前一期': '前一观察期',
    '后一期': '后一观察期',
    '前一期机构持股比例': '前一报告期机构持股比例',
    '后一期机构持股比例': '后一报告期机构持股比例',
    '前一期股东总户数': '前一观察期股东总户数',
    '后一期股东总户数': '后一观察期股东总户数',
}

MISSING_HEADERS = {
    '股票代码',
    '前一期',
    '后一期',
    '前一期机构持股比例',
    '后一期机构持股比例',
    '前一期股东总户数',
    '后一期股东总户数',
}

BACKFILL_FIELDS = (
    ('前一期', '前一期机构持股比例', '前一期机构持股比例'),
    ('后一期', '后一期机构持股比例', '后一期机构持股比例'),
    ('前一期', '前一期股东总户数', '前一期股东总户数'),
    ('后一期', '后一期股东总户数', '后一期股东总户数'),
)


def get_header_map(worksheet):
    header_map = {}
    for column_index, cell in enumerate(worksheet[1], start=1):
        header = cell.value
        if header in header_map:
            raise ValueError(f'{worksheet.title} 表头重复：{header}')
        header_map[header] = column_index
    return header_map


def require_headers(header_map, required_headers, label):
    missing = sorted(set(required_headers) - set(header_map))
    if missing:
        raise KeyError(f'{label}缺少列：{missing}')


def normalize_code(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return str(value).upper()
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    normalized = str(value).strip().upper()
    return normalized or None


def normalize_period(value, label):
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        converted = from_excel(value)
        return converted.date() if isinstance(converted, datetime) else converted
    if isinstance(value, str):
        value = value.strip()
        for date_format in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%Y/%m/%d', '%Y%m%d'):
            try:
                return datetime.strptime(value, date_format).date()
            except ValueError:
                pass
    raise ValueError(f'{label}不是可识别的日期：{value!r}')


def is_blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def is_missing_target(value):
    if is_blank(value):
        return True
    if isinstance(value, bool):
        return False
    try:
        return float(value) == 0
    except (TypeError, ValueError):
        return False


def values_equal(left, right):
    if isinstance(left, (int, float)) and not isinstance(left, bool):
        if isinstance(right, (int, float)) and not isinstance(right, bool):
            return math.isclose(float(left), float(right), rel_tol=0, abs_tol=1e-12)
    return left == right


def add_source_value(mapping, key, value, field_name, row_number):
    if is_blank(value):
        return
    if key in mapping:
        previous_value, previous_row = mapping[key]
        if not values_equal(previous_value, value):
            raise ValueError(
                f'缺失信息存在冲突：{field_name}，代码/考察期={key}，'
                f'第{previous_row}行={previous_value!r}，第{row_number}行={value!r}'
            )
        return
    mapping[key] = (value, row_number)


def build_source_maps(missing_file):
    workbook = load_workbook(missing_file, data_only=True, read_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    header_map = get_header_map(worksheet)
    require_headers(header_map, MISSING_HEADERS, '缺失信息文件')

    mappings = {field_name: {} for _, field_name, _ in BACKFILL_FIELDS}
    source_row_count = 0
    for row_number in range(2, worksheet.max_row + 1):
        code = normalize_code(
            worksheet.cell(row_number, header_map['股票代码']).value
        )
        row_has_value = any(
            not is_blank(worksheet.cell(row_number, header_map[header]).value)
            for header in MISSING_HEADERS
        )
        if not row_has_value:
            continue
        if code is None:
            workbook.close()
            raise ValueError(f'缺失信息文件第{row_number}行缺少股票代码')
        source_row_count += 1

        periods = {}
        for period_field in ('前一期', '后一期'):
            periods[period_field] = normalize_period(
                worksheet.cell(row_number, header_map[period_field]).value,
                f'缺失信息文件第{row_number}行{period_field}',
            )

        for period_field, source_field, mapping_name in BACKFILL_FIELDS:
            value = worksheet.cell(row_number, header_map[source_field]).value
            if is_blank(value):
                continue
            period = periods[period_field]
            if period is None:
                workbook.close()
                raise ValueError(
                    f'缺失信息文件第{row_number}行{source_field}有值，但{period_field}为空'
                )
            add_source_value(
                mappings[mapping_name],
                (code, period),
                value,
                source_field,
                row_number,
            )

    workbook.close()
    return mappings, source_row_count


def workbook_value_snapshot(workbook):
    snapshot = {}
    for worksheet in workbook.worksheets:
        for row in worksheet.iter_rows():
            for cell in row:
                snapshot[(worksheet.title, cell.coordinate)] = cell.value
    return snapshot


def workbook_formula_snapshot(workbook):
    return {
        (worksheet.title, cell.coordinate): cell.value
        for worksheet in workbook.worksheets
        for row in worksheet.iter_rows()
        for cell in row
        if cell.data_type == 'f'
    }


def workbook_structure_snapshot(workbook):
    return [
        {
            'title': worksheet.title,
            'max_row': worksheet.max_row,
            'max_column': worksheet.max_column,
            'freeze_panes': str(worksheet.freeze_panes) if worksheet.freeze_panes else None,
            'auto_filter': worksheet.auto_filter.ref,
            'merged_ranges': tuple(sorted(str(item) for item in worksheet.merged_cells.ranges)),
            'conditional_formats': tuple(
                sorted(str(item.sqref) for item in worksheet.conditional_formatting)
            ),
            'tables': tuple(sorted(worksheet.tables.keys())),
        }
        for worksheet in workbook.worksheets
    ]


def file_sha256(path):
    digest = hashlib.sha256()
    with path.open('rb') as source:
        for block in iter(lambda: source.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def verify_saved_output(
        output_file,
        values_before,
        formulas_before,
        structure_before,
        changed_values,
):
    workbook = load_workbook(output_file, data_only=False, read_only=False)
    values_after = workbook_value_snapshot(workbook)
    formulas_after = workbook_formula_snapshot(workbook)
    structure_after = workbook_structure_snapshot(workbook)

    if formulas_after != formulas_before:
        workbook.close()
        raise RuntimeError('保存后公式发生变化，已拒绝交付输出文件')
    if structure_after != structure_before:
        workbook.close()
        raise RuntimeError('保存后工作表结构发生变化，已拒绝交付输出文件')

    changed_keys = set(changed_values)
    all_keys = set(values_before) | set(values_after)
    for key in all_keys:
        before = values_before.get(key)
        after = values_after.get(key)
        if key in changed_keys:
            expected = changed_values[key]
            if not values_equal(after, expected):
                workbook.close()
                raise RuntimeError(
                    f'保存后回填值校验失败：{key}，预期={expected!r}，实际={after!r}'
                )
        elif not values_equal(before, after):
            workbook.close()
            raise RuntimeError(
                f'检测到回填区域外的值发生变化：{key}，'
                f'原值={before!r}，新值={after!r}'
            )
    workbook.close()


def show_work_status(message, finished=False):
    """在终端同一行刷新当前工作状态。"""
    print(
        f'\r正在工作：{message:<50}',
        end='\n' if finished else '',
        flush=True,
    )


def backfill_ratio_file(ratio_file, missing_file, output_file, backup_dir):
    show_work_status('正在备份源比例文件')
    source_hash_before = file_sha256(ratio_file)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    backup_dir.mkdir(parents=True, exist_ok=True)
    source_backup = backup_dir / (
        f'{ratio_file.stem}_回填前备份_{timestamp}{ratio_file.suffix}'
    )
    shutil.copy2(ratio_file, source_backup)
    if file_sha256(source_backup) != source_hash_before:
        raise RuntimeError(f'源文件备份校验失败：{source_backup}')

    show_work_status('正在读取缺失信息文件')
    mappings, source_row_count = build_source_maps(missing_file)
    show_work_status('正在载入比例文件及公式缓存')
    workbook = load_workbook(ratio_file, data_only=False, read_only=False)
    # read_only 工作表不适合按行号随机访问；这里不写入，只在内存中读取公式缓存值。
    cached_workbook = load_workbook(ratio_file, data_only=True, read_only=False)
    worksheet = workbook[workbook.sheetnames[0]]
    cached_worksheet = cached_workbook[worksheet.title]
    header_map = get_header_map(worksheet)
    cached_header_map = get_header_map(cached_worksheet)
    require_headers(header_map, RATIO_HEADERS.values(), '比例文件')
    require_headers(cached_header_map, RATIO_HEADERS.values(), '比例文件缓存')

    show_work_status('正在建立回填前校验快照')
    values_before = workbook_value_snapshot(workbook)
    formulas_before = workbook_formula_snapshot(workbook)
    structure_before = workbook_structure_snapshot(workbook)
    changed_values = {}
    changes_by_field = defaultdict(int)
    matched_source_keys = {field_name: set() for field_name in mappings}
    protected_non_missing = defaultdict(int)
    source_codes = {
        key[0]
        for field_mapping in mappings.values()
        for key in field_mapping
    }

    total_rows = worksheet.max_row - 1
    for row_number in range(2, worksheet.max_row + 1):
        completed_rows = row_number - 1
        if completed_rows % 250 == 0 or completed_rows == total_rows:
            show_work_status(f'正在回填 {completed_rows}/{total_rows} 行')
        code = normalize_code(
            worksheet.cell(row_number, header_map[RATIO_HEADERS['股票代码']]).value
        )
        if code is None or code not in source_codes:
            continue

        periods = {}
        for period_field in ('前一期', '后一期'):
            period_header = RATIO_HEADERS[period_field]
            cached_value = cached_worksheet.cell(
                row_number, cached_header_map[period_header]
            ).value
            periods[period_field] = normalize_period(
                cached_value,
                f'比例文件第{row_number}行{period_header}缓存值',
            )
            if periods[period_field] is None:
                workbook.close()
                cached_workbook.close()
                raise ValueError(
                    f'比例文件第{row_number}行{period_header}公式没有缓存结果，'
                    '请先用 Excel 打开并保存比例文件后重试'
                )

        for period_field, source_field, mapping_name in BACKFILL_FIELDS:
            key = (code, periods[period_field])
            source_entry = mappings[mapping_name].get(key)
            if source_entry is None:
                continue
            source_value, _ = source_entry
            matched_source_keys[mapping_name].add(key)

            target_header = RATIO_HEADERS[source_field]
            target_cell = worksheet.cell(row_number, header_map[target_header])
            if target_cell.data_type == 'f':
                workbook.close()
                cached_workbook.close()
                raise ValueError(
                    f'拒绝覆盖公式单元格：{worksheet.title}!{target_cell.coordinate}'
                )
            if not is_missing_target(target_cell.value):
                protected_non_missing[target_header] += 1
                continue
            if values_equal(target_cell.value, source_value):
                continue

            target_cell.value = source_value
            snapshot_key = (worksheet.title, target_cell.coordinate)
            changed_values[snapshot_key] = source_value
            changes_by_field[target_header] += 1

    cached_workbook.close()

    temporary_output = output_file.with_name(
        f'.{output_file.stem}_{timestamp}.tmp.xlsx'
    )
    try:
        show_work_status('正在保存并校验输出文件')
        workbook.save(temporary_output)
        workbook.close()
        verify_saved_output(
            temporary_output,
            values_before,
            formulas_before,
            structure_before,
            changed_values,
        )

        previous_output_backup = None
        if output_file.exists():
            previous_output_backup = backup_dir / (
                f'{output_file.stem}_覆盖前备份_{timestamp}{output_file.suffix}'
            )
            shutil.copy2(output_file, previous_output_backup)
        os.replace(temporary_output, output_file)
    finally:
        if temporary_output.exists():
            temporary_output.unlink()
        try:
            workbook.close()
        except Exception:
            pass

    if file_sha256(ratio_file) != source_hash_before:
        raise RuntimeError(f'源比例文件被意外修改：{ratio_file}')
    show_work_status('缺失回填完成', finished=True)

    unused_source_keys = {
        field_name: len(set(field_mapping) - matched_source_keys[field_name])
        for field_name, field_mapping in mappings.items()
    }
    return {
        'source_backup': source_backup,
        'previous_output_backup': previous_output_backup,
        'source_row_count': source_row_count,
        'source_key_count': {name: len(values) for name, values in mappings.items()},
        'unused_source_keys': unused_source_keys,
        'changed_cell_count': len(changed_values),
        'changes_by_field': dict(changes_by_field),
        'protected_non_missing': dict(protected_non_missing),
        'formula_count': len(formulas_before),
    }

"""""""""
   主体
"""""""""
default_year = int(yr)
ratio_file, date1 = select_latest_dated_file(
    DATA_DIR, RATIO_FILE_PATTERN, '买方反应-比例文件', default_year
)
missing_file, date2 = select_latest_dated_file(
    DATA_DIR, MISSING_FILE_PATTERN, '买方反应-缺失信息文件', default_year
)
output_file = DATA_DIR / f'买方反应-比例填充{date1}.xlsx'

result = backfill_ratio_file(
    ratio_file=ratio_file,
    missing_file=missing_file,
    output_file=output_file,
    backup_dir=BACKUP_DIR,
)
"""""""""
   主体
"""""""""

## 导出...
print(f'比例文件（date1={date1}）：{ratio_file}')
print(f'缺失信息文件（date2={date2}）：{missing_file}')
print(f'源文件备份：{result["source_backup"]}')
if result['previous_output_backup'] is not None:
    print(f'原输出备份：{result["previous_output_backup"]}')
print(f'缺失信息记录：{result["source_row_count"]}')
print(f'回填单元格：{result["changed_cell_count"]}')
for field_name, count in result['changes_by_field'].items():
    print(f'  {field_name}：{count}')
for field_name, count in result['protected_non_missing'].items():
    print(f'保留已有非缺失值 {field_name}：{count}')
for field_name, count in result['unused_source_keys'].items():
    print(f'未匹配来源键 {field_name}：{count}')
print(f'公式校验：{result["formula_count"]} 个公式保持不变')
print(f'输出文件：{output_file}')
print('success')

# 添加时间记录
insert_bookmark(__file__, 1, rel, start)
if os.environ.get('GB_SKIP_FLUSH_BUFFER') != '1':
    open_file('flush_buffer.py')


