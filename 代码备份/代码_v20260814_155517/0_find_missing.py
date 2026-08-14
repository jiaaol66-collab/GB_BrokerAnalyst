import sys, os, time, warnings
import unicodedata
from collections import defaultdict
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
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

warnings.filterwarnings('ignore')

from userlib.configs import analyst_config as C

rel = find_rel(__file__, '参评信息')
start = insert_bookmark(__file__,0, rel)
impt = find_rel(__file__, '重要文件')
mtr = find_rel(__file__, '资料')
processed = find_rel(__file__, '已处理')
yr = get_year(rel)

CFG = C.for_year(yr)
dataset_cfg = CFG.dataset_ingestion
dataset_path = get_terminal_dataset_path(yr, __file__)
if not dataset_path:
    raise FileNotFoundError(
        f'未找到 {yr} 年 Terminal 原始数据集配置，请先在 Terminal 中指定数据集位置。'
    )

bundle_root = Path(dataset_path).joinpath(*dataset_cfg['bundle_parts'])
source_dirs = [bundle_root / item for item in dataset_cfg['source_dirs']]
user_src = bundle_root / dataset_cfg['info_file']

missing_paths = [str(path) for path in [*source_dirs, user_src] if not path.exists()]
if missing_paths:
    raise FileNotFoundError(
        '分析师原始数据集缺少 CFG 指定的文件或目录：\n' + '\n'.join(missing_paths)
    )

processed_year_root = Path(processed) / str(yr)
upload_data_root = processed_year_root / 'goldenbull_fxs'
supplement_data_root = processed_year_root / 'goldenbull_fxs_ex'
user_dst = f'{processed}{yr}/用户报名表.xlsx'
output_dir = Path(__file__).resolve().parent.parent / 'step_1_手动补齐报名表'
make_dir(str(output_dir))

try:
    user = pd.read_excel(
        user_src,
        skiprows=dataset_cfg['skiprows'],
        dtype={dataset_cfg['id_column']: str},
    )
except Exception as exc:
    raise RuntimeError(f'用户报名表无法读取：{user_src}') from exc

required_columns = [
    dataset_cfg['id_column'],
    dataset_cfg['company_short_name_column'],
    dataset_cfg['company_name_column'],
]
missing_columns = [column for column in required_columns if column not in user.columns]
if missing_columns:
    raise KeyError(
        f'用户报名表缺少 CFG 指定列：{missing_columns}；实际列：{list(user.columns)}'
    )

path2 = r'../step_2_分析师股票提取清洗信息'
# make_dir(path2+r'/1_初步提取')

print('---------------正在寻找缺失表头的报名表{}--------------'.format(datetime.now()))

no_head = []

dataset_tag = Path(dataset_path).name
institution_root = Path(processed) / str(yr) / '机构报名表' / dataset_tag
make_dir(str(institution_root))

id_column = dataset_cfg['id_column']
short_name_column = dataset_cfg['company_short_name_column']
company_name_column = dataset_cfg['company_name_column']

institution_info = user.loc[
    user[id_column].notna(),
    [id_column, short_name_column, company_name_column],
].copy()
for column in [id_column, short_name_column, company_name_column]:
    institution_info[column] = institution_info[column].astype(str).str.strip()
institution_info = institution_info.drop_duplicates(subset=[id_column], keep='first')
institution_lookup = institution_info.set_index(id_column).to_dict(orient='index')

# 遍历 CFG 指定的上传资料目录
for source_dir in source_dirs:
    for subdir, dirs, files in os.walk(source_dir):
        folder_name = os.path.basename(subdir)
        all_ = len(files)
        cur_ = 0
        for file in files:
            org_path = os.path.join(subdir,file)
            cur_ += 1
            org_folder = os.path.basename(get_path(2,org_path)['n'].strip('/'))
            org_folder = f'(来自{org_folder}).'
            # 检查文件是否为支持的 Excel 格式
            print_progress(cur_, all_, 50, f'正在搜索{file}')
            institution_folder = institution_root / folder_name
            if not institution_folder.exists():
                make_dir(str(institution_folder), 1)
            if '机构' in file:
                copy_file(
                    org_path,
                    str(institution_folder / org_folder.join(file.split("."))),
                    0,
                    1,
                )
            if Path(file).suffix.lower() in {'.xlsx', '.xls'}:
                file_path = os.path.join(subdir, file)
                # 尝试读取Excel文件
                try:
                    xls = pd.ExcelFile(file_path)  # 按扩展名自动选择 openpyxl/xlrd
                    sheet_names = xls.sheet_names  # 获取所有工作表名
                    # 遍历所有工作表
                    for sheet_name in sheet_names:
                        df = pd.read_excel(xls,sheet_name=sheet_name)
                        if '分析师个人信息表' not in df.columns:
                            if (df.iloc[0][0] == '分析师所属部门' or df.iloc[0][0] == '分析师所在机构名称'):
                                no_head.append({
                                    'institution_id': folder_name,
                                    'file_name': os.path.basename(file_path),
                                    'file_path': str(Path(file_path).resolve()),
                                    'sheet_name': sheet_name,
                                })
                except:
                    pass


def local_absolute_link(path):
    """生成保留中文的 Windows 绝对路径，禁止使用会产生 %xx 编码的 as_uri()。"""
    return Path(path).resolve().as_posix()


def markdown_text(value):
    return (
        str(value or '-')
        .replace('\\', '\\\\')
        .replace('|', '\\|')
        .replace('[', '\\[')
        .replace(']', '\\]')
        .replace('\r', ' ')
        .replace('\n', ' ')
    )


FILE_MANIFEST_COLUMNS = [
    '上传文件',
    '补充资料',
    '机构编码',
    '机构简称',
    '来源',
    '文件名',
    '上传时间',
    '核对状态',
]
FILE_MANIFEST_NOISE = {'.DS_Store', 'Thumbs.db'}
SEAL_FILE_KEYWORDS = ('用印', '盖章', '公章', '印章', '签章', '用章')
REGISTRY_SUMMARY_COLUMNS = [
    '上传文件',
    '补充资料',
    '机构编码',
    '机构简称',
    '机构名称',
    '填报人',
    '电话',
    '邮箱',
    '上传文件数量',
    '补充数据数量',
    '是否含用印文件',
    '是否有上传文件缺失',
    '是否有补充数据缺失',
]
REGISTRY_MISSING_COLUMNS = [
    '机构编码',
    '机构简称',
    '填报人',
    '电话',
    '邮箱',
    '缺失上传文件名称',
    '缺失补充数据缺失名称',
]


def manifest_clean_text(value):
    """统一清单索引中的空值、空格和 Unicode 表示。"""
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    return unicodedata.normalize('NFC', str(value).strip())


def manifest_key(value):
    return manifest_clean_text(value).casefold()


institution_lookup_by_key = {
    manifest_key(institution_id): company
    for institution_id, company in institution_lookup.items()
}


def manifest_institution_short_name(institution_hash):
    company = institution_lookup_by_key.get(
        manifest_key(institution_hash),
        {},
    )
    return manifest_clean_text(company.get(short_name_column, ''))


def manifest_file_name(value):
    text = manifest_clean_text(value).replace('\\', '/')
    return text.rsplit('/', 1)[-1] if text else ''


def manifest_excel_time(value):
    text = manifest_clean_text(value)
    if not text:
        return pd.NaT
    parsed = pd.to_datetime(value, errors='coerce')
    if pd.notna(parsed):
        if getattr(parsed, 'tzinfo', None) is not None:
            parsed = parsed.tz_localize(None)
        return parsed.to_pydatetime()
    return text


def manifest_relative_path(source_path, workbook_parent):
    """生成从文件清单所在文件夹出发的真实相对路径。"""
    return os.path.relpath(source_path, workbook_parent).replace('\\', '/')


def manifest_display_path(source_path):
    """生成清单中展示用的逻辑短路径。"""
    relative_path = os.path.relpath(
        source_path,
        processed_year_root,
    ).replace('\\', '/')
    return f'../{relative_path}'


def manifest_actual_files(processed_dir, workbook_parent):
    records = []
    for path in sorted(
            processed_dir.rglob('*'),
            key=lambda item: item.as_posix().casefold(),
    ):
        if (
                not path.is_file()
                or path.name in FILE_MANIFEST_NOISE
                or path.name.startswith('~$')
        ):
            continue

        relative_inside_dir = path.relative_to(processed_dir)
        parts = relative_inside_dir.parts
        institution_hash = manifest_clean_text(parts[0]) if len(parts) >= 2 else ''
        records.append({
            '_key': (
                manifest_key(institution_hash),
                manifest_key(path.name),
            ),
            '_机构哈希': institution_hash,
            '_文件名': manifest_clean_text(path.name),
            '_链接路径': path.resolve(),
            '_来源相对路径': manifest_relative_path(
                path.resolve(),
                workbook_parent,
            ),
            '_结构异常': len(parts) < 2,
        })
    return records


def manifest_index_groups(index_frame):
    groups = defaultdict(list)
    for institution_raw, file_raw, time_raw, index_order in index_frame.itertuples(
            index=False,
            name=None,
    ):
        institution_hash = manifest_clean_text(institution_raw)
        file_name = manifest_file_name(file_raw)
        parsed_time = pd.to_datetime(time_raw, errors='coerce')
        sort_time = parsed_time.value if pd.notna(parsed_time) else -1
        key = (manifest_key(institution_hash), manifest_key(file_name))
        groups[key].append({
            '_机构哈希': institution_hash,
            '_文件名': file_name,
            '_上传时间': manifest_excel_time(time_raw),
            '_索引顺序': index_order,
            '_时间排序': sort_time,
        })
    return groups


def manifest_output_record(
        institution_hash,
        file_name,
        upload_time,
        source_path,
        workbook_parent,
        status,
):
    resolved = Path(source_path).resolve()
    upload_folder_path = (
        upload_data_root / institution_hash
    ).resolve()
    supplement_folder_path = (
        supplement_data_root / institution_hash
    ).resolve()
    return {
        '上传文件': (
            '打开文件夹' if upload_folder_path.is_dir() else '无文件夹'
        ),
        '补充资料': (
            '打开文件夹'
            if supplement_folder_path.is_dir()
            else '无文件夹'
        ),
        '机构编码': institution_hash,
        '机构简称': manifest_institution_short_name(institution_hash),
        '来源': manifest_display_path(resolved),
        '文件名': file_name,
        '上传时间': upload_time,
        '核对状态': status,
        '_链接路径': resolved,
        '_上传文件夹路径': upload_folder_path,
        '_补充文件夹路径': supplement_folder_path,
    }


def build_file_manifest_sheet(table_cfg, processed_dir, workbook_parent):
    index_path = bundle_root / table_cfg['index_file']
    if not index_path.is_file():
        raise FileNotFoundError(f'文件清单索引表不存在：{index_path}')
    if not processed_dir.is_dir():
        raise FileNotFoundError(f'文件清单资料目录不存在：{processed_dir}')

    index_frame = pd.read_excel(
        index_path,
        sheet_name=table_cfg['sheet_name'],
        dtype=object,
        skiprows=table_cfg.get('skiprows', []),
    )
    required_columns = [
        table_cfg['id_column'],
        table_cfg['file_column'],
        table_cfg['time_column'],
    ]
    missing = [
        column
        for column in required_columns
        if column not in index_frame.columns
    ]
    if missing:
        raise KeyError(
            f'文件清单索引表 {index_path.name} 缺少列：{missing}；'
            f'实际列：{list(index_frame.columns)}'
        )

    index_frame = index_frame[required_columns].copy()
    index_frame.columns = ['_机构哈希', '_文件名', '_上传时间']
    index_frame['_索引顺序'] = range(len(index_frame))
    grouped_index = manifest_index_groups(index_frame)

    actual_files = manifest_actual_files(processed_dir, workbook_parent)
    actual_by_key = defaultdict(list)
    for actual in actual_files:
        actual_by_key[actual['_key']].append(actual)

    rows = []
    consumed_paths = set()
    for key, records in grouped_index.items():
        selected = max(
            records,
            key=lambda item: (item['_时间排序'], item['_索引顺序']),
        )
        duplicate_note = f'索引重复{len(records)}条；' if len(records) > 1 else ''
        matches = actual_by_key.get(key, [])

        if matches:
            for match in matches:
                status = duplicate_note + '索引与文件匹配'
                if match['_结构异常']:
                    status += '；文件未位于哈希目录'
                rows.append(manifest_output_record(
                    match['_机构哈希'] or selected['_机构哈希'],
                    match['_文件名'],
                    selected['_上传时间'],
                    match['_链接路径'],
                    workbook_parent,
                    status,
                ))
                consumed_paths.add(
                    os.path.normcase(str(match['_链接路径']))
                )
            continue

        expected_path = (
            processed_dir
            / selected['_机构哈希']
            / selected['_文件名']
        )
        status = duplicate_note + '索引有记录但文件缺失'
        if not selected['_机构哈希']:
            status += '；机构哈希为空'
        if not selected['_文件名']:
            status += '；文件名为空'
        rows.append(manifest_output_record(
            selected['_机构哈希'],
            selected['_文件名'],
            selected['_上传时间'],
            expected_path,
            workbook_parent,
            status,
        ))

    for actual in actual_files:
        path_key = os.path.normcase(str(actual['_链接路径']))
        if path_key in consumed_paths:
            continue
        status = '文件存在但索引无记录'
        if actual['_结构异常']:
            status += '；文件未位于哈希目录'
        rows.append(manifest_output_record(
            actual['_机构哈希'],
            actual['_文件名'],
            pd.NaT,
            actual['_链接路径'],
            workbook_parent,
            status,
        ))

    frame = pd.DataFrame(
        rows,
        columns=FILE_MANIFEST_COLUMNS + [
            '_链接路径',
            '_上传文件夹路径',
            '_补充文件夹路径',
        ],
    )
    if not frame.empty:
        frame = frame.sort_values(
            ['机构编码', '来源'],
            kind='stable',
            key=lambda series: series.map(manifest_key),
        ).reset_index(drop=True)
    return frame


def write_file_manifest():
    manifest_tables = dataset_cfg.get('manifest_tables', {})
    if not manifest_tables:
        print(f'{yr} 年未配置文件清单索引，跳过文件清单构建。')
        return None

    output_subdir = dataset_cfg.get(
        'manifest_output_subdir',
        '',
    )
    workbook_parent = (processed_year_root / output_subdir).resolve()
    make_dir(str(workbook_parent), 1)
    workbook_path = workbook_parent / '文件清单.xlsx'

    frames = {}
    for sheet_name in ['上传文件', '补充数据']:
        if sheet_name not in manifest_tables:
            raise KeyError(f'分析师配置缺少文件清单 Sheet：{sheet_name}')
        table_cfg = manifest_tables[sheet_name]
        processed_dir = (
            processed_year_root
            / table_cfg['processed_dir']
        ).resolve()
        frames[sheet_name] = build_file_manifest_sheet(
            table_cfg,
            processed_dir,
            workbook_parent,
        )

    check_with_retry(str(workbook_path))
    with pd.ExcelWriter(workbook_path, engine='openpyxl') as writer:
        for sheet_name, frame in frames.items():
            frame[FILE_MANIFEST_COLUMNS].to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

    workbook = load_workbook(workbook_path)
    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    warning_fill = PatternFill('solid', fgColor='FFF2CC')
    error_fill = PatternFill('solid', fgColor='FCE4D6')
    duplicate_fill = PatternFill('solid', fgColor='DDEBF7')
    widths = [35, 35, 20, 22, 65, 55, 23, 32]

    for sheet_name, frame in frames.items():
        worksheet = workbook[sheet_name]
        worksheet.freeze_panes = 'A2'
        worksheet.sheet_view.showGridLines = False
        worksheet.auto_filter.ref = f'A1:H{max(worksheet.max_row, 1)}'
        worksheet.row_dimensions[1].height = 23

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal='center',
                vertical='center',
            )
        for column_index, width in enumerate(widths, start=1):
            worksheet.column_dimensions[
                get_column_letter(column_index)
            ].width = width

        for row_index, link_path in enumerate(
                frame['_链接路径'],
                start=2,
        ):
            upload_folder_path = frame.iloc[
                row_index - 2
            ]['_上传文件夹路径']
            upload_folder_cell = worksheet.cell(
                row=row_index,
                column=1,
            )
            if Path(upload_folder_path).is_dir():
                upload_folder_cell.hyperlink = manifest_relative_path(
                    upload_folder_path,
                    workbook_parent,
                )
                upload_folder_cell.style = 'Hyperlink'

            supplement_folder_path = frame.iloc[
                row_index - 2
            ]['_补充文件夹路径']
            supplement_folder_cell = worksheet.cell(
                row=row_index,
                column=2,
            )
            if Path(supplement_folder_path).is_dir():
                supplement_folder_cell.hyperlink = manifest_relative_path(
                    supplement_folder_path,
                    workbook_parent,
                )
                supplement_folder_cell.style = 'Hyperlink'

            relative_cell = worksheet.cell(row=row_index, column=5)
            relative_cell.hyperlink = manifest_relative_path(
                link_path,
                workbook_parent,
            )
            relative_cell.style = 'Hyperlink'
            worksheet.cell(row=row_index, column=7).number_format = (
                'yyyy-mm-dd hh:mm:ss'
            )

            status_cell = worksheet.cell(row=row_index, column=8)
            status = manifest_clean_text(status_cell.value)
            if '文件缺失' in status:
                status_cell.fill = error_fill
            elif '索引无记录' in status:
                status_cell.fill = warning_fill
            elif '索引重复' in status:
                status_cell.fill = duplicate_fill

            for column_index in range(1, len(FILE_MANIFEST_COLUMNS) + 1):
                worksheet.cell(
                    row=row_index,
                    column=column_index,
                ).alignment = Alignment(vertical='top')

    workbook.save(workbook_path)
    print(f'文件清单：{workbook_path}')
    for sheet_name, frame in frames.items():
        summary = frame['核对状态'].value_counts(dropna=False).to_dict()
        print(f'{sheet_name}：{len(frame)}行，{summary}')
    return workbook_path, frames


def manifest_frame_statistics(frame):
    """仅从已生成的文件清单 DataFrame 汇总机构文件状态。"""
    file_counts = defaultdict(int)
    missing_files = defaultdict(list)
    seal_institution_ids = set()
    for record in frame.to_dict(orient='records'):
        institution_id_key = manifest_key(record['机构编码'])
        if not institution_id_key:
            continue

        status = manifest_clean_text(record['核对状态'])
        file_name = manifest_clean_text(record['文件名'])
        if '文件缺失' not in status:
            file_counts[institution_id_key] += 1
            if any(
                    keyword in file_name
                    for keyword in SEAL_FILE_KEYWORDS
            ):
                seal_institution_ids.add(institution_id_key)
            continue
        if not file_name:
            continue
        known_names = {
            manifest_key(item)
            for item in missing_files[institution_id_key]
        }
        if manifest_key(file_name) not in known_names:
            missing_files[institution_id_key].append(file_name)

    return {
        'file_counts': file_counts,
        'missing_files': missing_files,
        'seal_institution_ids': seal_institution_ids,
    }


def write_user_registry_audit(manifest_frames):
    """用文件清单 DataFrame 在原始用户注册表旁生成机构汇总。"""
    required_sheets = ['上传文件', '补充数据']
    missing_sheets = [
        sheet_name
        for sheet_name in required_sheets
        if sheet_name not in manifest_frames
    ]
    if missing_sheets:
        raise KeyError(f'文件清单 DataFrame 缺少 Sheet：{missing_sheets}')

    registry_columns = [
        id_column,
        short_name_column,
        company_name_column,
        'f_register',
        'f_phone',
        'f_email',
    ]
    missing_registry_columns = [
        column for column in registry_columns
        if column not in user.columns
    ]
    if missing_registry_columns:
        raise KeyError(
            f'用户注册表缺少机构核对列：{missing_registry_columns}；'
            f'实际列：{list(user.columns)}'
        )

    upload_stats = manifest_frame_statistics(
        manifest_frames['上传文件'],
    )
    supplement_stats = manifest_frame_statistics(
        manifest_frames['补充数据'],
    )

    summary_rows = []
    missing_rows = []
    for row in user[registry_columns].to_dict(orient='records'):
        institution_id = manifest_clean_text(row[id_column])
        institution_id_key = manifest_key(institution_id)
        company_short_name = manifest_clean_text(row[short_name_column])
        company_name = manifest_clean_text(row[company_name_column])
        registrant = manifest_clean_text(row['f_register'])
        phone = manifest_clean_text(row['f_phone'])
        email = manifest_clean_text(row['f_email'])

        upload_folder = (source_dirs[0] / institution_id).resolve()
        supplement_folder = (
            source_dirs[1] / institution_id
        ).resolve()

        missing_upload_files = upload_stats['missing_files'].get(
            institution_id_key,
            [],
        )
        missing_supplement_files = supplement_stats['missing_files'].get(
            institution_id_key,
            [],
        )
        contains_seal_file = (
            institution_id_key in upload_stats['seal_institution_ids']
            or institution_id_key
            in supplement_stats['seal_institution_ids']
        )

        summary_rows.append({
            '上传文件': (
                '打开文件夹' if upload_folder.is_dir() else '无文件夹'
            ),
            '补充资料': (
                '打开文件夹'
                if supplement_folder.is_dir()
                else '无文件夹'
            ),
            '机构编码': institution_id,
            '机构简称': company_short_name,
            '机构名称': company_name,
            '填报人': registrant,
            '电话': phone,
            '邮箱': email,
            '上传文件数量': upload_stats['file_counts'].get(
                institution_id_key,
                0,
            ),
            '补充数据数量': supplement_stats['file_counts'].get(
                institution_id_key,
                0,
            ),
            '是否含用印文件': '是' if contains_seal_file else '否',
            '是否有上传文件缺失': (
                '是' if missing_upload_files else '否'
            ),
            '是否有补充数据缺失': (
                '是' if missing_supplement_files else '否'
            ),
            '_上传文件夹路径': upload_folder,
            '_补充文件夹路径': supplement_folder,
        })

        if missing_upload_files or missing_supplement_files:
            missing_rows.append({
                '机构编码': institution_id,
                '机构简称': company_short_name,
                '填报人': registrant,
                '电话': phone,
                '邮箱': email,
                '缺失上传文件名称': '\n'.join(missing_upload_files),
                '缺失补充数据缺失名称': '\n'.join(
                    missing_supplement_files
                ),
            })

    summary_frame = pd.DataFrame(
        summary_rows,
        columns=REGISTRY_SUMMARY_COLUMNS + [
            '_上传文件夹路径',
            '_补充文件夹路径',
        ],
    )
    missing_frame = pd.DataFrame(
        missing_rows,
        columns=REGISTRY_MISSING_COLUMNS,
    )

    workbook_path = user_src.with_name(
        f'{user_src.stem}_文件汇总.xlsx'
    ).resolve()
    workbook_parent = workbook_path.parent
    check_with_retry(str(workbook_path))
    with pd.ExcelWriter(workbook_path, engine='openpyxl') as writer:
        summary_frame[REGISTRY_SUMMARY_COLUMNS].to_excel(
            writer,
            sheet_name='机构汇总',
            index=False,
        )
        missing_frame[REGISTRY_MISSING_COLUMNS].to_excel(
            writer,
            sheet_name='缺失表',
            index=False,
        )

    workbook = load_workbook(workbook_path)
    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    warning_fill = PatternFill('solid', fgColor='FFF2CC')
    error_fill = PatternFill('solid', fgColor='FCE4D6')
    sheet_settings = {
        '机构汇总': (
            REGISTRY_SUMMARY_COLUMNS,
            [35, 35, 20, 22, 38, 18, 18, 32, 16, 16, 18, 22, 22],
        ),
        '缺失表': (
            REGISTRY_MISSING_COLUMNS,
            [20, 22, 18, 18, 32, 55, 55],
        ),
    }

    for sheet_name, (columns, widths) in sheet_settings.items():
        worksheet = workbook[sheet_name]
        worksheet.freeze_panes = 'A2'
        worksheet.sheet_view.showGridLines = False
        worksheet.auto_filter.ref = (
            f'A1:{get_column_letter(len(columns))}'
            f'{max(worksheet.max_row, 1)}'
        )
        worksheet.row_dimensions[1].height = 28
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(
                horizontal='center',
                vertical='center',
                wrap_text=True,
            )
        for column_index, width in enumerate(widths, start=1):
            worksheet.column_dimensions[
                get_column_letter(column_index)
            ].width = width

        for row_index in range(2, worksheet.max_row + 1):
            for column_index in range(1, len(columns) + 1):
                worksheet.cell(
                    row=row_index,
                    column=column_index,
                ).alignment = Alignment(
                    vertical='top',
                    wrap_text=(
                        sheet_name == '缺失表'
                        and column_index in (6, 7)
                    ),
                )

            if sheet_name == '机构汇总':
                if worksheet.cell(row=row_index, column=11).value == '是':
                    worksheet.cell(
                        row=row_index,
                        column=11,
                    ).fill = warning_fill
                for column_index in (12, 13):
                    if (
                            worksheet.cell(
                                row=row_index,
                                column=column_index,
                            ).value == '是'
                    ):
                        worksheet.cell(
                            row=row_index,
                            column=column_index,
                        ).fill = error_fill
            else:
                line_count = max(
                    len(
                        str(
                            worksheet.cell(
                                row=row_index,
                                column=column_index,
                            ).value
                            or ''
                        ).splitlines()
                    )
                    for column_index in (6, 7)
                )
                worksheet.row_dimensions[row_index].height = min(
                    15 * max(line_count, 1),
                    90,
                )

    summary_sheet = workbook['机构汇总']
    for row_index, upload_folder_path in enumerate(
            summary_frame['_上传文件夹路径'],
            start=2,
    ):
        if Path(upload_folder_path).is_dir():
            upload_folder_cell = summary_sheet.cell(
                row=row_index,
                column=1,
            )
            upload_folder_cell.hyperlink = manifest_relative_path(
                upload_folder_path,
                workbook_parent,
            )
            upload_folder_cell.style = 'Hyperlink'

        supplement_folder_path = summary_frame.iloc[
            row_index - 2
        ]['_补充文件夹路径']
        if Path(supplement_folder_path).is_dir():
            supplement_folder_cell = summary_sheet.cell(
                row=row_index,
                column=2,
            )
            supplement_folder_cell.hyperlink = manifest_relative_path(
                supplement_folder_path,
                workbook_parent,
            )
            supplement_folder_cell.style = 'Hyperlink'

    workbook.save(workbook_path)
    print(
        f'用户注册表文件汇总：{workbook_path}；'
        f'机构汇总{len(summary_frame)}行，缺失表{len(missing_frame)}行'
    )
    return workbook_path


missing_report_path = output_dir / '1_缺失表头报名表位置.md'
delete_file(str(output_dir / '1_缺失表头报名表位置.txt'))
delete_file(str(missing_report_path))
with missing_report_path.open('w', encoding='utf-8', newline='\n') as f:
    f.write('# 缺失表头报名表位置\n\n')
    f.write(f'- 数据集：`{markdown_text(dataset_path)}`\n')
    f.write(f'- 错误数量：{len(no_head)}\n\n')
    f.write('| 机构简称 | 机构名称 | 机构哈希编码 | 工作表 | 错误文件（真实绝对路径） |\n')
    f.write('| --- | --- | --- | --- | --- |\n')
    for item in no_head:
        institution_id = item['institution_id']
        company = institution_lookup.get(institution_id, {})
        company_short_name = company.get(short_name_column, '-')
        company_name = company.get(company_name_column, '-')
        file_link_target = local_absolute_link(item['file_path'])
        link_label = markdown_text(f"打开：{item['file_name']}")
        f.write(
            f'| {markdown_text(company_short_name)}'
            f' | {markdown_text(company_name)}'
            f' | `{markdown_text(institution_id)}`'
            f' | {markdown_text(item["sheet_name"])}'
            f' | [{link_label}](<{file_link_target}>) |\n'
        )
        print(
            company_short_name,
            company_name,
            institution_id,
            item['file_name'],
            item['sheet_name'],
        )

institution_index = institution_info.rename(columns={
    short_name_column: '机构简称',
    company_name_column: '机构名称',
    id_column: '链接位置占位',
})[['机构简称', '机构名称', '链接位置占位']]

copy_file(str(source_dirs[0]), str(upload_data_root))
copy_file(str(source_dirs[1]), str(supplement_data_root))

supplement_hashes = {
    path.name
    for path in source_dirs[1].iterdir()
    if path.is_dir()
}

institution_folder_paths = []
upload_folder_paths = []
supplement_folder_paths = []
for institution_id in institution_index['链接位置占位']:
    institution_folder = (institution_root / institution_id).resolve()
    upload_folder = (upload_data_root / institution_id).resolve()
    supplement_folder = (
        (supplement_data_root / institution_id).resolve()
        if institution_id in supplement_hashes
        else None
    )
    make_dir(str(institution_folder))
    make_dir(str(upload_folder))
    institution_folder_paths.append(institution_folder)
    upload_folder_paths.append(upload_folder)
    supplement_folder_paths.append(supplement_folder)

institution_index['上传数据'] = institution_index['链接位置占位']
institution_index['补充资料'] = [
    institution_id if supplement_folder is not None else ''
    for institution_id, supplement_folder in zip(
        institution_index['链接位置占位'],
        supplement_folder_paths,
    )
]

institution_index_path = Path(processed) / str(yr) / '机构报名表' / '机构报名表.xlsx'
check_with_retry(str(institution_index_path))
with pd.ExcelWriter(institution_index_path, engine='openpyxl') as writer:
    institution_index.to_excel(writer, sheet_name='机构报名表索引', index=False)
    worksheet = writer.sheets['机构报名表索引']
    worksheet.freeze_panes = 'A2'
    worksheet.auto_filter.ref = worksheet.dimensions
    worksheet.column_dimensions['A'].width = 22
    worksheet.column_dimensions['B'].width = 38
    worksheet.column_dimensions['C'].width = 24
    worksheet.column_dimensions['D'].width = 24
    worksheet.column_dimensions['E'].width = 24
    for row_number, institution_folder in enumerate(institution_folder_paths, start=2):
        link_cell = worksheet.cell(row=row_number, column=3)
        link_cell.hyperlink = local_absolute_link(institution_folder)
        link_cell.style = 'Hyperlink'
    for row_number, upload_folder in enumerate(upload_folder_paths, start=2):
        link_cell = worksheet.cell(row=row_number, column=4)
        link_cell.hyperlink = local_absolute_link(upload_folder)
        link_cell.style = 'Hyperlink'
    for row_number, supplement_folder in enumerate(supplement_folder_paths, start=2):
        if supplement_folder is None:
            continue
        link_cell = worksheet.cell(row=row_number, column=5)
        link_cell.hyperlink = local_absolute_link(supplement_folder)
        link_cell.style = 'Hyperlink'

file_manifest_path, file_manifest_frames = write_file_manifest()
registry_audit_path = write_user_registry_audit(file_manifest_frames)

print('success')

copy_file(str(user_src), user_dst)

# open_file('../step_2_分析师股票提取清洗信息/1_初步提取')
# open_file('../step_1_手动补齐报名表/0_手动补齐说明.docx')
open_file(f'{processed}/{yr}/goldenbull_fxs')

insert_bookmark(__file__, 1, rel, start)
