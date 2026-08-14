import sys, time, os, warnings, re
from pathlib import Path
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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
# eg.
# df = pd.read_excel(f'{impt}filename.xlsx', parse_dates=['成立日期','清盘日期'])

"""""""""
   主体
"""""""""

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
PROCESSED_YEAR_DIR = Path(processed) / str(yr)
SUBJECTIVE_DIR = PROCESSED_YEAR_DIR / '主观评分'
RAW_DIR = SUBJECTIVE_DIR / '原始数据'
STAT_DIR = SUBJECTIVE_DIR / '统计数据'
RESULT_ROOT = SCRIPT_DIR / '2_计算结果'

ANALYST_PATTERN = re.compile(
    r'^分析师投票_(?P<date>\d{8})\.xlsx$', re.IGNORECASE
)
ORG_PATTERN = re.compile(
    r'^机构投票_(?P<date>\d{8})\.xlsx$', re.IGNORECASE
)


def clean_text(value):
    """清理参与匹配的名称，避免首尾空格造成漏匹配。"""
    if pd.isna(value):
        return value
    return str(value).strip()


def read_export(file_path):
    """跳过中文说明行，使用第二行英文字段名读取导出文件。"""
    data = pd.read_excel(file_path, header=1)
    data.columns = [clean_text(column) for column in data.columns]
    for column in data.select_dtypes(include='object').columns:
        data[column] = data[column].map(clean_text)
    return data


def require_columns(data, columns, label):
    missing = [column for column in columns if column not in data.columns]
    if missing:
        raise ValueError(f'{label}缺少字段：{missing}')


def ensure_unique(data, columns, label):
    duplicated = data[data.duplicated(columns, keep=False)]
    if not duplicated.empty:
        samples = duplicated[columns].drop_duplicates().head(10).to_dict('records')
        raise ValueError(f'{label}的匹配键不唯一：{samples}')


def make_excel_style(file_path, sheet_names=None):
    """统一输出格式；分析师记分表保持 2025 年双表头结构。"""
    wb = load_workbook(file_path)
    thin = Side(style='thin', color='D9D9D9')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill('solid', fgColor='B4C6E7')
    field_fill = PatternFill('solid', fgColor='D9E2F3')

    for ws in wb.worksheets:
        if sheet_names is not None and ws.title not in sheet_names:
            continue
        ws.freeze_panes = 'A3' if ws.title == '分析师名单表' else 'A2'
        ws.auto_filter.ref = ws.dimensions
        ws.sheet_view.showGridLines = False
        for cell in ws[1]:
            cell.font = Font(name='宋体', size=11, bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        if ws.title == '分析师名单表':
            for cell in ws[2]:
                cell.font = Font(name='Calibri', size=11, bold=True)
                cell.fill = field_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            data_start = 3
        else:
            data_start = 2
        for row in ws.iter_rows(min_row=data_start):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical='center')
        for column_index, column_cells in enumerate(ws.iter_cols(), start=1):
            max_length = max(
                (len(str(cell.value)) for cell in column_cells if cell.value is not None),
                default=0,
            )
            ws.column_dimensions[get_column_letter(column_index)].width = min(
                max(max_length + 2, 10), 44
            )

    wb.save(file_path)


def build_analyst_scores(vote_file, roster_file, output_file):
    votes = read_export(vote_file)
    required_vote_columns = [
        'f_analysis_name', 'f_award_name', 'f_user_coefficient'
    ]
    require_columns(votes, required_vote_columns, '分析师投票')

    roster = pd.read_excel(roster_file)
    roster.columns = [clean_text(column) for column in roster.columns]
    require_columns(
        roster,
        ['序号', '分析师姓名', '研究领域', '公司名称'],
        '参赛名单',
    )
    roster = roster[['序号', '分析师姓名', '研究领域', '公司名称']].copy()
    for column in ['分析师姓名', '研究领域', '公司名称']:
        roster[column] = roster[column].map(clean_text)
    ensure_unique(roster, ['分析师姓名', '研究领域'], '参赛名单')

    votes['f_user_coefficient'] = pd.to_numeric(
        votes['f_user_coefficient'], errors='raise'
    )
    grouped = (
        votes.groupby(
            ['f_analysis_name', 'f_award_name'],
            as_index=False,
            dropna=False,
        )['f_user_coefficient']
        .sum()
        .rename(columns={
            'f_analysis_name': '分析师姓名',
            'f_award_name': '研究领域',
            'f_user_coefficient': '系数累加票数',
        })
    )
    ensure_unique(grouped, ['分析师姓名', '研究领域'], '分析师投票汇总')

    # 全部分析师中最高的系数累加票数对应 100 分。
    max_votes = grouped['系数累加票数'].max()
    grouped['主观评价分数'] = (
        grouped['系数累加票数'] / max_votes * 100
        if pd.notna(max_votes) and max_votes != 0
        else 0
    )

    analyst_scores = roster.merge(
        grouped[['分析师姓名', '研究领域', '主观评价分数']],
        on=['分析师姓名', '研究领域'],
        how='left',
        validate='one_to_one',
    )
    analyst_scores['主观评价分数'] = (
        analyst_scores['主观评价分数'].fillna(0)
    )
    analyst_scores = analyst_scores.rename(columns={'序号': '分析师ID'})[
        ['分析师ID', '分析师姓名', '研究领域', '公司名称', '主观评价分数']
    ]

    english_header = pd.DataFrame([{
        '分析师ID': 'id',
        '分析师姓名': 'analysis_name',
        '研究领域': 'research_domain',
        '公司名称': 'company_name',
        '主观评价分数': 'score',
    }])
    output_data = pd.concat([english_header, analyst_scores], ignore_index=True)
    Path(output_file).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        output_data.to_excel(writer, sheet_name='分析师名单表', index=False)
    make_excel_style(output_file, {'分析师名单表'})
    return analyst_scores, grouped


def safe_sheet_name(name, used_names):
    cleaned = re.sub(r'[\\/*?:\[\]]', '_', clean_text(name))[:31] or '未命名奖项'
    sheet_name = cleaned
    suffix = 2
    while sheet_name in used_names:
        suffix_text = f'_{suffix}'
        sheet_name = f'{cleaned[:31-len(suffix_text)]}{suffix_text}'
        suffix += 1
    used_names.add(sheet_name)
    return sheet_name


def build_organization_scores(vote_file, output_file):
    votes = read_export(vote_file)
    required_columns = [
        'f_award_id', 'f_award_name', 'f_entry_id', 'f_entry_name',
        'f_user_coefficient',
    ]
    require_columns(votes, required_columns, '机构投票')
    votes['f_user_coefficient'] = pd.to_numeric(
        votes['f_user_coefficient'], errors='raise'
    )

    grouped = (
        votes.groupby(
            ['f_award_id', 'f_award_name', 'f_entry_id', 'f_entry_name'],
            as_index=False,
            dropna=False,
        )['f_user_coefficient']
        .sum()
        .rename(columns={
            'f_award_id': '奖项ID',
            'f_award_name': '奖项名称',
            'f_entry_id': '机构ID',
            'f_entry_name': '机构名称',
            'f_user_coefficient': '得分',
        })
    )

    Path(output_file).parent.mkdir(parents=True, exist_ok=True)
    used_names = set()
    sheet_names = []
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for award_name, award_data in grouped.groupby('奖项名称', sort=True):
            award_data = award_data.copy().sort_values(
                ['得分', '机构名称', '机构ID'],
                ascending=[False, True, True],
                kind='stable',
            )
            award_data.insert(
                0,
                '排名',
                award_data['得分'].rank(method='dense', ascending=False).astype(int),
            )
            award_data = award_data[
                ['排名', '机构ID', '机构名称', '奖项ID', '奖项名称', '得分']
            ]
            sheet_name = safe_sheet_name(award_name, used_names)
            sheet_names.append(sheet_name)
            award_data.to_excel(writer, sheet_name=sheet_name, index=False)
    make_excel_style(output_file, set(sheet_names))
    return grouped, sheet_names


analyst_vote_file, analyst_date = select_latest_dated_file(
    RAW_DIR,
    ANALYST_PATTERN,
    '分析师投票文件',
    int(yr),
    output_format='%m%d',
)
org_vote_file, org_date = select_latest_dated_file(
    RAW_DIR,
    ORG_PATTERN,
    '机构投票文件',
    int(yr),
    output_format='%m%d',
)
if analyst_date != org_date:
    raise ValueError(
        f'分析师与机构投票日期不一致：{analyst_date} / {org_date}'
    )

date = analyst_date  # 文件名与结果文件夹统一使用四位 MMDD。
roster_file = PROCESSED_YEAR_DIR / '参赛名单0807.xlsx'
if not roster_file.exists():
    raise FileNotFoundError(f'参赛名单不存在：{roster_file}')

analyst_output = STAT_DIR / f'分析师记分统计{date}.xlsx'
result_dir = RESULT_ROOT / date
organization_output = result_dir / f'机构奖投票结果{date}.xlsx'

print(f'分析师投票：{analyst_vote_file.name}', flush=True)
print(f'机构投票：{org_vote_file.name}', flush=True)
print(f'四位日期：{date}', flush=True)

analyst_scores, analyst_vote_scores = build_analyst_scores(
    analyst_vote_file,
    roster_file,
    analyst_output,
)
organization_scores, organization_sheets = build_organization_scores(
    org_vote_file,
    organization_output,
)

print(
    f'分析师记分统计：{len(analyst_scores)} 人，'
    f'{len(analyst_vote_scores)} 人有票，输出 {analyst_output}',
    flush=True,
)
print(
    f'机构奖统计：{len(organization_scores)} 条，'
    f'{len(organization_sheets)} 个奖项，输出 {organization_output}',
    flush=True,
)
"""""""""
   主体
"""""""""

print('success')

# 添加时间记录
insert_bookmark(__file__, 1, rel, start)
open_file('flush_buffer.py')


