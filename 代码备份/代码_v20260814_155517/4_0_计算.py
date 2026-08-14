import sys, os, time
import json
import re
import subprocess
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
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

rel = find_rel(__file__, '参评信息')
start = insert_bookmark(__file__,0, rel)
impt = find_rel(__file__, '重要文件')
mtr = find_rel(__file__, '资料')
processed = find_rel(__file__, '已处理')
org = find_rel(__file__, '原始数据集')
yr = get_year(rel)

WORKSPACE_DIR = Path(__file__).resolve().parents[1]
IMPORTANT_DIR = WORKSPACE_DIR / '重要文件'


def require_columns(frame, required_columns, label):
    missing_columns = sorted(set(required_columns) - set(frame.columns))
    if missing_columns:
        raise KeyError(f'{label}缺少列：{missing_columns}')


def require_sequence_column(frame, label, unique=True):
    if frame.columns.empty or frame.columns[0] != '序号':
        first_column = None if frame.columns.empty else frame.columns[0]
        raise ValueError(f'{label}第一列必须是“序号”，实际为：{first_column}')
    if frame['序号'].isna().any():
        raise ValueError(f'{label}存在空序号')
    if unique and not frame['序号'].is_unique:
        raise ValueError(f'{label}存在重复序号，不能进行一对一合并')


def read_formula_columns_with_excel(workbook_path, column_names, expected_rows):
    """通过Excel只读计算并批量读取公式列，不保存或修改源工作簿。"""
    powershell_script = r'''
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$WorkbookPath = $env:GB_ANALYST_WORKBOOK_PATH
$ColumnNames = ConvertFrom-Json $env:GB_ANALYST_COLUMN_NAMES
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    # 仅在Excel内存中计算，关闭时明确不保存；部分Office版本传ReadOnly可选参数会打开失败。
    $workbook = $excel.Workbooks.Open($WorkbookPath)
    try {
        $worksheet = $workbook.Worksheets.Item(1)
        $lastRow = $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row
        $lastColumn = $worksheet.Cells(1, $worksheet.Columns.Count).End(-4159).Column
        $worksheet.Calculate()
        $headerColumns = @{}
        for ($column = 1; $column -le $lastColumn; $column++) {
            $headerText = [string]$worksheet.Cells(1, $column).Value2
            if (-not [string]::IsNullOrEmpty($headerText)) {
                if ($headerColumns.ContainsKey($headerText)) {
                    throw "表头重复，无法唯一读取：$headerText"
                }
                $headerColumns[$headerText] = $column
            }
        }
        $result = [ordered]@{}
        foreach ($ColumnName in $ColumnNames) {
            if (-not $headerColumns.ContainsKey([string]$ColumnName)) {
                throw "未找到列：$ColumnName"
            }
            $column = $headerColumns[[string]$ColumnName]
            $range = $worksheet.Range(
                $worksheet.Cells(2, $column),
                $worksheet.Cells($lastRow, $column)
            )
            $raw = $range.Value2
            $values = [System.Collections.Generic.List[object]]::new()
            $rowLower = $raw.GetLowerBound(0)
            $rowUpper = $raw.GetUpperBound(0)
            $columnLower = $raw.GetLowerBound(1)
            for ($row = $rowLower; $row -le $rowUpper; $row++) {
                $value = $raw.GetValue($row, $columnLower)
                # Excel错误值经COM会变成约-21468262xx的Int32，不能当作业务数值。
                if (
                    $value -is [System.Int32] -and
                    $value -ge -2146826300 -and
                    $value -le -2146826200
                ) {
                    $value = $null
                }
                $values.Add($value)
            }
            $result[$ColumnName] = @($values)
        }
        [PSCustomObject]$result | ConvertTo-Json -Compress -Depth 4
    }
    finally {
        $workbook.Close($false)
    }
}
finally {
    $excel.Quit()
}
'''
    command_environment = os.environ.copy()
    command_environment['GB_ANALYST_WORKBOOK_PATH'] = str(
        Path(workbook_path).resolve()
    )
    command_environment['GB_ANALYST_COLUMN_NAMES'] = json.dumps(
        list(column_names), ensure_ascii=False
    )
    completed = subprocess.run(
        [
            'powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass',
            '-Command', powershell_script,
        ],
        capture_output=True,
        text=True,
        encoding='utf-8-sig',
        errors='replace',
        timeout=180,
        env=command_environment,
    )
    if completed.returncode != 0:
        error_message = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(
            f'Excel读取公式列失败：{error_message}'
        )
    try:
        values_by_column = json.loads(completed.stdout.strip())
    except (json.JSONDecodeError, TypeError) as exc:
        raise RuntimeError(
            'Excel返回的公式列数据格式无法识别'
        ) from exc
    for column_name in column_names:
        values = values_by_column.get(column_name)
        if not isinstance(values, list) or len(values) != expected_rows:
            actual_rows = None if not isinstance(values, list) else len(values)
            raise RuntimeError(
                f'Excel返回“{column_name}”{actual_rows}行，预期{expected_rows}行'
            )
    return pd.DataFrame(values_by_column, columns=column_names)


default_year = int(yr)
industry_file, industry_date = select_latest_dated_file(
    IMPORTANT_DIR / '同行业回报率',
    re.compile(r'^同行业回报率(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'),
    '同行业回报率文件',
    default_year,
)
ratio_file, ratio_date = select_latest_dated_file(
    IMPORTANT_DIR / '分析师清洗版',
    re.compile(r'^买方反应-比例填充(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'),
    '买方反应-比例填充文件',
    default_year,
)
public_fund_file, public_fund_date = select_latest_dated_file(
    IMPORTANT_DIR / '分析师清洗版',
    re.compile(r'^买方反应-公募基金(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'),
    '买方反应-公募基金文件',
    default_year,
)
excluded_file, excluded_date = select_latest_dated_file(
    IMPORTANT_DIR / '名单',
    re.compile(r'^取消资格(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'),
    '取消资格文件',
    default_year,
)

subjective_directory = (
    Path(processed) / str(yr) / '主观评分' / '统计数据'
)
try:
    subjective_file, subjective_date = select_latest_dated_file(
        subjective_directory,
        re.compile(
            r'^分析师记分统计(?P<date>\d{4}|\d{6}|\d{8})\.xlsx$'
        ),
        '分析师主观评分统计文件',
        default_year,
        output_format='%m%d',
    )
except FileNotFoundError:
    subjective_file = None
    subjective_date = format_file_date_label(
        industry_date, default_year, output_format='%m%d'
    )

now = datetime.now()
nowdate = now.strftime('%Y%m%d')[2:]
print('---------------读取源文件---------------')
if subjective_file is not None:
    print(f'主观评分【{subjective_file.name}】')
print(f'同行业回报率【{industry_file.name}】')
print(f'买方反应-比例【{ratio_file.name}】')
print(f'买方反应-公募基金【{public_fund_file.name}】')
print(f'取消资格【{excluded_file.name}】')
time.sleep(1)

date = subjective_date

extractdate = '_' + industry_date
appe = extractdate

switch_method = ''
# switch_method = input('对于同一分析师的多个股票使用算数平均分排名还是最高分排名？(enter/else):')
if not switch_method:
    method = '算术平均'
else:
    method = '最大值'

# isobject = input('是否先进行客观评价排名?(1/else):')
isobject = '主观'
if not isobject: object = '_客观评价'
else:  
    if subjective_file is not None:
        object = ''

    else:
        print('没有主观评分相关的资料, 目前进行客观评价排名。')
        object = '_客观评价'
        isobject = ''

process_output_dir = Path('1_计算过程')
result_output_dir = Path('2_计算结果') / date
process_output_dir.mkdir(parents=True, exist_ok=True)
result_output_dir.mkdir(parents=True, exist_ok=True)

start = time.time()
df = pd.read_excel(industry_file)
ratio_data = pd.read_excel(ratio_file)
public_fund_data = pd.read_excel(public_fund_file)
excluded_data = pd.read_excel(excluded_file)

require_sequence_column(df, '同行业回报率文件')
require_sequence_column(ratio_data, '买方反应-比例填充文件')
require_sequence_column(public_fund_data, '买方反应-公募基金文件')
require_sequence_column(excluded_data, '取消资格文件', unique=False)
require_columns(
    df,
    {
        '序号', '标的股票名称', '推荐日期清洗版',
        '三个月回报率', '同行业三个月回报', '中期观察日',
        '三个月回报率_年化', '同行业三个月年化回报', '超额回报率三个月',
        '六个月回报率', '同行业六个月回报', '长期观察日',
        '六个月回报率_年化', '同行业六个月年化回报', '超额回报率六个月',
        '考察期回报率', '同行业考察期回报率', '距考察期截止天数',
        '考察期回报率_年化', '同行业考察期年化回报率', '超额回报率考察期',
        '推荐/目标价格',
        '6个月内最高价', '推荐时收盘价', '研报发布时标的价格', '扣分',
    },
    '同行业回报率文件',
)
require_columns(
    ratio_data,
    {
        '序号', '机构比例变化原始', '股东比例变化原始',
        '机构比例变化', '股东比例变化',
    },
    '买方反应-比例填充文件',
)
require_columns(
    public_fund_data,
    {'序号', '头部公募基金变化'},
    '买方反应-公募基金文件',
)

industry_formula_columns = [
    '标的股票名称', '推荐/目标价格', '扣分',
    '中期观察日', '三个月回报率_年化',
    '同行业三个月年化回报', '超额回报率三个月',
    '长期观察日', '六个月回报率_年化',
    '同行业六个月年化回报', '超额回报率六个月',
    '距考察期截止天数', '考察期回报率_年化',
    '同行业考察期年化回报率', '超额回报率考察期',
]
print('正在通过Excel只读获取公式结果...')
industry_formula_values = read_formula_columns_with_excel(
    industry_file, industry_formula_columns, len(df)
)
industry_formula_values.index = df.index
df[industry_formula_columns] = industry_formula_values[industry_formula_columns]

ratio_formula_columns = [
    '机构比例变化原始', '股东比例变化原始',
    '机构比例变化', '股东比例变化',
]
ratio_formula_values = read_formula_columns_with_excel(
    ratio_file, ratio_formula_columns, len(ratio_data)
)
ratio_formula_values.index = ratio_data.index
ratio_data[ratio_formula_columns] = ratio_formula_values[ratio_formula_columns]

public_formula_columns = ['头部公募基金变化']
public_formula_values = read_formula_columns_with_excel(
    public_fund_file, public_formula_columns, len(public_fund_data)
)
public_formula_values.index = public_fund_data.index
public_fund_data[public_formula_columns] = public_formula_values[
    public_formula_columns
]

ratio_data = ratio_data[[
    '序号', '机构比例变化原始', '股东比例变化原始',
    '机构比例变化', '股东比例变化',
]]
public_fund_data = public_fund_data[['序号', '头部公募基金变化']]
merge_columns = (
    set(ratio_data.columns[1:]) | set(public_fund_data.columns[1:])
)
overlapping_columns = sorted(merge_columns & set(df.columns))
if overlapping_columns:
    raise ValueError(f'同行业回报率文件已有待合并列，无法判断数据来源：{overlapping_columns}')

industry_row_count = len(df)
df = df.merge(ratio_data, on='序号', how='left', validate='one_to_one')
df = df.merge(public_fund_data, on='序号', how='left', validate='one_to_one')
if len(df) != industry_row_count:
    raise RuntimeError('按序号合并后行数发生变化，已停止计算')

excluded_ids = set(excluded_data['序号'].dropna())
excluded_row_count = int(df['序号'].isin(excluded_ids).sum())
df = df.loc[~df['序号'].isin(excluded_ids)].copy()
print(f'取消资格过滤：剔除 {excluded_row_count} 行，保留 {len(df)} 行')

recommendation_dates = pd.to_datetime(df['推荐日期清洗版'], errors='coerce')
if recommendation_dates.isna().any():
    invalid_ids = df.loc[recommendation_dates.isna(), '序号'].head(10).tolist()
    raise ValueError(f'取消资格过滤后仍有无法识别的推荐日期，序号示例：{invalid_ids}')
df['推荐日期清洗版'] = recommendation_dates
df['距考察期截止天数'] = pd.to_numeric(
    df['距考察期截止天数'], errors='coerce'
)
invalid_examination_period = (
    df['距考察期截止天数'].isna()
    | df['距考察期截止天数'].le(0)
)
if invalid_examination_period.any():
    invalid_ids = df.loc[invalid_examination_period, '序号'].head(10).tolist()
    raise ValueError(
        '同行业回报率文件的“距考察期截止天数”为空或不大于0，序号示例：'
        f'{invalid_ids}'
    )

df.rename(columns={'股东比例变化': '股东数量变化率'}, inplace=True)
for compatibility_column in ('是否相符_ins', '是否相符_shareholder', '是否相符_public'):
    df[compatibility_column] = pd.NA
if '主观评价分数' not in df.columns:
    df['主观评价分数'] = np.nan

raw_return_rate_columns = [
    '三个月回报率', '同行业三个月回报',
    '六个月回报率', '同行业六个月回报',
    '考察期回报率', '同行业考察期回报率',
]
df[raw_return_rate_columns] = (
    df[raw_return_rate_columns].apply(pd.to_numeric, errors='coerce') / 100
)
excel_numeric_columns = [
    '距考察期截止天数', '推荐/目标价格', '扣分',
    '三个月回报率_年化', '同行业三个月年化回报', '超额回报率三个月',
    '六个月回报率_年化', '同行业六个月年化回报', '超额回报率六个月',
    '考察期回报率_年化', '同行业考察期年化回报率', '超额回报率考察期',
]
df[excel_numeric_columns] = df[excel_numeric_columns].apply(
    pd.to_numeric, errors='coerce'
)
for observation_date_column in ('中期观察日', '长期观察日'):
    df[observation_date_column] = pd.to_datetime(
        pd.to_numeric(df[observation_date_column], errors='coerce'),
        unit='D',
        origin='1899-12-30',
        errors='coerce',
    )
buyer_ratio_columns = [
    '机构比例变化原始', '股东比例变化原始',
    '机构比例变化', '股东数量变化率',
]
df[buyer_ratio_columns] = (
    df[buyer_ratio_columns]
    .apply(pd.to_numeric, errors='coerce') / 100
)
df['头部公募基金变化'] = pd.to_numeric(
    df['头部公募基金变化'], errors='coerce'
)

if df['标的股票名称'].isna().any():
    invalid_ids = df.loc[df['标的股票名称'].isna(), '序号'].head(10).tolist()
    raise ValueError(f'同行业回报率文件的标的股票名称为空，序号示例：{invalid_ids}')

required_excel_results = [
    '距考察期截止天数', '扣分',
    '三个月回报率_年化', '同行业三个月年化回报', '超额回报率三个月',
    '六个月回报率_年化', '同行业六个月年化回报', '超额回报率六个月',
    '考察期回报率_年化', '同行业考察期年化回报率', '超额回报率考察期',
]
missing_excel_results = df[required_excel_results].isna().any(axis=1)
if missing_excel_results.any():
    invalid_ids = df.loc[missing_excel_results, '序号'].head(10).tolist()
    raise ValueError(f'Excel公式结果仍为空，序号示例：{invalid_ids}')

if isobject:
    domain_4 = pd.read_excel(r'../重要文件/名单/四大领域分析师.xlsx')
    # 第1行为中文字段名，第2行为英文接口字段；计算时跳过英文说明行。
    result_subjective = pd.read_excel(subjective_file, skiprows=[1])
    if '主观评价分数' in df.columns.to_list():
        df = df.drop(columns=['主观评价分数'])

path_br = str(process_output_dir / f'买方反应{nowdate}{extractdate}.xlsx')
path_er = str(process_output_dir / f'超额回报率{nowdate}{extractdate}.xlsx')
path_sum = str(
    process_output_dir / f'最终分数汇总{nowdate}{extractdate}{object}.xlsx'
)
file_path_all = str(
    result_output_dir / f'计算排名结果{nowdate}{method}{extractdate}{object}.xlsx'
)
path_er_examination = str(
    Path(path_er).with_name(f'{Path(path_er).stem}_考察期.xlsx')
)
path_sum_examination = str(
    Path(path_sum).with_name(f'{Path(path_sum).stem}_考察期.xlsx')
)
file_path_all_examination = str(
    Path(file_path_all).with_name(f'{Path(file_path_all).stem}_考察期.xlsx')
)


def require_identical_output_columns(current_frame, examination_frame, stage_name):
    """两个口径只允许数值不同，输出列名和列顺序必须完全一致。"""
    current_columns = current_frame.columns.tolist()
    examination_columns = examination_frame.columns.tolist()
    if current_columns != examination_columns:
        raise RuntimeError(
            f'{stage_name}两个版本的列名或列顺序不一致，已停止输出：'
            f'现有口径={current_columns}；考察期口径={examination_columns}'
        )


#%%    1、买方反应
def buyer_response(group):
    #决定使用排名赋分
    """mean_inst = group['机构比例变化'].mean()  
    std_inst = group['机构比例变化'].std()  
    group['机构比例变化_Z'] = ((group['机构比例变化'] - mean_inst) / std_inst)  
    group['机构比例变化_100'] = (group['机构比例变化_Z'] - group['机构比例变化_Z'].min()) / (group['机构比例变化_Z'].max() - group['机构比例变化_Z'].min()) * 100  #"""
    group['机构比例变化_Z'] = 0
    group['机构比例排名'] = 0
    group['机构比例变化_100'] = 0
    score = group[(~group['机构比例变化'].isnull())].copy()
    score['机构比例变化_Z'] = (score['机构比例变化']-score['机构比例变化'].mean())/score['机构比例变化'].std()
    score['机构比例排名'] = score['机构比例变化_Z'].rank(method='dense', ascending=False)
    n = score['机构比例排名'].max()
    score['机构比例变化_100'] = (100 - (score['机构比例排名']-1)*50/(n-1)).fillna(0)
    group.loc[score.index, '机构比例变化_Z'] = score['机构比例变化_Z']
    group.loc[score.index, '机构比例变化_100'] = score['机构比例变化_100']
    group.loc[group['机构比例变化'] < 0, '机构比例变化_100'] = 0
    group.loc[group['机构比例变化'] == 0, '机构比例变化_100'] = 50
    group.loc[score.index, '机构比例排名'] = score['机构比例排名']
        
    """max_share = group['股东数量变化率'].max()  
    min_share = group['股东数量变化率'].min()  
    group['股东数量变化率_100'] = ((group['股东数量变化率'] - min_share) / (max_share - min_share) * 100).fillna(0)#"""
    group['股东数量变化率_Z'] = 0
    group['股东数量变化率排名'] = 0
    group['股东数量变化率_100'] = 0
    score = group[(~group['股东数量变化率'].isnull())].copy()
    score['股东数量变化率_Z'] = (score['股东数量变化率']-score['股东数量变化率'].mean())/score['股东数量变化率'].std()
    score['股东数量变化率排名'] = score['股东数量变化率_Z'].rank(method='dense',ascending=False)
    n = score['股东数量变化率排名'].max()
    score['股东数量变化率_100'] = (100 - (score['股东数量变化率排名']-1)*50/(n-1)).fillna(0)
    group.loc[score.index, '股东数量变化率_Z'] = score['股东数量变化率_Z']
    group.loc[score.index, '股东数量变化率_100'] = score['股东数量变化率_100']
    group.loc[group['股东数量变化率'] < 0, '股东数量变化率_100'] = 0
    group.loc[group['股东数量变化率'] == 0, '股东数量变化率_100'] = 50
    group.loc[score.index, '股东数量变化率排名'] = score['股东数量变化率排名']
    
    """mean_fund = group['头部公募基金变化'].mean()  
    std_fund = group['头部公募基金变化'].std()  
    group['头部公募基金变化_Z'] = ((group['头部公募基金变化'] - mean_fund) / std_fund)  
    group['头部公募基金变化_100'] = ((group['头部公募基金变化_Z'] - group['头部公募基金变化_Z'].min()) / (group['头部公募基金变化_Z'].max() - group['头部公募基金变化_Z'].min()) * 100).fillna(0)#"""   
    group['头部公募基金变化_Z'] = 0
    group['头部公募基金排名'] = 0
    group['头部公募基金变化_100'] = 0    
    score = group[(~group['头部公募基金变化'].isnull())].copy()
    score['头部公募基金变化_Z'] = (score['头部公募基金变化']-score['头部公募基金变化'].mean())/score['头部公募基金变化'].std()
    score['头部公募基金排名'] = score['头部公募基金变化_Z'].rank(method='dense',ascending=False)
    n = score['头部公募基金排名'].max()
    score['头部公募基金变化_100'] = (100 - (score['头部公募基金排名']-1)*50/(n-1)).fillna(0)
    group.loc[score.index, '头部公募基金变化_Z'] = score['头部公募基金变化_Z']
    group.loc[score.index, '头部公募基金变化_100'] = score['头部公募基金变化_100']
    group.loc[group['头部公募基金变化'] < 0, '头部公募基金变化_100'] = 0
    group.loc[group['头部公募基金变化'] == 0, '头部公募基金变化_100'] = 50
    group.loc[score.index, '头部公募基金排名'] = score['头部公募基金排名']
    
    group['买方反应'] = 0.15*group['机构比例变化_100'] + 0.1*group['股东数量变化率_100'] + 0.05 * group['头部公募基金变化_100']
    return group  
  
df_br = df[['序号', '研究领域清洗版', '看涨/看跌', '权重',
            '是否相符_ins', '机构比例变化原始', '机构比例变化',
            '是否相符_shareholder', '股东比例变化原始', '股东数量变化率',
            '是否相符_public', '头部公募基金变化']].copy()

result_response = df_br.groupby('研究领域清洗版').apply(buyer_response).reset_index(drop=True)  

result_response = result_response[['序号','研究领域清洗版','看涨/看跌','权重','是否相符_ins',
                                   '机构比例变化原始','机构比例变化','机构比例变化_Z','机构比例排名','机构比例变化_100',
                                   '是否相符_shareholder','股东比例变化原始','股东数量变化率','股东数量变化率_Z',
                                   '股东数量变化率排名','股东数量变化率_100','是否相符_public',
                                   '头部公募基金变化','头部公募基金变化_Z','头部公募基金排名',
                                   '头部公募基金变化_100','买方反应']]

print(result_response.head())
result_response.to_excel(path_br, index=False)
#%%    2、超额回报率的计算
#提取同行业数据
# try:
#     weight = df[['序号','权重','推荐/目标价格','6个月内最高价']].copy()
#     df_pct = df_pct.merge(weight,on='序号',how='left')
# except Exception as e:
#     print('error', e)
#     weight = df[['序号','权重']]
#     df_pct = df_pct.merge(weight,on='序号',how='left')
    
# df_pct = df_pct[df_pct['序号'].isin(df['序号'])]

df_pct = df.copy()

# def excess_return(group):
#     #对Z_score加权后排名赋分，最高100分，最低0分
#     group['超额回报率三个月年化'] = group['三个月回报率_年化'] - group['同行业三个月年化回报']
#     mean3 = group['超额回报率三个月年化'].mean()
#     std3 = group['超额回报率三个月年化'].std()
#     group['超额回报率三个月年化Z'] = ((group['超额回报率三个月年化']-mean3)/std3)*group['权重']
#     group['超额年化回报率三个月排名'] = group['超额回报率三个月年化Z'].rank(method='dense',ascending=False)
#     n = group['超额年化回报率三个月排名'].max()
#     group['超额年化回报率三个月得分'] = (100 - (group['超额年化回报率三个月排名']-1)*100/n).fillna(100)
#
#     group['超额回报率六个月年化'] = group['六个月回报率_年化'] - group['同行业六个月年化回报']
#     mean6 = group['超额回报率六个月年化'].mean()
#     std6 = group['超额回报率六个月年化'].std()
#     group['超额回报率六个月年化Z'] = ((group['超额回报率六个月年化']-mean6)/std6)*group['权重']
#     group['超额年化回报率六个月排名'] = group['超额回报率六个月年化Z'].rank(method='dense',ascending=False)
#     n = group['超额年化回报率六个月排名'].max()
#     group['超额年化回报率六个月得分'] = (100 - (group['超额年化回报率六个月排名']-1)*100/n).fillna(100)
#     return group


def excess_return(group):
    # 对Z_score加权后排名赋分，最高100分，最低0分
    # group['超额回报率三个月年化'] = group['三个月回报率_年化'] - group['同行业三个月年化回报']
    mean3 = group['超额回报率三个月'].mean()
    std3 = group['超额回报率三个月'].std()
    group['超额回报率三个月年化Z'] = ((group['超额回报率三个月'] - mean3) / std3) * group['权重']
    group['超额年化回报率三个月排名'] = group['超额回报率三个月年化Z'].rank(method='dense', ascending=False)
    n = group['超额年化回报率三个月排名'].max()
    group['超额年化回报率三个月得分'] = (100 - (group['超额年化回报率三个月排名'] - 1) * 100 / n).fillna(100)

    # group['超额回报率六个月年化'] = group['六个月回报率_年化'] - group['同行业六个月年化回报']
    mean6 = group['超额回报率六个月'].mean()
    std6 = group['超额回报率六个月'].std()
    group['超额回报率六个月年化Z'] = ((group['超额回报率六个月'] - mean6) / std6) * group['权重']
    group['超额年化回报率六个月排名'] = group['超额回报率六个月年化Z'].rank(method='dense', ascending=False)
    n = group['超额年化回报率六个月排名'].max()
    group['超额年化回报率六个月得分'] = (100 - (group['超额年化回报率六个月排名'] - 1) * 100 / n).fillna(100)

    mean_examination = group['超额回报率考察期'].mean()
    std_examination = group['超额回报率考察期'].std()
    group['超额回报率考察期年化Z'] = (
        (group['超额回报率考察期'] - mean_examination)
        / std_examination
    ) * group['权重']
    group['超额年化回报率考察期排名'] = group[
        '超额回报率考察期年化Z'
    ].rank(method='dense', ascending=False)
    n = group['超额年化回报率考察期排名'].max()
    group['超额年化回报率考察期得分'] = (
        100
        - (group['超额年化回报率考察期排名'] - 1) * 100 / n
    ).fillna(100)

    return group


#对于有目标价格的，6个月内最高价超过目标价格的给10分，未达到给0分
df_pct['是否达到目标价格'] = 0 
df_pct.loc[df_pct['推荐/目标价格'].isnull(),'是否达到目标价格'] = -1
df_pct.loc[(df_pct['推荐/目标价格']<=df_pct['6个月内最高价']),'是否达到目标价格'] = 10
df_pct['分组方式'] = 2
df_pct.loc[df_pct['距考察期截止天数']<182,'分组方式'] = 1

if appe in ['-1']:#现在想来还是三者分别算，然后再分别按排名得分会比较公平，即不用分第二次组
    result_return_base = df_pct.groupby(
        ['研究领域清洗版', '分组方式']
    ).apply(excess_return).reset_index(drop=True)
else:
    result_return_base = df_pct.groupby(
        '研究领域清洗版'
    ).apply(excess_return).reset_index(drop=True)


def calculate_excess_return_total(frame, under_90_score_column):
    """按同一权重规则计算超额回报率总分，仅切换不满90天的得分来源。"""
    output = frame.copy()
    output['超额回报率分数1'] = output['超额年化回报率三个月得分'] * 0.25
    at_least_six_months = output['距考察期截止天数'] >= 182
    under_three_months = output['距考察期截止天数'] < 90
    output.loc[at_least_six_months, '超额回报率分数1'] = (
        output['超额年化回报率三个月得分'] * 0.12
        + output['超额年化回报率六个月得分'] * 0.18
    )
    output.loc[under_three_months, '超额回报率分数1'] = (
        output[under_90_score_column] * 0.2
    )
    output['超额回报率分数'] = output['超额回报率分数1'] * 4 / 3
    has_target_price = output['是否达到目标价格'] != -1
    output.loc[has_target_price, '超额回报率分数'] = (
        output['是否达到目标价格'] + output['超额回报率分数1']
    )
    output['超额回报率分数'] = (
        output['超额回报率分数'] + output['扣分'].fillna(0)
    ).clip(lower=0)
    return output


# 现有口径保持不变：不满90天使用三个月得分。
result_return = calculate_excess_return_total(
    result_return_base,
    '超额年化回报率三个月得分',
)
# 考察期口径：仅不满90天改用考察期得分，其他权重、推荐价格和扣分不变。
result_return_examination = calculate_excess_return_total(
    result_return_base,
    '超额年化回报率考察期得分',
)

print(result_return.head())

return_output_columns = ['序号', '推荐日期清洗版', '研究领域清洗版', '距考察期截止天数', '中期观察日', '长期观察日',
        '权重',
        '三个月回报率', '同行业三个月回报', '三个月回报率_年化', '同行业三个月年化回报',
        '超额回报率三个月', '超额回报率三个月年化Z', '超额年化回报率三个月排名','超额年化回报率三个月得分',
        '六个月回报率', '同行业六个月回报', '六个月回报率_年化', '同行业六个月年化回报',
        '超额回报率六个月', '超额回报率六个月年化Z', '超额年化回报率六个月排名', '超额年化回报率六个月得分',
        '考察期回报率', '同行业考察期回报率', '考察期回报率_年化',
        '同行业考察期年化回报率', '超额回报率考察期',
        '超额回报率考察期年化Z', '超额年化回报率考察期排名',
        '超额年化回报率考察期得分', '分组方式',
        '推荐/目标价格', '6个月内最高价', '是否达到目标价格',
        '研报发布时标的价格', '推荐时收盘价', '扣分',
        '超额回报率分数1', '超额回报率分数']
result_return = result_return[return_output_columns]
result_return_examination = result_return_examination[return_output_columns]
require_identical_output_columns(
    result_return,
    result_return_examination,
    '超额回报率过程表',
)

result_return.to_excel(path_er,index=False)
result_return_examination.to_excel(path_er_examination, index=False)
#%%    3、合并主观评价分数（如果有）【if strat != '1':】
if isobject:
    result_subjective = result_subjective.rename(columns={'研究领域':'研究领域清洗版'})
    result_subjective.info()
    matching_subject = ['分析师姓名', '公司名称', '研究领域清洗版', '主观评价分数']
    match = ['分析师姓名', '公司名称', '研究领域清洗版']
    domain_4 = domain_4.merge(result_subjective[matching_subject], on=match, how='left')
    domain_4.rename(columns = {'研究领域清洗版':'研究领域'})
    domain_4['主观评价分数'].fillna(0, inplace=True)
    df = df.merge(result_subjective[matching_subject], on=match, how='left')
    df['主观评价分数'].fillna(0, inplace=True)
#%%    4、排名算分
#先将模板提出
df2 = df[['序号','分析师姓名','公司名称','主观评价分数','唯一标识','研究领域清洗版','标的股票代码','标的股票名称','推荐日期清洗版','推荐强度清洗版','看涨/看跌','权重','推荐/目标价格','6个月内最高价','推荐时收盘价','研报发布时标的价格','扣分','机构比例变化原始','机构比例变化','股东比例变化原始','股东数量变化率','头部公募基金变化']].copy()
print('df2')
print(df2.head())
time.sleep(4)
df_concat = []
df_group = df2.groupby('唯一标识')
for name,group in df_group:
    group['推荐股票总数'] = group['标的股票代码'].nunique()
    df_concat.append(group)
df2 = pd.concat(df_concat)
if isobject:
   df2 = pd.concat([df2,domain_4])

# ##修改--取消资格的名单
# check = ['刘乐文','竺劲']
# df2 = df2[~df2['分析师姓名'].isin(check)]

#买方反应
reaction = result_response[['序号','机构比例排名','机构比例变化_100','股东数量变化率排名','股东数量变化率_100','头部公募基金排名','头部公募基金变化_100','买方反应']]

#超额回报率
return0_columns = ['序号','三个月回报率','三个月回报率_年化', '同行业三个月年化回报','超额回报率三个月', '超额回报率三个月年化Z', '超额年化回报率三个月排名', '超额年化回报率三个月得分',
                   '同行业三个月回报',
                   '六个月回报率','六个月回报率_年化', '同行业六个月年化回报', '超额回报率六个月', '超额回报率六个月年化Z','超额年化回报率六个月排名','超额年化回报率六个月得分',
                   '同行业六个月回报',
                   '考察期回报率', '同行业考察期回报率', '距考察期截止天数', '考察期回报率_年化',
                   '同行业考察期年化回报率', '超额回报率考察期',
                   '超额回报率考察期年化Z', '超额年化回报率考察期排名',
                   '超额年化回报率考察期得分','是否达到目标价格',
                   '超额回报率分数1','超额回报率分数','分组方式']
detail_columns = ['序号', '分析师姓名', '公司名称', '唯一标识', '研究领域', '标的股票代码', '标的股票名称','推荐日期', '推荐强度',
               '看涨/看跌', '权重','机构比例变化原始','机构比例变化','机构比例排名','机构比例变化_100', '股东比例变化原始','股东数量变化率','股东数量变化率排名','股东数量变化率_100','头部公募基金变化', '头部公募基金排名','头部公募基金变化_100', '买方反应',
               '三个月回报率', '三个月回报率_年化', '同行业三个月年化回报','超额回报率三个月', '超额回报率三个月年化Z', '超额年化回报率三个月排名', '超额年化回报率三个月得分',
               '同行业三个月回报',
               '六个月回报率', '六个月回报率_年化', '同行业六个月年化回报', '超额回报率六个月', '超额回报率六个月年化Z','超额年化回报率六个月排名','超额年化回报率六个月得分',
               '同行业六个月回报',
               '考察期回报率', '同行业考察期回报率', '距考察期截止天数', '考察期回报率_年化',
               '同行业考察期年化回报率', '超额回报率考察期',
               '超额回报率考察期年化Z', '超额年化回报率考察期排名',
               '超额年化回报率考察期得分','分组方式','超额回报率(未合并推荐分数)','推荐/目标价格', '6个月内最高价', '是否达到目标价格','研报发布时标的价格','推荐时收盘价','扣分',
               '超额回报率', '客观评价分数',
               '主观评价分数(原始)','推荐股票总数']


def build_scored_result(return_frame):
    """两个口径分别计算，始终使用同一套正式字段名。"""
    return0 = return_frame[return0_columns]
    output = df2.merge(reaction, on='序号', how='left').merge(
        return0, on='序号', how='left'
    )
    output.rename(columns={
        '超额回报率分数1': '超额回报率(未合并推荐分数)',
        '超额回报率分数': '超额回报率',
        '研究领域清洗版': '研究领域',
        '推荐日期清洗版': '推荐日期',
        '推荐强度清洗版': '推荐强度',
        '主观评价分数': '主观评价分数(原始)',
    }, inplace=True)
    output['客观评价分数'] = (
        output['买方反应'] + output['超额回报率']
    ).fillna(0)
    output = output[detail_columns]
    output['客观评价分数(平均)'] = output.groupby(
        '唯一标识'
    )['客观评价分数'].transform('mean')
    output['主观评价分数(原始)'] = output[
        '主观评价分数(原始)'
    ].fillna(0)
    output['主观评价分数'] = output['主观评价分数(原始)'] * 0.3
    output['得分'] = output['客观评价分数'] + output['主观评价分数']
    output['平均得分'] = (
        output['客观评价分数(平均)'] + output['主观评价分数']
    )
    output['推荐日期'] = output['推荐日期'].dt.strftime('%Y-%m-%d')
    output.loc[
        output['主观评价分数'] == 0,
        ['主观评价分数', '主观评价分数(原始)'],
    ] = np.nan
    return output.sort_values(
        by=['唯一标识', '客观评价分数'], ascending=False
    )


result = build_scored_result(result_return)
result_examination = build_scored_result(result_return_examination)
print('超额回报率统计：')
print(result['超额回报率'].describe())
print('买方反应统计：')
print(result['买方反应'].describe())

result2_columns = ['分析师姓名', '公司名称', '得分', '平均得分', '研究领域', '标的股票代码', '标的股票名称', '推荐日期', '距考察期截止天数',
       '推荐强度', '看涨/看跌', '权重', '机构比例变化原始', '机构比例变化', '股东比例变化原始', '股东数量变化率', '头部公募基金变化', '机构比例变化_100',
       '股东数量变化率_100', '头部公募基金变化_100', '买方反应',
       '三个月回报率', '同行业三个月回报', '超额年化回报率三个月得分',
       '六个月回报率', '同行业六个月回报', '超额年化回报率六个月得分',
       '考察期回报率', '同行业考察期回报率', '超额年化回报率考察期得分',
       '推荐/目标价格', '6个月内最高价', '是否达到目标价格',
       '研报发布时标的价格', '推荐时收盘价', '扣分',
       '超额回报率', '客观评价分数', '客观评价分数(平均)',
       '主观评价分数','推荐股票总数']
result2 = result[result2_columns].copy()
result2_examination = result_examination[result2_columns].copy()
require_identical_output_columns(
    result2,
    result2_examination,
    '最终分数汇总表',
)
    

#按照研究领域>平均得分>得分的顺序排列
if not isobject:
    result2 = result2.sort_values(by=['研究领域','客观评价分数(平均)','客观评价分数'],ascending=False)
    result2_examination = result2_examination.sort_values(
        by=['研究领域', '客观评价分数(平均)', '客观评价分数'],
        ascending=False,
    )
else:
    result2 = result2.sort_values(by=['研究领域','平均得分','得分'],ascending=False)
    result2_examination = result2_examination.sort_values(
        by=['研究领域', '平均得分', '得分'],
        ascending=False,
    )

#直接复制用于提取同行业回报率的字典代码

if yr != 2024:
    ind = [[] for _ in range(35)]
    industryID = {
        '宏观': ind[0], '策略': ind[1], '金融工程': ind[2], '固定收益': ind[3], '农林牧渔': ind[4],
        '基础化工': ind[5], '钢铁': ind[6], '有色金属': ind[7], '电子': ind[8], '汽车': ind[9],
        '家用电器': ind[10], '食品饮料': ind[11], '纺织服饰': ind[12], '轻工制造': ind[13],
        '医药生物': ind[14], '公用事业': ind[15], '交通运输': ind[16], '房地产': ind[17],
        '商贸零售': ind[18], '社会服务': ind[19], '银行': ind[20], '非银金融': ind[21],
        '综合': ind[22], '建筑材料': ind[23], '建筑装饰': ind[24], '电力设备': ind[25],
        '机械设备': ind[26], '国防军工': ind[27], '计算机': ind[28], '传媒': ind[29],
        '通信': ind[30], '煤炭': ind[31], '石油石化': ind[32], '环保': ind[33], '美容护理': ind[34]
    }

else:
    ind = [[] for _ in range(32)]
    


def write_score_summary(frame, target_path):
    with pd.ExcelWriter(target_path, engine='openpyxl') as writer:
        for ind in industryID.keys():
            result_ind = frame[frame['研究领域'] == ind]
            result_ind = result_ind.sort_values(by=['平均得分'], ascending=False)
            result_ind.index = result_ind['平均得分'].rank(
                method='dense', ascending=False
            )
            result_ind.index.name = '排名'
            result_ind.to_excel(writer, sheet_name=f'{ind}')


write_score_summary(result2, path_sum)
write_score_summary(result2_examination, path_sum_examination)

match = ['分析师姓名', '公司名称', '研究领域']
#暂且保留每位分析师的最高分股票到最终呈现的排名
result2 = result2.drop_duplicates(subset=match)#默认保留最上面那条，也就是分最高的那只股票
result2_examination = result2_examination.drop_duplicates(subset=match)

# 最终排名与汇总严格使用同一套列名和列顺序。
result3_columns = result2_columns.copy()
result3 = result2[result3_columns]
result3_examination = result2_examination[result3_columns]
require_identical_output_columns(
    result3,
    result3_examination,
    '最终排名结果',
)


def write_ranking_result(frame, target_path):
    check_with_retry(target_path)
    with pd.ExcelWriter(target_path, engine='openpyxl') as writer:
        for ind in industryID.keys():
            result_ind = frame[frame['研究领域'] == ind]
            if not isobject:
                if ind in ['固定收益','金融工程','宏观','策略']:
                    result_ind = result_ind.sort_values(by='得分',ascending=False)
                    result_ind = result_ind[['分析师姓名','公司名称','得分','研究领域']]
                    result_ind['得分'] = result_ind['得分']*10/3
                    result_ind.index = result_ind['得分'].rank(method='dense',ascending=False)
                else:
                    if not switch_method: # 算术平均
                        result_ind = result_ind.sort_values(by=['客观评价分数(平均)'],ascending=False)
                        result_ind.index = result_ind['客观评价分数(平均)'].rank(method='dense',ascending=False)
                    else:
                        result_ind = result_ind.sort_values(by=['客观评价分数'],ascending=False)
                        result_ind.index = result_ind['客观评价分数'].rank(method='dense',ascending=False)
            else:
                if ind in ['固定收益','金融工程','宏观','策略']:
                    result_ind = result_ind.sort_values(by='得分',ascending=False)
                    result_ind = result_ind[['分析师姓名','公司名称','得分','研究领域']]
                    result_ind['得分'] = result_ind['得分']*10/3
                    result_ind.index = result_ind['得分'].rank(method='dense',ascending=False)
                else:
                    if not switch_method:
                        result_ind = result_ind.sort_values(by=['平均得分'],ascending=False)
                        result_ind.index = result_ind['平均得分'].rank(method='dense',ascending=False)
                    else:
                        result_ind = result_ind.sort_values(by=['得分'],ascending=False)
                        result_ind.index = result_ind['得分'].rank(method='dense',ascending=False)
            result_ind.index.name = '排名'
            result_ind.to_excel(writer, sheet_name=f'{ind}')


write_ranking_result(result3, file_path_all)
write_ranking_result(result3_examination, file_path_all_examination)


"""import sys
if not isobject:#暂时不进行后续步骤
    end = time.time()
    print(f'花了{round(end-start,3)}秒')
    sys.exit(0)"""
#%%    5、更改文件呈现方式
from openpyxl import load_workbook  
from openpyxl.styles import PatternFill  
from openpyxl.utils import get_column_letter

def set_column_number_format(sheet, column_letter, number_format='0.00'):  
    """设置指定列的数值格式"""  
    for cell in sheet[column_letter]:  
        if cell.value is not None:  
            cell.number_format = number_format  
            
def set_column_pct_format(sheet, column_letter, number_format='0.00%'):  
    """设置指定列的数值格式"""  
    for cell in sheet[column_letter]:  
        if cell.value is not None:  
            cell.number_format = number_format 

def estimate_column_width(sheet, column_letter, default_width=10):  
    """基于内容长度估算列宽（非常近似的方法）"""  
    max_length = 0  
    for cell in sheet[column_letter]:  
        if cell.value is not None:  
            cell_content = str(cell.value)  
            cell_length = len(cell_content)  
            if cell_length > max_length:  
                max_length = cell_length  
    # 返回一个稍微大于最长内容长度的宽度（加上额外的缓冲）  
    return max(default_width, (max_length + 2))  
  
def highlight_and_adjust_columns(
    file_path, highlight_columns, adjust_only_columns, number_columns
):  
    """  
    对多个特定列进行高亮和列宽调整。  
    highlight_columns: 需要高亮的列字母列表。  
    adjust_only_columns: 仅需要调整列宽的列字母列表，这些列不会被高亮。  
    number_columns: 需要按普通数值显示并调整列宽的列字母列表。
    """  
    wb = load_workbook(file_path)  
    hl_columns = highlight_columns
    ajo_columns = adjust_only_columns
    for sheet_name in wb.sheetnames:  
        sheet = wb[sheet_name]  
        if sheet_name not in ['固定收益','金融工程','宏观','策略']:
            adjust_only_columns = ajo_columns
            highlight_columns = hl_columns
            active_number_columns = number_columns
            # 冻结窗格位置
            sheet.freeze_panes = 'F7'
        if sheet_name in ['固定收益','金融工程','宏观','策略']:
            highlight_columns = ['D']
            active_number_columns = []
            # 冻结窗格位置
            sheet.freeze_panes = 'E7'
        for column_letter in highlight_columns:  
            cell = sheet[f"{column_letter}1"]  
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  
            cell.fill = fill  
  
            set_column_number_format(sheet, column_letter)  
  
            estimated_width = estimate_column_width(sheet, column_letter)  
            sheet.column_dimensions[column_letter].width = estimated_width  
  
        for column_letter in adjust_only_columns:  
            estimated_width = estimate_column_width(sheet, column_letter)  
            sheet.column_dimensions[column_letter].width = estimated_width
            set_column_pct_format(sheet, column_letter)#设置百分比格式

        for column_letter in active_number_columns:
            estimated_width = estimate_column_width(sheet, column_letter)
            sheet.column_dimensions[column_letter].width = estimated_width
            set_column_number_format(sheet, column_letter)

  
    wb.save(file_path)  
    
file_paths = [file_path_all, file_path_all_examination]


def final_output_column_letter(column_name):
    """最终展示版首列为排名，因此数据字段需要整体右移一列。"""
    return get_column_letter(result3_columns.index(column_name) + 2)


highlight_columns = [
    final_output_column_letter(column_name)
    for column_name in [
        '买方反应', '平均得分', '得分', '超额回报率',
        '客观评价分数', '客观评价分数(平均)', '主观评价分数',
    ]
]
adjust_columns = [
    final_output_column_letter(column_name)
    for column_name in [
        '机构比例变化原始', '机构比例变化',
        '股东比例变化原始', '股东数量变化率',
        '三个月回报率', '同行业三个月回报',
        '六个月回报率', '同行业六个月回报',
        '考察期回报率', '同行业考察期回报率',
    ]
]  # 回报率列按百分比显示
number_columns = [
    final_output_column_letter(column_name)
    for column_name in [
        '推荐/目标价格', '6个月内最高价',
        '研报发布时标的价格', '推荐时收盘价',
    ]
]
center_columns = [
    'A',
    *[
        final_output_column_letter(column_name)
        for column_name in [
            '分析师姓名', '得分', '平均得分', '研究领域',
            '标的股票代码', '标的股票名称', '距考察期截止天数',
            '买方反应', '扣分', '超额回报率', '客观评价分数',
            '客观评价分数(平均)', '主观评价分数',
        ]
    ],
]

for file_path in file_paths:
    highlight_and_adjust_columns(
        file_path, highlight_columns, adjust_columns, number_columns
    )
    auto_fit_columns(file_path, file_path)
    format_excel_columns(file_path,
            {tuple(center_columns): {'center': True}})

end = time.time()
print(f'总共花费{round(end-start,3)}秒')
print(f'现有三个月口径结果：{file_path_all}')
print(f'不满三个月使用考察期口径结果：{file_path_all_examination}')

open_file(file_path_all)
open_file(file_path_all_examination)
