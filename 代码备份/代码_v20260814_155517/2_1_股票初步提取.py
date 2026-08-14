#%%    1_初步提取(核心步骤，务必事先处理原文件，因为表格格式不一定一模一样，所以CIY)
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

rel = find_rel(__file__, '参评信息')
start = insert_bookmark(__file__,0, rel)
impt = find_rel(__file__, '重要文件')
mtr = find_rel(__file__, '资料')
processed = find_rel(__file__, '已处理')
org = find_rel(__file__, '原始数据集')
yr = get_year(rel)
make_dir('输出文件')

analysist = []
leader = []
stock_code = []
stock_name = []
recommend_date = []
recommend = []
at_moment_price = []
recommend_price = []
institute_number = []
domains = []
titles = []
err = []
no_name = []
source_records = []


STOCK_VALUE_COLUMN_INDEXES = (1, 2, 3, 4, 8, 9, 11)


def is_empty_stock_value(value):
    """股票填报字段的空占位判断；中间步骤统一以“-”保留。"""
    if pd.isna(value):
        return True
    return str(value).strip().lower() in {'', '-', 'nan', 'none', 'nat'}


def select_stock_row_indices(df, start_row_index, max_stock_rows=10):
    """
    固定检查股票表头后的全部可用填报行，不把中间空行当作结束标记。

    - 任一股票字段存在内容：保留最多10行，包含其中的空行；
    - 代码、名称、日期、标题、强度及价格全部为空：只保留第一条空占位；
    - 表头后没有实际行：返回一个 None，由调用方生成一条空占位。
    """
    stock_row_count = min(
        max_stock_rows,
        max(0, len(df) - start_row_index),
    )
    row_indices = [start_row_index + offset for offset in range(stock_row_count)]
    if not row_indices:
        return [None]

    has_any_stock_value = any(
        any(
            not is_empty_stock_value(df.iloc[row_index, column_index])
            for column_index in STOCK_VALUE_COLUMN_INDEXES
        )
        for row_index in row_indices
    )
    return row_indices if has_any_stock_value else row_indices[:1]

def traverse_folder_stock(root_folder):
    # 遍历根文件夹下的所有子文件夹和文件
    for subdir, dirs, files in os.walk(root_folder):
        folder_name = os.path.basename(subdir)
        v = 0
        for file in files:
            # 检查文件是否为支持的 Excel 格式
            if '~$' in file:
                continue
            if Path(file).suffix.lower() in {'.xlsx', '.xls'}:
                file_path = os.path.join(subdir, file)
                absolute_file_path = str(Path(file_path).resolve())
                sheet_name = '-'
                #print(f"找到文件: {file_path}")
                # 尝试读取Excel文件
                try:
                    xls = pd.ExcelFile(file_path) # 不要随意设engine
                    sheet_names = xls.sheet_names  # 获取所有工作表名
                    # 遍历所有工作表
                    v = v + 1

                    print_progress(v, len(files), 50, f'正在读取{folder_name}..')

                    b = folder_name
                    for sheet_name in sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name,dtype={'Unnamed: 1':str,'Unnamed: 3':str})
                        #如果报名表过长，则说明有问题需要查看
                        #print(f"正在处理工作表: {sheet_name}")
                        if '分析师个人信息表' not in df.columns:
                            #print('不是报名表')
                            continue
                        #%%检查有没有存在重叠的报名表
                        if len(df)>50:
                            err.append([
                                folder_name,
                                os.path.basename(file_path),
                                sheet_name,
                                absolute_file_path,
                                '报名表行数超过50行',
                            ])
                        #%%CIY,对编号37的机构单独设计的读取位置,对应问题总结的情况1.II
                        # if folder_name == '37':
                        #     name1 = '-'
                        #     name2 = '-' if pd.isna(df.iloc[6,1]) else df.iloc[6,1]#去查看它的位置
                        #     #print(name2)
                        # else:
                        name1 = '-' if pd.isna(df.iloc[8,1]) else df.iloc[8,1]
                        #print(name1)
                        name2 = '-' if pd.isna(df.iloc[11,1]) else df.iloc[11,1]
                        # if folder_name in ['48']:#CIY
                        #     name2 = '-'
                        #print(name2)
                        domain = '-' if pd.isna(df.iloc[2,1]) else df.iloc[2,1]
                        #%%如果没提取到姓名，则记录
                        if name1 == '-' and name2 == '-' and domain != '-':
                            no_name.append([
                                folder_name,
                                os.path.basename(file_path),
                                sheet_name,
                                absolute_file_path,
                            ])
                        #%%定位读取
                        #定位所需信息的行数
                        search_string = '标的股票代码'#CIY
                        if search_string in df['Unnamed: 1'].to_list():
                            mask = df['Unnamed: 1'].str.contains(search_string, case=False, na=False)
                            matching_rows = df[mask].index.tolist()
                            j = matching_rows[0] + 1
                            # 固定扫描最多10个股票填报行。不能用中间空行作为结束标记：
                            # 报名表可能前面为空、后面才填股票代码。
                            stock_row_indices = select_stock_row_indices(df, j)
                            for stock_row_index in stock_row_indices:
                                analysist.append(name1)
                                leader.append(name2)
                                domains.append(domain)
                                stockCode = '-' if (
                                    stock_row_index is None
                                    or pd.isna(df.iloc[stock_row_index, 1])
                                ) else df.iloc[stock_row_index, 1]
                                stockCode = str(stockCode)#先将其强行改为str型
                                #print(stockCode)
                                stock_code.append(stockCode)
                                stockName = '-' if (
                                    stock_row_index is None
                                    or pd.isna(df.iloc[stock_row_index, 2])
                                ) else df.iloc[stock_row_index, 2]
                                #print(stockName)
                                stock_name.append(stockName)
                                recommendDate = '-' if (
                                    stock_row_index is None
                                    or pd.isna(df.iloc[stock_row_index, 3])
                                ) else df.iloc[stock_row_index, 3]
                                recommendDate = str(recommendDate)
                                #print(recommendDate)
                                recommend_date.append(recommendDate)
                                atmomentprice = '-' if (
                                    stock_row_index is None
                                    or pd.isna(df.iloc[stock_row_index, 8])
                                ) else df.iloc[stock_row_index, 8]
                                #print(atmomentprice)
                                at_moment_price.append(atmomentprice)
                                recommendprice = '-' if (
                                    stock_row_index is None
                                    or pd.isna(df.iloc[stock_row_index, 9])
                                ) else df.iloc[stock_row_index, 9]
                                #print(recommendprice)
                                recommend_price.append(recommendprice)
                                recommendation = '-' if (
                                    stock_row_index is None
                                    or pd.isna(df.iloc[stock_row_index, 11])
                                ) else df.iloc[stock_row_index, 11]
                                #print(recommendation)
                                recommend.append(recommendation)
                                title = '-' if (
                                    stock_row_index is None
                                    or pd.isna(df.iloc[stock_row_index, 4])
                                ) else df.iloc[stock_row_index, 4]
                                titles.append(title)
                                #print(title)
                                institute_number.append(folder_name)
                                source_records.append({
                                    '机构哈希': folder_name,
                                    '文件名': os.path.basename(file_path),
                                    '工作表': sheet_name,
                                    '文件路径': absolute_file_path,
                                    '股票信息Excel行': (
                                        '-' if stock_row_index is None
                                        else stock_row_index + 2
                                    ),
                                })
                        else:
                            analysist.append(name1)
                            leader.append(name2)
                            domains.append(domain)
                            stockCode = '-'
                            stockCode = str(stockCode)
                            #print(stockCode)
                            stock_code.append(stockCode)
                            stockName = '-'
                            #print(stockName)
                            stock_name.append(stockName)
                            recommendDate = '-'
                            recommendDate = str(recommendDate)
                            #print(recommendDate)
                            recommend_date.append(recommendDate)
                            atmomentprice = '-'
                            #print(atmomentprice)
                            at_moment_price.append(atmomentprice)
                            recommendprice = '-'
                            #print(recommendprice)
                            recommend_price.append(recommendprice)
                            recommendation = '-'
                            #print(recommendation)
                            recommend.append(recommendation)
                            title = '-' 
                            titles.append(title)
                            #print(title)
                            institute_number.append(folder_name)
                            source_records.append({
                                '机构哈希': folder_name,
                                '文件名': os.path.basename(file_path),
                                '工作表': sheet_name,
                                '文件路径': absolute_file_path,
                                '股票信息Excel行': '-',
                            })
                except Exception as e:
                    print(f"无法读取文件 {file_path}: {e}")
                    err.append([
                        folder_name,
                        os.path.basename(file_path),
                        sheet_name,
                        absolute_file_path,
                        f'无法读取：{e}',
                    ])

# 指定根文件夹的路径
print('---------------正在读取更改后的数据集{}--------------'.format(datetime.now()))
for data_folder_name in ['goldenbull_fxs', 'goldenbull_fxs_ex']:
    data_root = Path(processed) / str(yr) / data_folder_name
    if data_root.is_dir():
        print(f'正在读取：{data_root}')
        traverse_folder_stock(str(data_root))
print('success')

#%%输出
#将读取的数据转化为DataFrame
output = {'分析师':analysist,'组长':leader,'研究领域':domains,'标的股票代码':stock_code,'标的股票名称':stock_name,
          '推荐日期':recommend_date,'推荐强度':recommend,'研报发布时标的价格':at_moment_price,'研报推荐的目标价格':recommend_price,
          '机构编号':institute_number,'研报标题':titles}

df0 = pd.DataFrame(output)

path2 = r'../2_清洗分析师姓名'

make_dir(f'{impt}补充信息')
make_dir(path2)
make_dir('输出文件')
if os.path.exists(f'{processed}{yr}/已提取数据/append.xlsx'):
    copy_file(f'{processed}{yr}/已提取数据/append.xlsx', f'{impt}补充信息/append.xlsx',1)
else:
    copy_file(f'{mtr}模板/append_ept.xlsx', '3_append.xlsx',1)

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


DATE_VALIDATION_FORMATS = [
    '%Y 年 %m 月 %d 日',
    '%Y年%m月%d日',
    '%Y.%m',
    '%Y年%m月%d号',
    '%Y年-%m月-%d日',
    '%Y 年 %m月  %d日',
    '%Y-%m-%d %H:%M',
    '%Y-%m-%d %H:%M:%S',
]
NAME_ALLOWED_PATTERN = re.compile(
    r"^[\u3400-\u4dbf\u4e00-\u9fffA-Za-z"
    r" \t·•・.'’,-，、/／;；&]+$"
)
NAME_CHARACTER_PATTERN = re.compile(r'[\u3400-\u4dbf\u4e00-\u9fffA-Za-z]')
STOCK_CODE_PATTERN = re.compile(
    r'^\(?\s*(?:(?:SH|SZ|BJ|HK)\.?\d{5,6}|'
    r'\d{5,6}\.?(?:SH|SZ|BJ|HK))\s*\)?$',
    re.IGNORECASE,
)
SWAPPED_STOCK_CODE_TOKEN_PATTERN = re.compile(
    r'^(?:(?:SH|SZ|BJ|HK)\.?\d{1,7}|'
    r'(?:\d{1,7}|[A-Z]{1,8})\.?(?:SH|SZ|BJ|HK|O|N)|'
    r'\d{1,7})$',
    re.IGNORECASE,
)
STOCK_CODE_NORMALIZATION_MAP = str.maketrans({
    '（': '(',
    '）': ')',
    '。': '.',
    '．': '.',
    '／': '/',
})
DATE_LIKE_PATTERN = re.compile(
    r'^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}'
    r'(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$'
)
PLACEHOLDER_VALUES = {'', '-', 'nan', 'none', 'nat'}


def normalized_text(value):
    return '' if pd.isna(value) else str(value).strip()


def is_placeholder(value):
    return normalized_text(value).lower() in PLACEHOLDER_VALUES


def is_stock_code_like(value):
    """识别名称列中常见的 A/H/美股代码写法，用于提示代码与名称疑似填反。"""
    value_text = normalized_text(value).upper().translate(STOCK_CODE_NORMALIZATION_MAP)
    value_text = re.sub(r'\s+', '', value_text).strip('()')
    if not value_text:
        return False
    code_parts = [part.strip('()') for part in value_text.split('/')]
    return all(
        part and SWAPPED_STOCK_CODE_TOKEN_PATTERN.fullmatch(part)
        for part in code_parts
    )


field_validation_issues = []
field_validation_issue_keys = set()


def add_field_validation_issue(row_index, field_name, reason, value, excel_row):
    source = source_records[row_index]
    issue_key = (
        source['文件路径'],
        source['工作表'],
        field_name,
        excel_row,
        normalized_text(value),
        reason,
    )
    if issue_key in field_validation_issue_keys:
        return
    field_validation_issue_keys.add(issue_key)
    field_validation_issues.append({
        **source,
        '字段': field_name,
        '错误说明': reason,
        '提取值': normalized_text(value) or '-',
        '来源Excel行': excel_row,
    })


if len(df0) != len(source_records):
    raise ValueError(
        f'提取结果与来源路径数量不一致：结果 {len(df0)} 行，来源 {len(source_records)} 行'
    )

# 姓名只允许中文、英文字母及常见姓名分隔符；日期、股票代码或其他字符均提前提示。
for field_name, excel_row in [('分析师', 10), ('组长', 13)]:
    for row_index, value in df0[field_name].items():
        value_text = normalized_text(value)
        if is_placeholder(value_text):
            continue
        if STOCK_CODE_PATTERN.fullmatch(value_text) or DATE_LIKE_PATTERN.fullmatch(value_text):
            add_field_validation_issue(
                row_index,
                field_name,
                '姓名位置疑似填入股票代码或日期',
                value_text,
                excel_row,
            )
        elif (
            not NAME_ALLOWED_PATTERN.fullmatch(value_text)
            or not NAME_CHARACTER_PATTERN.search(value_text)
        ):
            add_field_validation_issue(
                row_index,
                field_name,
                '姓名包含非中英文字符，可能发生错位',
                value_text,
                excel_row,
            )

# 股票代码列出现完整日期，通常说明股票代码与推荐日期发生了列错位。
for row_index, value in df0['标的股票代码'].items():
    value_text = normalized_text(value)
    if not is_placeholder(value_text) and DATE_LIKE_PATTERN.fullmatch(value_text):
        add_field_validation_issue(
            row_index,
            '标的股票代码',
            '股票代码位置疑似填入推荐日期',
            value_text,
            source_records[row_index]['股票信息Excel行'],
        )

# 名称列像股票代码、代码列反而不像代码时，提示两列疑似填反；仅标记，不自动换值。
for row_index, row in df0[['标的股票代码', '标的股票名称']].iterrows():
    stock_code_text = normalized_text(row['标的股票代码'])
    stock_name_text = normalized_text(row['标的股票名称'])
    if is_placeholder(stock_code_text) or is_placeholder(stock_name_text):
        continue
    if is_stock_code_like(stock_name_text) and not is_stock_code_like(stock_code_text):
        add_field_validation_issue(
            row_index,
            '标的股票代码/标的股票名称',
            '股票名称列像股票代码、股票代码列不像代码，疑似两列填反',
            f'{stock_code_text} ↔ {stock_name_text}',
            source_records[row_index]['股票信息Excel行'],
        )

# 推荐日期先识别“填入股票代码”，再检查是否能按现有日期规则解析。
recommend_date_text = df0['推荐日期'].apply(normalized_text)
parsed_recommend_dates = parse_dates(
    recommend_date_text.copy(),
    DATE_VALIDATION_FORMATS,
)
for row_index, value_text in recommend_date_text.items():
    if is_placeholder(value_text):
        continue
    source_excel_row = source_records[row_index]['股票信息Excel行']
    if STOCK_CODE_PATTERN.fullmatch(value_text):
        add_field_validation_issue(
            row_index,
            '推荐日期',
            '推荐日期位置疑似填入股票代码',
            value_text,
            source_excel_row,
        )
    elif pd.isna(parsed_recommend_dates.loc[row_index]):
        add_field_validation_issue(
            row_index,
            '推荐日期',
            '推荐日期格式无法识别',
            value_text,
            source_excel_row,
        )


institution_lookup = {}
institution_index_root = Path(processed) / str(yr) / '机构报名表'
institution_index_candidates = sorted(
    institution_index_root.glob('机构报名表*.xlsx'),
    key=lambda path: path.stat().st_mtime,
    reverse=True,
)
for institution_index_path in institution_index_candidates:
    try:
        institution_index = pd.read_excel(
            institution_index_path,
            dtype={'链接位置占位': str},
        )
    except Exception:
        continue
    required_index_columns = {'机构简称', '机构名称', '链接位置占位'}
    if not required_index_columns.issubset(institution_index.columns):
        continue
    institution_index['链接位置占位'] = (
        institution_index['链接位置占位'].fillna('').astype(str).str.strip()
    )
    institution_lookup = (
        institution_index.loc[
            institution_index['链接位置占位'] != '',
            ['机构简称', '机构名称', '链接位置占位'],
        ]
        .drop_duplicates(subset=['链接位置占位'], keep='first')
        .set_index('链接位置占位')[['机构简称', '机构名称']]
        .to_dict(orient='index')
    )
    break


legacy_error_report_path = '2_信息有问题的报名表.txt'
error_report_path = '2_信息有问题的报名表.md'

# 字段问题按“同一文件 + 同一工作表 + 同一问题类型”合并。
# 同一张表中的不同字段或不同错误说明仍分别写一行。
field_validation_groups = {}
for issue in field_validation_issues:
    group_key = (
        issue['文件路径'],
        issue['工作表'],
        issue['字段'],
        issue['错误说明'],
    )
    field_validation_groups.setdefault(group_key, []).append(issue)

# 排序只看文件和工作表；Python 的稳定排序会保留同一张表内各问题类型的原始顺序。
# 因此不同问题仍分行，但同一文件同一张表的所有问题会连续排列。
field_validation_group_items = sorted(
    field_validation_groups.items(),
    key=lambda item: (
        os.path.normcase(item[0][0]),
        normalized_text(item[0][1]),
    ),
)

for report_path in [legacy_error_report_path, error_report_path]:
    if os.path.exists(report_path):
        os.remove(report_path)

with open(error_report_path, 'w', encoding='utf-8', newline='\n') as f:
    f.write('# 信息有问题的报名表\n\n')
    f.write(f'- 读取失败或结构异常：{len(err)}\n')
    f.write(f'- 分析师姓名缺失：{len(no_name)}\n')
    f.write(
        f'- 字段前置校验异常：{len(field_validation_issues)} 条'
        f'（按文件/工作表/问题类型合并为 {len(field_validation_groups)} 组）\n\n'
    )

    f.write('## 读取失败或结构异常\n\n')
    f.write('| 机构简称 | 机构名称 | 机构哈希 | 错误说明 | 工作表 | 报名表（真实绝对路径） |\n')
    f.write('| --- | --- | --- | --- | --- | --- |\n')
    for institution_hash, file_name, sheet_name, file_path, reason in err:
        institution = institution_lookup.get(str(institution_hash), {})
        institution_short_name = institution.get('机构简称', '-')
        institution_name = institution.get('机构名称', '-')
        file_link = local_absolute_link(file_path)
        f.write(
            f'| {markdown_text(institution_short_name)}'
            f' | {markdown_text(institution_name)}'
            f' | `{markdown_text(institution_hash)}`'
            f' | {markdown_text(reason)}'
            f' | {markdown_text(sheet_name)}'
            f' | [打开：{markdown_text(file_name)}](<{file_link}>) |\n'
        )
        print(institution_short_name, institution_name, file_name, sheet_name, reason)

    f.write('\n## 分析师姓名可能有问题\n\n')
    f.write('| 机构简称 | 机构名称 | 机构哈希 | 工作表 | 报名表（真实绝对路径） |\n')
    f.write('| --- | --- | --- | --- | --- |\n')
    for institution_hash, file_name, sheet_name, file_path in no_name:
        institution = institution_lookup.get(str(institution_hash), {})
        institution_short_name = institution.get('机构简称', '-')
        institution_name = institution.get('机构名称', '-')
        file_link = local_absolute_link(file_path)
        f.write(
            f'| {markdown_text(institution_short_name)}'
            f' | {markdown_text(institution_name)}'
            f' | `{markdown_text(institution_hash)}`'
            f' | {markdown_text(sheet_name)}'
            f' | [打开：{markdown_text(file_name)}](<{file_link}>) |\n'
        )
        print(institution_short_name, institution_name, file_name, sheet_name)

    f.write('\n## 字段前置校验异常\n\n')
    f.write(
        '| 机构简称 | 机构名称 | 机构哈希 | 问题明细'
        ' | 工作表 | 报名表（真实绝对路径） |\n'
    )
    f.write('| --- | --- | --- | --- | --- | --- |\n')
    for (file_path, sheet_name, _field_name, _reason), issues in field_validation_group_items:
        first_issue = issues[0]
        institution_hash = first_issue['机构哈希']
        institution = institution_lookup.get(str(institution_hash), {})
        institution_short_name = institution.get('机构简称', '-')
        institution_name = institution.get('机构名称', '-')
        file_link = local_absolute_link(file_path)
        problem_details = '<br>'.join(
            f'Excel行 {markdown_text(issue["来源Excel行"])}：'
            f'{markdown_text(issue["字段"])}；'
            f'{markdown_text(issue["错误说明"])}；'
            f'值 `{markdown_text(issue["提取值"])}`'
            for issue in issues
        )
        f.write(
            f'| {markdown_text(institution_short_name)}'
            f' | {markdown_text(institution_name)}'
            f' | `{markdown_text(institution_hash)}`'
            f' | {problem_details}'
            f' | {markdown_text(sheet_name)}'
            f' | [打开：{markdown_text(first_issue["文件名"])}](<{file_link}>) |\n'
        )
        print(institution_short_name, institution_name, sheet_name, f'{len(issues)} 条问题')

file_path = '输出文件/output_初步提取.xlsx'
check_with_retry(file_path)
df0.to_excel(file_path,index=False)
copy_file(file_path,f'{impt}股票信息/1_初步提取.xlsx', 1)
print('success')

open_file(file_path)
open_file('3_append.xlsx')

insert_bookmark(__file__, 1, rel, start)
