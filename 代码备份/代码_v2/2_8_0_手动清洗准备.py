import sys, os, time
from pathlib import Path


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

#%%去重、更改顺序，输出清洗完毕后文件，后可制作成模板文件
df3 = pd.read_excel(f'{impt}股票信息/7_已初步清洗股票代码.xlsx')
df3 = df3.sort_values(by=['机构编号','分析师姓名'])
df3['序号'] = None
df3['序号'].iloc[0] = 1
df3 = df3[['序号','分析师姓名','唯一标识','公司名称','研究领域','研究领域清洗版','标的股票代码(原)','标的股票代码(清洗)','标的股票代码','标的股票名称(原)','标的股票名称','标的股票名称(对比)','推荐日期','推荐日期清洗版','推荐强度','推荐强度清洗版','看涨/看跌','权重','机构编号','研报发布时标的价格(原)','研报推荐的目标价格(原)','研报推荐的目标价格','研报推荐的目标价格下限','上限','研报发布时标的价格']]
df3.info()
print(len(df3[df3.duplicated(keep=False)]),df3[df3.duplicated(keep=False)])
df3 = df3.sort_values(by=['机构编号','分析师姓名'])
df3.drop_duplicates(inplace=True)#对应情况1,共8条
df3.info()
df5 = df3[df3['研究领域清洗版'].isin(['金融工程','固定收益','宏观','策略'])]
df7 = df3[~df3['研究领域清洗版'].isin(['金融工程','固定收益','宏观','策略'])]#对应情况2，共105条
df5 = df5[['分析师姓名','研究领域清洗版','公司名称']]
df5.drop_duplicates(subset=['分析师姓名','公司名称'],inplace=True)


path_sub = f'{impt}名单'
df5.to_excel(path_sub +r'/四大领域分析师.xlsx',index=False)

df6 = df7[df7['标的股票名称(原)'] == '-']
df7 = df7[df7['标的股票名称(原)'] != '-']#对应情况3，共34条
df7.info()

path3 = '输出文件/1_数据准备/1_股票数据清洗/'
make_dir(path3)
copy_file(f'{mtr}模板/{yr}/模板复制.xlsx', '模板复制.xlsx')
open_file('模板复制.xlsx')
open_file('2_手动清洗说明.docx')
# copy_file(r'../../资料/代码/8.准备数据.py','7.准备数据【按步骤清洗完毕后运行此文件】.py')
copy_file(f'../4_清洗研究领域/2_复数领域分析师.txt', '复数领域分析师.txt')

file_path8 = path3 + '1_股票信息部分清洗版.xlsx'
check_with_retry(file_path8)

# 定义颜色代码
PERIWINKLE_PURPLE = "#E4DFEC"  # 培安紫
CORNFLOWER_BLUE = "#DAEEF3"  # 矢车菊蓝
CHOCOLATE_YELLOW = "#FDE9D9"  # 巧克力黄
SCARLET_RED = "#F9DBDF"  # 猩红

# 创建Excel文件
with pd.ExcelWriter(file_path8, engine='xlsxwriter') as writer:
    df7.to_excel(writer, sheet_name='Sheet1', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    # 获取数据范围（从第2行开始到最后一行）
    last_row = len(df7) + 1  # +1是因为标题行

    for i in range(len(df7)): # 设置K列公式
        row_index = i + 2
        # 消除J列额外的空格以及转义符
        formula = f'=CLEAN(TRIM(J{row_index}))'
        worksheet.write_formula(f'K{row_index}', formula)

    for i in range(len(df7)): # 设置I列公式
        row_index = i + 2
        # 由K列提取的WIND股票代码
        formula = f'=to_windcode(K{row_index})'
        worksheet.write_formula(f'I{row_index}', formula)

    for i in range(len(df7)): # 设置L列公式
        row_index = i + 2
        #
        formula = f'=s_info_name1(G{row_index},N{row_index})'
        worksheet.write_formula(f'L{row_index}', formula)

    # 1. H列条件格式：=LEFT(H2,7)<>LEFT(I2,7) 粗略的展示股票代码前部的差异
    format_periwinkle = workbook.add_format({'bg_color': PERIWINKLE_PURPLE})
    worksheet.conditional_format(f'H2:H{last_row}', {
        'type': 'formula',
        'criteria': f'=LEFT(H2,7)<>LEFT(I2,7)',
        'format': format_periwinkle
    })

    # 2. I列条件格式：=I2<>H2 表示提取的代码与step.2 第7步生成的代码有任何差异
    format_cornflower = workbook.add_format({'bg_color': CORNFLOWER_BLUE})
    worksheet.conditional_format(f'I2:I{last_row}', {
        'type': 'formula',
        'criteria': f'=I2<>H2',
        'format': format_cornflower
    })

    # 3. J列条件格式：=NOT(ISNUMBER(SEARCH(K2,L2))) 粗略展示K2与L2之间的字段不存在包含关系
    format_chocolate = workbook.add_format({'bg_color': CHOCOLATE_YELLOW})
    worksheet.conditional_format(f'J2:J{last_row}', {
        'type': 'formula',
        'criteria': f'=NOT(ISNUMBER(SEARCH(K2,L2)))',
        'format': format_chocolate
    })

    # 4. 整个工作表条件格式：=(CELL("row")=ROW()) 简易高亮行工具，建议常按F9更新位置
    format_scarlet = workbook.add_format({'bg_color': SCARLET_RED})

    # 获取工作表的列范围（从A到最后一列）
    last_col_letter = chr(65 + len(df7.columns) - 1)

    # 应用整个工作表的条件格式
    worksheet.conditional_format(f'A2:{last_col_letter}{last_row}', {
        'type': 'formula',
        'criteria': '=(CELL("row")=ROW())',
        'format': format_scarlet
    })

open_file(file_path8)

check_with_retry(path3 + '取消资格.xlsx')
df6.to_excel(path3 + '取消资格.xlsx', index=False)
open_file(path3 + '取消资格.xlsx')

