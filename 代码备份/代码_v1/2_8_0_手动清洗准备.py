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

df6 = df7[df7['标的股票代码(原)'] == '-']
df7 = df7[df7['标的股票代码(原)'] != '-']#对应情况3，共34条
df7.info()

path3 = '输出文件/1_数据准备/1_股票数据清洗/'
make_dir(path3)
copy_file(f'{mtr}模板/{yr}/模板复制.xlsx', '模板复制.xlsx')
open_file('模板复制.xlsx')
open_file('2_手动清洗说明.docx')
# copy_file(r'../../资料/代码/8.准备数据.py','7.准备数据【按步骤清洗完毕后运行此文件】.py')
copy_file(f'../4_清洗研究领域/2_复数领域分析师.txt', '复数领域分析师.txt')

file_path8 = path3 + '1_股票信息部分清洗版.xlsx'
df7.to_excel(file_path8, index=False)
open_file(file_path8)

df6.to_excel(path3 + '取消资格.xlsx', index=False)
open_file(path3 + '取消资格.xlsx')

