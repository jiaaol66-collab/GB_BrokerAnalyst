import sys, os, time
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

df3 = pd.read_excel(f'{impt}股票信息/5_已清洗推荐强度.xlsx')

df3.rename(columns={'标的股票代码':'标的股票代码(原)','标的股票名称':'标的股票名称(原)','研报发布时标的价格':'研报发布时标的价格(原)','研报推荐的目标价格':'研报推荐的目标价格(原)'},inplace=True)
df3['机构编号'] = df3['机构编号'].astype(str)
df3['唯一标识'] = df3['分析师姓名'] + df3['机构编号']#作为唯一标识分析师的依据
# df3['机构编号'] = df3['机构编号'].astype(int)
df3[['标的股票代码(清洗)','标的股票代码','标的股票名称','标的股票名称(对比)','研报发布时标的价格','研报推荐的目标价格','研报推荐的目标价格下限','上限']] = None
df3['标的股票代码(清洗)'] = df3['标的股票代码(原)']

#%%推荐价格与标的价格的清洗(CIY)
#清洗推荐目标价格和发布时标的价格
def convert(x):
    if isinstance(x,str):#CIY,虽然我估计下面这一串能应付大多数情况...
        if '（港股）' in x:
            x = x.split('（港股）')[1]
        x = x.replace(',','.').replace('。','.').replace('，','.').replace('分红后价格','').replace('前复权','').replace('未复权','').replace('不复权','').replace('美金','').replace('美元','').replace('港币','').replace('港元','').replace('元','').replace('/股','').replace(' ','').replace('$','').replace('＄','').replace('RMB','').replace('CNY','').replace('HKD','').replace('HK','').replace('复权','')
        for i in ['(','（']:
            if i in x:
                x = x.split(i)[0]
        try:
            x = float(x)
            return x
        except :
            print(f'{x}不能被更改')
            return x
    else:
        return x

df3['研报发布时标的价格'] = df3['研报发布时标的价格(原)'].apply(lambda x: None if x in ['——','无','-','/'] else x)#预处理
df3['研报发布时标的价格'] = df3['研报发布时标的价格'].apply(convert)
a = df3[(df3['研报发布时标的价格'].apply(lambda x: isinstance(x, float)))&(~df3['研报发布时标的价格(原)'].apply(lambda x:isinstance(x, float)))][['研报发布时标的价格','研报发布时标的价格(原)']]
df3['研报推荐的目标价格'] = df3['研报推荐的目标价格(原)'].apply(lambda x: x.replace('~','-').replace('到','-') if isinstance(x,str) else x)
df3['研报推荐的目标价格'] = df3['研报推荐的目标价格'].apply(lambda x: None if (x in ['——','-','/'] or (isinstance(x,str) and ('未'or'不') in x)) else x)
df3['研报推荐的目标价格下限'] = df3['研报推荐的目标价格'].apply(lambda x: x.split('-')[0] if (isinstance(x,str) and '-' in x) else x)
df3['上限'] = df3['研报推荐的目标价格'].apply(lambda x: x.split('-')[1] if (isinstance(x,str) and '-' in x) else None)
df3['研报推荐的目标价格下限'] = df3['研报推荐的目标价格下限'].apply(convert)
b = df3[(df3['研报推荐的目标价格下限'].apply(lambda x: isinstance(x, float)))&(~df3['研报推荐的目标价格(原)'].apply(lambda x:isinstance(x, float)))][['研报推荐的目标价格下限','研报推荐的目标价格(原)']]
df3['上限'] = df3['上限'].apply(convert)

path7 = r'../7_清洗股票代码'
make_dir(path7)

with open('2_价格清洗.txt','w',encoding = 'utf-8') as f:
    f.write('标的价格\n')
    for i in range(len(a)):
        f.write(f"{i}原:{a.iloc[i]['研报发布时标的价格(原)']},清洗后:{a.iloc[i]['研报发布时标的价格']}\n")
    f.write('推荐价格下限\n')
    for i in range(len(b)):
        f.write(f"{i}原:{b.iloc[i]['研报推荐的目标价格(原)']},清洗后:{b.iloc[i]['研报推荐的目标价格下限']}\n")


file_path6 = '输出文件/output_分析师已初步清洗价格.xlsx'

print('success')

check_with_retry(file_path6)
df3.to_excel(file_path6, index=False)
open_file(file_path6)

copy_file(file_path6, f'{impt}股票信息/6_已初步清洗价格.xlsx', 1)
insert_bookmark(__file__, 1, rel, start)
