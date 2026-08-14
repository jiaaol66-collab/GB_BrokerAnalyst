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

tag_append = 0

print('---------------正在读取文件{}--------------'.format(datetime.now()))
if os.path.exists(f'{impt}补充信息/append.xlsx'):
    extra = pd.read_excel(f'{impt}补充信息/append.xlsx')
    tag_append = 1
else:
    if os.path.exists(f'../1_初步提取/3_append.xlsx'):
        extra = pd.read_excel(f'../1_初步提取/3_append.xlsx')
        tag_append = 1

if tag_append == 1:
    len2 = len(extra)


df0 = pd.read_excel(f'{impt}股票信息/1_初步提取.xlsx')
len0 = len(df0)

if tag_append == 1:
    df0 = pd.concat([df0,extra],ignore_index=True)

df0['机构编号'] = df0['机构编号'].astype(int)

#%%去空行 选取研究领域不是'-'的所有行，即删去是'-'的行
len1 = len(df0[df0['研究领域'] == '-'])
df0 =  df0[df0['研究领域'] != '-']
#%%将分析师姓名与组长统一到[分析师姓名]列
def name_merge(row):
    row['分析师姓名'] = row['分析师'] if row['分析师'] != '-' else row['组长']
    if row['分析师']!='-'and row['组长']!='-'and row['分析师']!=row['组长']:#对应情况1
        row['分析师姓名'] = row['分析师'] +'、'+str(row['组长'])
    #唐海清和潘暕一个是分析师，一个是组长（...)CIY
    if row['分析师姓名'] =='唐海清':
        row['分析师姓名']='唐海清、潘暕'
    return row

df1 = df0.apply(name_merge,axis=1).copy()

# #更改几个没有提到名字的"核心组员"对应情况2
# mask = (df1['标的股票代码'] == '600060.SH')&(df1['分析师姓名'] == '-')
# df1.loc[mask,'分析师姓名'] = '王刚'
# mask = (df1['分析师姓名'] == '-')&(df1['机构编号'] != 12)
# df1.loc[mask,'分析师姓名'] = '李少青'

# #更改写反了的情况，对应情况3
# mask = df1['分析师姓名']=='宗建树'
# a = df1.loc[mask,'标的股票代码']
# b = df1.loc[mask,'标的股票名称']
# df1.loc[df1['分析师姓名']=='宗建树','标的股票名称'] = a
# df1.loc[df1['分析师姓名']=='宗建树','标的股票代码'] = b


print(f'提取到{len0}条分析师信息,通过删除空行删去{len1}条信息',end='')
if tag_append == 1:
    print(f'上一步补充{len2}条信息')


path3 = r'../3_清洗日期'
make_dir(path3)

file_path2 = '输出文件/output_分析师已清洗姓名.xlsx'

check_with_retry(file_path2)
df1.to_excel(file_path2, index=False)
copy_file(file_path2, f'{impt}股票信息/2_已清洗姓名.xlsx',1)

print('success')

open_file(file_path2)

insert_bookmark(__file__, 1, rel, start)
