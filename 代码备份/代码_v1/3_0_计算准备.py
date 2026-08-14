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

make_dir('3_万得提取')

sdir = []
for subdirs,dirs,files in os.walk('2_数据已提取'):
    i = 1
    for dir in dirs:
        try:
            dir = int(dir[-4:])
            sdir.append(dir)
        except:
            continue   
    if i == 1: break
appe = str(max(sdir))

appe = str(appe)

if len(appe) == 3:
   appe = '0' + appe

print(appe)

import time
time.sleep(1)

path = rf'2_数据已提取/1.数据准备_{appe}'
path2 = rf'2_数据已提取/数据提取_{appe}'

make_dir(r'3_万得提取/数据提取')
make_dir(r'../重要文件/模板')
make_dir(r'../重要文件/主观评分')
# copy_file(r'../资料/代码/10_计算.py', r'../step_4 计算/2024分析师计算.py')
# copy_file(r'../资料/说明/分析师金牛奖计算步骤2024.docx', r'../step_4 计算/分析师金牛奖计算细节.docx')

if os.path.exists(path):
    print('1')
    #复制必要文件到计算步骤
    copy_file(path + rf'/1_股票代码清洗与数据提取/分析师清洗版_{appe}.xlsx',
              rf'../重要文件/分析师清洗版/分析师清洗版_{appe}.xlsx')
    copy_file(path + rf'/0_同行业回报率industry/已提取数据_{appe}.xlsx',
              rf'../重要文件/同行业回报率/已提取数据_{appe}.xlsx')
    #复制文件到数据更新步骤
    copy_file(path + r'/1_股票代码清洗与数据提取/分析师定量数据模板.xlsx',
              r'3_万得提取/数据提取/分析师定量数据模板.xlsx')
    copy_file(path + r'/0_同行业回报率industry/1_提取同行业回报率.ipynb',
              r'3_万得提取/数据提取/1_提取同行业回报率.ipynb')
    copy_file(r'../资料/说明/数据提取说明.docx', r'3_万得提取/数据提取/0_数据提取说明.docx')
    copy_file(fr"../资料/模板/{yr}/新建模板(1).xlsx", r'3_万得提取/数据提取/新建模板(1).xlsx')
    copy_file(r"../资料/说明/同行业回报率说明_2.docx", r'3_万得提取/数据提取/同行业回报率说明.docx')
    #保存已清洗数据至必要文件的文件夹
    copy_file(path + r'/1_股票代码清洗与数据提取/取消资格.xlsx', r'../重要文件/名单/取消资格.xlsx')
    copy_file(path + r'/1_股票代码清洗与数据提取/分析师定量数据模板.xlsx', r'../重要文件/模板/分析师定量数据模板.xlsx')
    if not os.path.exists(r'../重要文件/头部公募基金'):
        shutil.copytree(path + r'/2_头部公募基金/头部公募基金', r'../重要文件/头部公募基金')
    print('success')
    time.sleep(1.2)
    os.startfile(os.path.join('..','step_4 计算'))
elif os.path.exists(path2):
    print('2')
    copy_file(path2 + rf'/分析师清洗版_{appe}.xlsx',
              rf'../step_4 计算/0_准备数据/分析师清洗版/分析师清洗版_{appe}.xlsx')
    copy_file(path2 + rf'/已提取数据_{appe}.xlsx', rf'../step_4 计算/0_准备数据/同行业回报率/已提取数据_{appe}.xlsx')
    print('success')
    time.sleep(1.2)
    open_file(os.path.join('..','step_4 计算'))
else:
    print('error, 文件夹不存在!请观察命名是否正确或者跟随文件夹内步骤')
    time.sleep(2)

