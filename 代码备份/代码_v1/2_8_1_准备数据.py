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

path3 = r'../../step_3 提取万得数据/1_数据准备'
path_m = mtr

# copy_file(path_m + r'/提取万得数据说明.docx', r'../../step_3 提取万得数据/0.提取万得数据说明.docx')
# copy_file(path_m + r'/数据准备说明.docx', path3 + r'/0.数据准备说明.docx')
#复制文件至提取行业回报率的0.文件夹
# make_dir(path3 + r'/0.同行业回报率industry')
# copy_file(path_m + r'/同行业回报率说明_1.docx', path3 + r'/0.同行业回报率industry/0.同行业回报率说明.docx')
# copy_file(path_m + r'/代码/step3 提取同行业回报率(得到机房提取).ipynb', path3 + r'/0.同行业回报率industry/1.提取同行业回报率.ipynb')
copy_file('4_分析师定量数据.xlsx', path3 + r'/分析师定量数据.xlsx')
copy_file(path_m + rf'/模板/{yr}/新建模板(1).xlsx', path3 + r'/0.同行业回报率industry/新建模板(1).xlsx')

#复制文件至清洗股票代码与数据提取的1.文件夹
# copy_file(path_m + r'/股票清洗步骤.docx', path3 + r'/1.股票代码清洗与数据提取/0.股票清洗步骤.docx')
copy_file('4_分析师定量数据.xlsx', path3 + r'/分析师定量数据.xlsx')

#复制文件至提取头部公募基金的2.文件夹
# make_dir(path3 + r'/2.头部公募基金/头部公募基金')
# make_dir(r'../../step_3 提取万得数据/2_数据已提取')
# copy_file(path_m + '/头部公募基金说明.docx', path3 + '/2.头部公募基金/0.头部公募基金说明.docx')
# copy_file(path_m + '/代码/step3 头部公募基金提取.ipynb', path3 + '/2.头部公募基金/1.头部公募基金提取.ipynb')
# copy_file(path_m + '/公募基金复制列.xlsx', path3 + '/2.头部公募基金/公募基金复制列.xlsx' )
# copy_file(path_m + '/代码/3_0_计算准备.py', '../../step_3 提取万得数据/2.计算准备.py')
print('success')
import time
time.sleep(3)
open_file(os.path.join('..','..','step_3 提取万得数据'))