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
df3 = pd.read_excel(f'{impt}股票信息/4_已清洗研究领域.xlsx')

print(df3['推荐强度'].value_counts())

#清洗推荐强度
#买入（强烈推荐）增持（推荐）
#中性（观望）减持（卖出）
def recommend(row):
    #如果名称为空(即'-'),那么推荐强度缺失就必须补上，这里补上的是增持（推荐）
    if row['标的股票名称']!='-':
        if row['推荐强度'] in ['强烈推荐','优大于市','优于大市','买入','买入（强烈推荐）','买入评级',
                           '首推','买入（强烈推荐)','买入(强烈推荐)','强','买入-A ','强推']:#CIY
            return '买入（强烈推荐）'
        elif row['推荐强度'] in ['增持','增持（推荐）','推荐','看好（行业报告）','持有','是','增持-A','谨慎推荐']:#CIY
            return '增持（推荐）'
        elif row['推荐强度'] in ['一般','中性','中性（观望）']:#CIY
            return '中性（观望）'
        elif row['推荐强度'] in ['减持','减持（卖出）']:#CIY
            return '减持（卖出）'
        else:
            return '增持（推荐）'
    
df3['推荐强度清洗版'] = df3.apply(recommend,axis=1)

#对应规则中的描述
def label(row):
    if row['推荐强度清洗版'] == '买入（强烈推荐）':
        row['看涨/看跌'] = 1
        row['权重'] = 1
    elif row['推荐强度清洗版'] == '减持（卖出）':
        row['看涨/看跌'] = -1
        row['权重'] = 1
    elif row['推荐强度清洗版'] == '增持（推荐）':   
        row['看涨/看跌'] = 1
        row['权重'] = 0.8
    elif row['推荐强度清洗版'] == '中性（观望）':
        row['看涨/看跌'] = -1
        row['权重'] = 0.8
    else:
        row['看涨/看跌'] = 0
        row['权重'] = 0
    return row

df3 = df3.apply(label,axis=1)

path6 = r'../6_清洗价格'
make_dir(path6)

file_path5 = '输出文件/output_分析师已清洗推荐强度.xlsx'
df3.to_excel(file_path5,index=False)

print('success')

open_file(file_path5)

copy_file(file_path5, f'{impt}股票信息/5_已清洗推荐强度.xlsx',1)

insert_bookmark(__file__, 1, rel, start)
