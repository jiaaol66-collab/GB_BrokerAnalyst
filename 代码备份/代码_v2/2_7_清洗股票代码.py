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

df3 = pd.read_excel(f'{impt}股票信息/6_已初步清洗价格.xlsx')

#%%股票代码初步清洗(非常粗略，到头来还是要到楼下提取对比的，这里只是简化一下步骤）
def stockcode(code):
    if '.' not in code and code != '-':
        code = code.replace('。','.')#CIY
        if len(code) >= 6:
            code += '.SZ'  #在没加后缀的中股票代码中.SZ占比 较多
        else:
            if len(code)==5:#当长度为5的时候，分两种情况，如果第一个是0，那大概率是港股，否则是内地股票[不过也存在反例。。。]
                if code[0] != '0': 
                    code = '0' + code
                    code += '.SH'
                elif code[0] == '0':
                    code = code[1:]
                    code += '.HK'
            elif len(code) == 4:
                code += '.HK'
            elif len(code) == 3:
                code = '0' + code
                code += '.HK'
            elif len(code) == 2:
                code = '00' + code
                code += '.HK'
            elif len(code) == 1:
                code = '000' + code
                code += '.HK'
    else:
        code = code.upper()

    if '.HK' in code:
        if len(code) == 8:
            code = code[1:]
    return code

df3['标的股票代码(清洗)'] = df3['标的股票代码(清洗)'].apply(stockcode)

path8 = r'../8_输出excel手动清洗'
make_dir(path8)

file_path7 = '输出文件/output_分析师已初步清洗股票代码.xlsx'

print('success')

check_with_retry(file_path7)
df3.to_excel(file_path7, index=False)

open_file(file_path7)

copy_file(file_path7, f'{impt}股票信息/7_已初步清洗股票代码.xlsx', 1)
