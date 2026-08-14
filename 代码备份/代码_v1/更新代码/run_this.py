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

rel = '../参评信息/'
make_dir(rel, 1)
start = insert_bookmark(__file__, 0, rel)
org = find_rel(__file__, '原始数据集')
mtr = find_rel(__file__, '资料')


""" 建立工作区文件夹 """
info = pd.read_excel(f'{mtr}代码/工作区设置/文件结构.xlsx',sheet_name='工作区')
info.info()
info = info[info['备注'].isna()]

subdir = info.drop_duplicates(subset=['工作区文件夹','子文件夹1','子文件夹2','子文件夹3','子文件夹4'])[
    ['步骤', '工作区文件夹', '子文件夹1', '子文件夹2', '子文件夹3', '子文件夹4']].copy()

subdir = subdir.fillna('.')
print(subdir.head(20))

print(f'创建工作区文件夹.. {datetime.now()}'.center(80,'-'))
for idx, row in subdir.iterrows():
    base_dir = 'step_' + str(row['步骤']) + ' ' +row['工作区文件夹'] if row['步骤'] != '.' else row['工作区文件夹']
    sub_dir1 = row['子文件夹1']
    sub_dir2 = row['子文件夹2']
    sub_dir3 = row['子文件夹3']
    sub_dir4 = row['子文件夹4']
    make_dir(f'../{base_dir}/{sub_dir1}/{sub_dir2}/{sub_dir3}/{sub_dir4}')

impt = find_rel(__file__, '重要文件')
tag_year = 0
now = datetime.now()
upper_year = now.year + 1


"""—————————————————————————— 自定义代码块 ——————————————————————————————"""

if not os.path.exists(f'{rel}年份信息.txt'):
    """ 自定义 field """
    field = [
        {
            'label': '评奖年份',
            'default': f'{now.year}',
            'hint': f'请输入2024-{upper_year}之间的年份(包括2024)',
            'validate': lambda v: validate_numeric(v, lower=2024, upper=upper_year, type_='int')
        },
        {
            'label': '数据集',
            'default': '20250928',
            'hint': '输入待提取的数据集信息',
            'validate': lambda v: validate_numeric(v,  lower = 19000101, upper= 22001231, type_='int')
        }
    ]
    result = custom_input_dialog('请输入新工作区的必需内容', field)
    tag_year = 1
    yr = int(result['评奖年份'])

    with open(f'{rel}年份信息.txt','w',encoding='utf-8') as f:
        f.write(f'{yr}')

    datasource = int(result['数据集'])

    with open(f'{rel}数据集信息.txt','w',endcoding='utf-8') as f:
        f.write(f'{datasource}')

    """↓ 自定义代码块，目的是与原始数据集产生连接，在此可以自定义操作 ↓"""

else:
    yr = get_year(rel)

# 无需清除缓冲区的步骤, 默认为'.'
step_folders = ['.', '前期准备']

"""—————————————————————————— 自定义代码块 ——————————————————————————————"""


print()
print(f'正在部署代码.. {datetime.now()}'.center(80,'-'))
code = info[info['保存在资料中的名称']!='.'][['步骤', '工作区文件夹','子文件夹1', '子文件夹2', '相应代码和说明文件', '代码/说明',
                                              '代码子文件夹','保存在资料中的名称']].copy()

code = code.fillna('.')
for idx, row in code.iterrows():
    step = 'step_' + str(row['步骤']) + ' ' + row['工作区文件夹'] if row['步骤'] != '.' else '.'
    dst_name = row['相应代码和说明文件']
    sav_name = row['保存在资料中的名称']
    sub_dir1 = row['子文件夹1']
    sub_dir2 = row['子文件夹2']
    sub_code = row['代码子文件夹']
    code_istr = row['代码/说明']
    tof = copy_file(f'{mtr}{code_istr}/{sub_code}/{sav_name}', f'../{step}/{sub_dir1}/{sub_dir2}/{dst_name}',
                    0, 0, 1)
    if not tof:
        if code_istr == '代码':
            copy_file(f'{mtr}{code_istr}/default_code.py', f'../{step}/{sub_dir1}/{sub_dir2}/{dst_name}')
        elif code_istr == '说明':
            copy_file(f'{mtr}{code_istr}/readme.docx', f'../{step}/{sub_dir1}/{sub_dir2}/{dst_name}')
        else:
            pass

    copy_file(f'{mtr}代码/run.bat', f'../{step}/{sub_dir1}/{sub_dir2}/run.bat', 0)
    if row['工作区文件夹'] not in step_folders:
        copy_file(f'{mtr}代码/flush_buffer.py', f'../{step}/{sub_dir1}/{sub_dir2}/flush_buffer.py', 0)

print()
insert_bookmark(__file__, 1, rel, start)