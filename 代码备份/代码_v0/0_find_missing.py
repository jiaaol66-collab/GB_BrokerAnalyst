import sys, os, time, warnings
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
        time.sleep(2)
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

warnings.filterwarnings('ignore')

rel = find_rel(__file__, '参评信息')
start = insert_bookmark(__file__,0, rel)
impt = find_rel(__file__, '重要文件')
mtr = find_rel(__file__, '资料')
processed = find_rel(__file__, '已处理')
org = find_rel(__file__, '原始数据集')
yr = get_year(rel)

src =f'{org}{yr}/报名表/数据资料'
 
dst = f'{processed}{yr}/数据资料'
path = f'../step_1 手动补齐报名表/'
#复制选定目录的所有文件到此目录
if not os.path.exists(dst):
    print('---------------正在复制文件{}--------------'.format(datetime.now()))
    copy_file(src, dst)

try:
    user = pd.read_excel(f'{org}{yr}/用户报名表.xlsx')
except:
    print('用户报名表无法读取...请跟随说明进行操作...')

path2 = r'../step_2 分析师股票提取清洗信息'
make_dir(path2+r'/1_初步提取')

print('---------------正在寻找缺失表头的报名表{}--------------'.format(datetime.now()))

no_head = []
def traverse(root_folder):
    # 遍历根文件夹下的所有子文件夹和文件
    for subdir, dirs, files in os.walk(root_folder):
        folder_name = os.path.basename(subdir)  
        for file in files:
            # 检查文件名中是否包含'.xlsx'
            if '.xlsx' in file.lower():
                file_path = os.path.join(subdir, file)
                # 尝试读取Excel文件
                try:
                    xls = pd.ExcelFile(file_path, engine='openpyxl')
                    sheet_names = xls.sheet_names  # 获取所有工作表名
                    # 遍历所有工作表
                    for sheet_name in sheet_names:
                        df = pd.read_excel(xls,sheet_name=sheet_name)
                        if '分析师个人信息表' not in df.columns:
                            if (df.iloc[0][0] == '分析师所属部门' or df.iloc[0][0] == '分析师所在机构名称'):
                                no_head.append([folder_name ,file_path.split('//')[-1],sheet_name])
                except:
                    pass

traverse(src)
file_path = path + r'1_缺失表头报名表位置.txt'

delete_file(file_path)
with open(file_path,'w',encoding='utf-8') as f:
    for i in range(len(no_head)):
        f.write(no_head[i][0]+'_'+no_head[i][1]+'_'+no_head[i][2]+'\n')
        print(no_head[i][0],no_head[i][1],no_head[i][2])
print('success')

copy_file(f'{org}/{yr}/报名表/数据资料/', f'{processed}/{yr}/数据资料')


open_file('../step_2 分析师股票提取清洗信息/1_初步提取')
open_file('../step_1 手动补齐报名表/0_手动补齐说明.docx')
open_file(f'{processed}/{yr}/数据资料')

insert_bookmark(__file__, 1, rel, start)