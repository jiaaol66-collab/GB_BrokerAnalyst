#%%    1_初步提取(核心步骤，务必事先处理原文件，因为表格格式不一定一模一样，所以CIY)
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

analysist = []
leader = []
stock_code = []
stock_name = []
recommend_date = []
recommend = []
at_moment_price = []
recommend_price = []
institute_number = []
domains = []
titles = []
err = []
no_name = []

def traverse_folder_stock(root_folder):
    # 遍历根文件夹下的所有子文件夹和文件
    for subdir, dirs, files in os.walk(root_folder):
        folder_name = os.path.basename(subdir)
        v = 0
        for file in files:
            # 检查文件名中是否包含'.xlsx'
            if '~$' in file:
                continue
            if '.xlsx' in file.lower():
                file_path = os.path.join(subdir, file)
                #print(f"找到文件: {file_path}")
                # 尝试读取Excel文件
                try:
                    xls = pd.ExcelFile(file_path) # 不要随意设engine
                    sheet_names = xls.sheet_names  # 获取所有工作表名
                    # 遍历所有工作表
                    v = v + 1

                    print_progress(v, len(files), 50, f'正在读取{folder_name}..')

                    b = folder_name
                    for sheet_name in sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet_name,dtype={'Unnamed: 1':str,'Unnamed: 3':str})
                        #如果报名表过长，则说明有问题需要查看
                        #print(f"正在处理工作表: {sheet_name}")
                        if '分析师个人信息表' not in df.columns:
                            #print('不是报名表')
                            continue
                        #%%检查有没有存在重叠的报名表
                        if len(df)>45:
                            err.append([folder_name ,file_path.split('\\')[-1],sheet_name])
                        #%%CIY,对编号37的机构单独设计的读取位置,对应问题总结的情况1.II
                        # if folder_name == '37':
                        #     name1 = '-'
                        #     name2 = '-' if pd.isna(df.iloc[6,1]) else df.iloc[6,1]#去查看它的位置
                        #     #print(name2)
                        # else:
                        name1 = '-' if pd.isna(df.iloc[8,1]) else df.iloc[8,1]
                        #print(name1)
                        name2 = '-' if pd.isna(df.iloc[11,1]) else df.iloc[11,1]
                        # if folder_name in ['48']:#CIY
                        #     name2 = '-'
                        #print(name2)
                        domain = '-' if pd.isna(df.iloc[2,1]) else df.iloc[2,1]
                        #%%如果没提取到姓名，则记录
                        if name1 == '-' and name2 == '-' and domain != '-':
                            no_name.append([folder_name ,file_path.split('\\')[-1],sheet_name])
                        #%%定位读取
                        #定位所需信息的行数
                        search_string = '标的股票代码'#CIY
                        if search_string in df['Unnamed: 1'].to_list():
                            mask = df['Unnamed: 1'].str.contains(search_string, case=False, na=False)
                            matching_rows = df[mask].index.tolist()
                            j = matching_rows[0] + 1
                            #因最多10个股票所以只最多选取10个(即便他大于10个, 所幸本次没有这个情况)
                            for i in range(10):
                                analysist.append(name1)
                                leader.append(name2)
                                domains.append(domain)
                                stockCode = '-' if pd.isna(df.iloc[j+i,1]) else df.iloc[j+i,1]
                                stockCode = str(stockCode)#先将其强行改为str型
                                #print(stockCode)
                                stock_code.append(stockCode)
                                stockName = '-' if pd.isna(df.iloc[j+i,2]) else df.iloc[j+i,2]
                                #print(stockName)
                                stock_name.append(stockName)
                                recommendDate = '-' if pd.isna(df.iloc[j+i,3]) else df.iloc[j+i,3]
                                recommendDate = str(recommendDate)
                                #print(recommendDate)
                                recommend_date.append(recommendDate)
                                atmomentprice = '-' if pd.isna(df.iloc[j+i,8]) else df.iloc[j+i,8]
                                #print(atmomentprice)
                                at_moment_price.append(atmomentprice)
                                recommendprice = '-' if pd.isna(df.iloc[j+i,9]) else df.iloc[j+i,9]
                                #print(recommendprice)
                                recommend_price.append(recommendprice)
                                recommendation = '-' if pd.isna(df.iloc[j+i,11]) else df.iloc[j+i,11]
                                #print(recommendation)
                                recommend.append(recommendation)
                                title = '-' if pd.isna(df.iloc[j+i,4]) else df.iloc[j+i,4]
                                titles.append(title)
                                #print(title)
                                institute_number.append(folder_name)
                                if pd.isna(df.iloc[j+1+i,2]):#当接下来的标的股票名称为空的时候,就停止。
                                    break
                        else:
                            analysist.append(name1)
                            leader.append(name2)
                            domains.append(domain)
                            stockCode = '-'
                            stockCode = str(stockCode)
                            #print(stockCode)
                            stock_code.append(stockCode)
                            stockName = '-'
                            #print(stockName)
                            stock_name.append(stockName)
                            recommendDate = '-'
                            recommendDate = str(recommendDate)
                            #print(recommendDate)
                            recommend_date.append(recommendDate)
                            atmomentprice = '-'
                            #print(atmomentprice)
                            at_moment_price.append(atmomentprice)
                            recommendprice = '-'
                            #print(recommendprice)
                            recommend_price.append(recommendprice)
                            recommendation = '-'
                            #print(recommendation)
                            recommend.append(recommendation)
                            title = '-' 
                            titles.append(title)
                            #print(title)
                            institute_number.append(folder_name)
                except Exception as e:
                    print(f"无法读取文件 {file_path}: {e}")
                    err.append([folder_name ,file_path.split('\\')[-1],sheet_name])

# 指定根文件夹的路径
print('---------------正在读取更改后的数据集{}--------------'.format(datetime.now()))
traverse_folder_stock(f"{processed}{yr}/数据资料/")
print('success')

#%%输出
#将读取的数据转化为DataFrame
output = {'分析师':analysist,'组长':leader,'研究领域':domains,'标的股票代码':stock_code,'标的股票名称':stock_name,
          '推荐日期':recommend_date,'推荐强度':recommend,'研报发布时标的价格':at_moment_price,'研报推荐的目标价格':recommend_price,
          '机构编号':institute_number,'研报标题':titles}

df0 = pd.DataFrame(output)

path2 = r'../2_清洗分析师姓名'

make_dir(f'{impt}补充信息')
make_dir(path2)
make_dir('输出文件')
if os.path.exists(f'{processed}{yr}/已提取数据/append.xlsx'):
    copy_file(f'{processed}{yr}/已提取数据/append.xlsx', f'{impt}补充信息/append.xlsx',1)
else:
    copy_file(f'{mtr}模板/append_ept.xlsx', '3_append.xlsx',1)

file_path2 = '2_信息有问题的报名表.txt'

if os.path.exists(file_path2):
    os.remove(file_path2)

with open(file_path2,'w',encoding='utf-8') as f:
    f.write('读取不到数据的\n')
    print('读取不到数据的')
    for i in range(len(err)):
        f.write(err[i][0]+'_'+err[i][1]+'_'+err[i][2]+'\n')
        print(err[i][0],err[i][1],err[i][2])
    f.write('分析师姓名可能有问题的\n')
    print('分析师姓名可能有问题的')
    for i in range(len(no_name)):
        f.write(no_name[i][0]+'_'+no_name[i][1]+'_'+no_name[i][2]+'\n')
        print(no_name[i][0],no_name[i][1],no_name[i][2])

file_path = '输出文件/output_初步提取.xlsx'
check_with_retry(file_path)
df0.to_excel(file_path,index=False)
copy_file(file_path,f'{impt}股票信息/1_初步提取.xlsx', 1)
print('success')

open_file(file_path)
open_file('3_append.xlsx')

insert_bookmark(__file__, 1, rel, start)