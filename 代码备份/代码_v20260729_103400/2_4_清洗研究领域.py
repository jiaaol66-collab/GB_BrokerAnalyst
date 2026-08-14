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

df2 = pd.read_excel(f'{impt}股票信息/3_初步清洗日期.xlsx')
print(df2['研究领域'].unique())

if yr == 2024:
    def domain(row):
        #计算机、通信和其他电子设备制造业
        if row['研究领域'] in ['计算机、通信和其他电子设备制造业','计算机','电子','通信','C39']:#CIY
            return '计算机、通信和其他电子设备制造业'
        #机械设备制造业
        elif row['研究领域'] in ['机械设备制造业','机械设备','专用设备制造业']:#CIY
            return '机械设备制造业'
        #电气机械和器材制造业
        elif row['研究领域'] in ['电气机械和器材制造业','C38 电气机械和器材制造业','家用电器']:#CIY
            return '电气机械和器材制造业'
        #基础化工制造业
        elif row['研究领域'] in ['基础化工制造业','基础化工','基础化工制造业（C26）']:#CIY
            return '基础化工制造业'
        #医药制造业
        elif row['研究领域'] in ['医药制造业','医药生物']:#CIY
            return '医药制造业'
        #汽车制造业
        elif row['研究领域'] in ['汽车','汽车制造业']:#CIY
            return '汽车制造业'
        #轻工制造业
        elif row['研究领域'] in ['轻工制造','轻工制造业']:#CIY
            return '轻工制造业'
        #非金属矿物制品业
        elif row['研究领域'] in ['非金属矿物制品业']:#CIY
            return '非金属矿物制品业'
        #金属制品业
        elif row['研究领域'] in ['金属制品业','有色金属']:#CIY
            return '金属制品业'
        #金属冶炼和压延加工业
        elif row['研究领域'] in ['金属冶炼和压延加工业','有色金属冶炼和压延加工业']:#CIY
            return '金属冶炼和压延加工业'
        #铁路、船舶、航空航天和其他运输设备制造业
        elif row['研究领域'] in ['铁路、船舶、航空航天和其他运输设备制造业','国防军工','15']:#CIY
            return '铁路、船舶、航空航天和其他运输设备制造业'
        #食品饮料制造加工业
        elif row['研究领域'] in ['食品饮料制造加工业','食品饮料','食品饮料加工制造业']:#CIY
            return '食品饮料制造加工业'
        #纺织服饰制造业
        elif row['研究领域'] in ['纺织服饰制造业','纺织服装、服饰业','C18']:#CIY
            return '纺织服饰制造业'
        #其他制造业
        elif row['研究领域'] in ['其他制造业']:#CIY
            return '其他制造业'
        #信息传输、软件和信息技术服务业
        elif row['研究领域'] in ['信息传输、软件和信息技术服务业']:#CIY
            return '信息传输、软件和信息技术服务业'
        #批发和零售业-无需清洗
        elif row['研究领域'] in ['批发和零售业']:#CIY
            return '批发和零售业'
        #电力、热力、燃气及水生产和供应业
        elif row['研究领域'] in ['电力、热力、燃气及水生产和供应业','电力设备新能源']:#CIY
            return '电力、热力、燃气及水生产和供应业'
        #金融业
        elif row['研究领域'] in ['金融业','非银金融','银行业']:#CIY
            return '金融业'
        #科学研究和技术服务业
        elif row['研究领域'] in ['科学研究和技术服务业']:#CIY
            return '科学研究和技术服务业'
        #交通运输、仓储和邮政业
        elif row['研究领域'] in ['交通运输','交通运输、仓储和邮政业']:#CIY
            return '交通运输、仓储和邮政业'
        #建筑业
        elif row['研究领域'] in ['建筑业','建筑材料','建筑']:#CIY
            return '建筑业'
        #房地产业
        elif row['研究领域'] in ['房地产业','房地产']:#CIY
            return '房地产业'
        #水利、环境和公共设施管理业
        elif row['研究领域'] in ['水利、环境和公共设施管理业']:#CIY
            return '水利、环境和公共设施管理业'
        #采矿业
        elif row['研究领域'] in ['采矿业','煤炭']:#CIY
            return '采矿业'
        #租赁和商务服务业
        elif row['研究领域'] in ['租赁和商务服务业']:#CIY
            return '租赁和商务服务业'
        #文化、体育和娱乐业
        elif row['研究领域'] in ['文化、体育和娱乐业','传媒行业','传媒']:#CIY
            return '文化、体育和娱乐业'
        #农、林、牧、渔业
        elif row['研究领域'] in ['农、林、牧、渔业','农林牧渔']:#CIY
            return '农、林、牧、渔业'
        #社会服务
        elif row['研究领域'] in ['社会服务']:#CIY
            return '社会服务'
        #固定收益
        elif row['研究领域'] in ['固定收益']:#CIY
            return '固定收益'
        #策略
        elif row['研究领域'] in ['策略']:#CIY
            return '策略'
        #宏观
        elif row['研究领域'] in ['宏观','宏观经济']:#CIY
            return '宏观'
        #金融工程
        elif row['研究领域'] in ['3 金融工程','金融工程']:#CIY
            return '金融工程'
        else:
            return None
else:
    def domain(row):
        # 1 宏观
        if row['研究领域'] in ['宏观', '宏观经济']:  # CIY
            return '宏观'
        # 2 策略
        elif row['研究领域'] in ['策略']:  # CIY
            return '策略'
        # 3 金融工程
        elif row['研究领域'] in ['金融工程']:  # CIY
            return '金融工程'
        # 4 固定收益
        elif row['研究领域'] in ['固定收益']:  # CIY
            return '固定收益'
        # 5 农林牧渔
        elif row['研究领域'] in ['农林牧渔']:  # CIY
            return '农林牧渔'
        # 6 基础化工
        elif row['研究领域'] in ['基础化工']:  # CIY
            return '基础化工'
        # 7 钢铁
        elif row['研究领域'] in ['钢铁']:  # CIY
            return '钢铁'
        # 8 有色金属
        elif row['研究领域'] in ['有色金属']:  # CIY
            return '有色金属'
        # 9 电子
        elif row['研究领域'] in ['电子', '801080.SI']:  # CIY
            return '电子'
        # 10 汽车
        elif row['研究领域'] in ['汽车']:  # CIY
            return '汽车'
        # 11 家用电器
        elif row['研究领域'] in ['家用电器', '家电']:  # CIY
            return '家用电器'
        # 12 食品饮料
        elif row['研究领域'] in ['食品饮料', '食品饮料制造加工业']:  # CIY
            return '食品饮料'
        # 13 纺织服饰
        elif row['研究领域'] in ['纺织服饰', '纺织服装']:  # CIY
            return '纺织服饰'
        # 14 轻工制造
        elif row['研究领域'] in ['轻工制造']:  # CIY
            return '轻工制造'
        # 15 医药生物
        elif row['研究领域'] in ['医药生物']:  # CIY
            return '医药生物'
        # 16 公用事业
        elif row['研究领域'] in ['公用事业']:  # CIY
            return '公用事业'
        # 17 交通运输
        elif row['研究领域'] in ['交通运输']:  # CIY
            return '交通运输'
        # 18 房地产
        elif row['研究领域'] in ['房地产']:  # CIY
            return '房地产'
        # 19 商贸零售
        elif row['研究领域'] in ['商贸零售', '批发和零售业']:  # CIY
            return '商贸零售'
        # 20 社会服务
        elif row['研究领域'] in ['社会服务']:  # CIY
            return '社会服务'
        # 21 银行
        elif row['研究领域'] in ['银行']:  # CIY
            return '银行'
        # 22 非银金融
        elif row['研究领域'] in ['非银金融', '非银行金融']:  # CIY
            return '非银金融'
        # 23 综合
        elif row['研究领域'] in ['综合']:  # CIY
            return '综合'
        # 24 建筑材料
        elif row['研究领域'] in ['建筑材料']:  # CIY
            return '建筑材料'
        # 25 建筑装饰
        elif row['研究领域'] in ['建筑装饰']:  # CIY
            return '建筑装饰'
        # 26 电力设备
        elif row['研究领域'] in ['电力设备']:  # CIY
            return '电力设备'
        # 27 机械设备
        elif row['研究领域'] in ['机械设备']:  # CIY
            return '机械设备'
        # 28 国防军工
        elif row['研究领域'] in ['国防军工']:  # CIY
            return '国防军工'
        # 29 计算机
        elif row['研究领域'] in ['计算机', '计算机、通信和其他电子设备制造业']:  # CIY
            return '计算机'
        # 30 传媒
        elif row['研究领域'] in ['传媒']:  # CIY
            return '传媒'
        # 31 通信
        elif row['研究领域'] in ['通信']:  # CIY
            return '通信'
        # 32 煤炭
        elif row['研究领域'] in ['煤炭']:  # CIY
            return '煤炭'
        # 33 石油石化
        elif row['研究领域'] in ['石油石化']:  # CIY
            return '石油石化'
        # 34 环保
        elif row['研究领域'] in ['环保']:  # CIY
            return '环保'
        # 35 美容护理
        elif row['研究领域'] in ['美容护理']:  # CIY
            return '美容护理'
        else:
            return None

df2['研究领域清洗版'] = df2.apply(domain,axis=1)

#%%导入用户报名表、合并公司名称
company = pd.read_excel(f'{processed}{yr}/用户报名表.xlsx')
company.info()
company_name = 'company_name' if yr == 2024 else 'f_company_name'
id = 'ID' if yr == 2024 else 'id'
company.rename(columns={company_name : '公司名称', id : '机构编号'},inplace=True)
df3 = df2.merge(company[['公司名称','机构编号']], on='机构编号', how='left')

print('存在两个领域以上的分析师:')
df6 = df3.drop_duplicates(subset=['分析师姓名','公司名称','研究领域清洗版'])
a = df6[df6.duplicated(subset = ['分析师姓名','公司名称'])][['分析师姓名','机构编号']]
print(a)

b = df6[df6['研究领域清洗版'].isnull()][['分析师姓名','机构编号']]
print(b)
#%%导出分析师报名名单
df4 = df3.drop_duplicates(subset=['分析师姓名','研究领域清洗版','公司名称'],keep='first')[['分析师姓名','研究领域','研究领域清洗版','公司名称','机构编号']]
df4 = df4.sort_values(by=['机构编号','分析师姓名'])
df4['序号'] = range(1,len(df4)+1)
df4 = df4[['序号','分析师姓名','研究领域','研究领域清洗版','公司名称','机构编号']]
df4 = df4.rename(columns={'研究领域':'上报研究领域','研究领域清洗版':'研究领域'})
df4 = df4.drop(columns=['上报研究领域'])

len4 = len(df4)

print(f'参评分析师共{len4}位')

# path5 = r'../5_清洗推荐强度'
# make_dir(path5)

# copy_file(r'../../资料/代码/5.清洗推荐强度.py', path5 + r'/1_清洗推荐强度.py')
# copy_file(r'../../资料/说明/推荐强度说明.docx', path5 + r'/0_推荐强度说明.docx')

a = pd.concat([a,b], ignore_index=True)

with open('2_复数领域分析师.txt','w',encoding='utf-8') as f:
    f.write('填报了复数领域的分析师：\n')
    for i in range(len(a)):
        f.write(f"{a.iloc[i]['机构编号']}的{a.iloc[i]['分析师姓名']}/n")
        print(f"{a.iloc[i]['机构编号']}的{a.iloc[i]['分析师姓名']}")

path_sub = f'{impt}名单/参赛名单.xlsx'
file_path4 = '输出文件/output_分析师已清洗研究领域.xlsx'

check_with_retry(file_path4)
df3.to_excel(file_path4, index=False)

check_with_retry(path_sub)
df4.to_excel(path_sub, index=False)

print('success')

open_file(path_sub)
open_file(file_path4)
copy_file(file_path4, f'{impt}股票信息/4_已清洗研究领域.xlsx',1)
insert_bookmark(__file__, 1, rel, start)