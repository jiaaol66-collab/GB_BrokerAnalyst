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

now = datetime.now()
nowdate = now.strftime('%Y%m%d')[2:]
print('---------------输入与设置{}---------------'.format(datetime.now()))
datelist = []
data = pd.ExcelFile(impt + '提取模板/分析师定量数据.xlsx')
for dates in data.sheet_names:
    if dates != '模板':
        datelist.append(dates)

# date = input(f'提取分析师定量数据和同行业回报率的日期(格式MMDD)(一共有这些日期)\n{datelist}:')
# if not date:
#     date = datelist[-1]
date = max(datelist)

extractdate = '_' + date
appe = extractdate

switch_method = ''
# switch_method = input('对于同一分析师的多个股票使用算数平均分排名还是最高分排名？(enter/else):')
if not switch_method:
    method = '算术平均'
else:
    method = '最大值'

# isobject = input('是否先进行客观评价排名?(1/else):')
isobject = '主观'
if not isobject: object = '_客观评价'
else:  
    if os.path.exists('../重要文件/主观评分/分析师记分统计.xlsx'):
        object = ''

    else:
        print('没有主观评分相关的资料, 目前进行客观评价排名。')
        object = '_客观评价'
        isobject = ''

make_dir('1_计算过程')
make_dir('2_计算结果')

start = time.time()
import pandas as pd
df = pd.read_excel(data, sheet_name=date)

df[['三个月回报率','六个月回报率','7.1回报率']] /= 100
# df_pct = pd.read_excel(f"../重要文件/同行业回报率/已提取数据{extractdate}.xlsx")
df_res = pd.read_excel("../重要文件/分析师清洗版/买方反应_1028.xlsx")

if isobject:
    domain_4 = pd.read_excel(r'../重要文件/名单/四大领域分析师.xlsx')
    result_subjective = pd.read_excel(r'../重要文件/主观评分/分析师记分统计.xlsx')
    no_stock = pd.read_excel(r'../重要文件/名单/取消资格.xlsx')
    if '主观评价分数' in df.columns.to_list():
        df = df.drop(columns=['主观评价分数'])

path_br = rf'1_计算过程/买方反应{nowdate}{extractdate}.xlsx'
path_er = rf'1_计算过程/超额回报率{nowdate}{extractdate}.xlsx'
path_sum = rf'1_计算过程/最终分数汇总{nowdate}{extractdate}{object}.xlsx'
file_path_all = f'2_计算结果/计算排名结果{nowdate}{method}{extractdate}{object}.xlsx'

#%%    1、买方反应
df = df.merge(df_res.drop(columns=['距离7月1日天数','看涨/看跌']), on=['序号'], how='left')
df[['机构比例变化','股东数量变化率']] =df[['机构比例变化','股东数量变化率']]/100
def buyer_response(group):
    #决定使用排名赋分
    """mean_inst = group['机构比例变化'].mean()  
    std_inst = group['机构比例变化'].std()  
    group['机构比例变化_Z'] = ((group['机构比例变化'] - mean_inst) / std_inst)  
    group['机构比例变化_100'] = (group['机构比例变化_Z'] - group['机构比例变化_Z'].min()) / (group['机构比例变化_Z'].max() - group['机构比例变化_Z'].min()) * 100  #"""
    group['机构比例变化_Z'] = 0
    group['机构比例排名'] = 0
    group['机构比例变化_100'] = 0
    score = group[(~group['机构比例变化'].isnull())]
    score['机构比例变化_Z'] = (score['机构比例变化']-score['机构比例变化'].mean())/score['机构比例变化'].std()
    score['机构比例排名'] = score['机构比例变化_Z'].rank(method='dense', ascending=False)
    n = score['机构比例排名'].max()
    score['机构比例变化_100'] = (100 - (score['机构比例排名']-1)*50/(n-1)).fillna(0)
    group.loc[score.index, '机构比例变化_Z'] = score['机构比例变化_Z']
    group.loc[score.index, '机构比例变化_100'] = score['机构比例变化_100']
    #group.loc[group['机构比例变化']==0,'机构比例变化_100'] = 50
    group.loc[score.index, '机构比例排名'] = score['机构比例排名']
        
    """max_share = group['股东数量变化率'].max()  
    min_share = group['股东数量变化率'].min()  
    group['股东数量变化率_100'] = ((group['股东数量变化率'] - min_share) / (max_share - min_share) * 100).fillna(0)#"""
    group['股东数量变化率_Z'] = 0
    group['股东数量变化率排名'] = 0
    group['股东数量变化率_100'] = 0
    score = group[(~group['股东数量变化率'].isnull())]
    score['股东数量变化率_Z'] = (score['股东数量变化率']-score['股东数量变化率'].mean())/score['股东数量变化率'].std()
    score['股东数量变化率排名'] = score['股东数量变化率_Z'].rank(method='dense',ascending=False)
    n = score['股东数量变化率排名'].max()
    score['股东数量变化率_100'] = (100 - (score['股东数量变化率排名']-1)*50/(n-1)).fillna(0)
    group.loc[score.index, '股东数量变化率_Z'] = score['股东数量变化率_Z']
    group.loc[score.index, '股东数量变化率_100'] = score['股东数量变化率_100']
    #group.loc[group['股东数量变化率']==0,'股东数量变化率_100'] = 50
    group.loc[score.index, '股东数量变化率排名'] = score['股东数量变化率排名']
    
    """mean_fund = group['头部公募基金变化'].mean()  
    std_fund = group['头部公募基金变化'].std()  
    group['头部公募基金变化_Z'] = ((group['头部公募基金变化'] - mean_fund) / std_fund)  
    group['头部公募基金变化_100'] = ((group['头部公募基金变化_Z'] - group['头部公募基金变化_Z'].min()) / (group['头部公募基金变化_Z'].max() - group['头部公募基金变化_Z'].min()) * 100).fillna(0)#"""   
    group['头部公募基金变化_Z'] = 0
    group['头部公募基金排名'] = 0
    group['头部公募基金变化_100'] = 0    
    score = group[(~group['头部公募基金变化'].isnull())]
    score['头部公募基金变化_Z'] = (score['头部公募基金变化']-score['头部公募基金变化'].mean())/score['头部公募基金变化'].std()
    score['头部公募基金排名'] = score['头部公募基金变化_Z'].rank(method='dense',ascending=False)
    n = score['头部公募基金排名'].max()
    score['头部公募基金变化_100'] = (100 - (score['头部公募基金排名']-1)*50/(n-1)).fillna(0)
    group.loc[score.index, '头部公募基金变化_Z'] = score['头部公募基金变化_Z']
    group.loc[score.index, '头部公募基金变化_100'] = score['头部公募基金变化_100']
    #group.loc[group['头部公募基金变化']==0,'头部公募基金变化_100'] = 50
    group.loc[score.index, '头部公募基金排名'] = score['头部公募基金排名']
    
    group['买方反应'] = 0.15*group['机构比例变化_100'] + 0.1*group['股东数量变化率_100'] + 0.05 * group['头部公募基金变化_100']
    return group  
  
df_br = df[['序号', '研究领域清洗版', '看涨/看跌', '权重', '是否相符_ins', '机构比例变化', '是否相符_shareholder', '股东数量变化率', '是否相符_public', '头部公募基金变化']].copy()

result_response = df_br.groupby('研究领域清洗版').apply(buyer_response).reset_index(drop=True)  

result_response = result_response[['序号','研究领域清洗版','看涨/看跌','权重','是否相符_ins',
                                   '机构比例变化','机构比例变化_Z','机构比例排名','机构比例变化_100',
                                   '是否相符_shareholder','股东数量变化率','股东数量变化率_Z',
                                   '股东数量变化率排名','股东数量变化率_100','是否相符_public',
                                   '头部公募基金变化','头部公募基金变化_Z','头部公募基金排名',
                                   '头部公募基金变化_100','买方反应']]

print(result_response.head())
result_response.to_excel(path_br, index=False)
#%%    2、超额回报率的计算
import pandas as pd
#提取同行业数据
# try:
#     weight = df[['序号','权重','推荐/目标价格','6个月内最高价']].copy()
#     df_pct = df_pct.merge(weight,on='序号',how='left')
# except Exception as e:
#     print('error', e)
#     weight = df[['序号','权重']]
#     df_pct = df_pct.merge(weight,on='序号',how='left')
    
# df_pct = df_pct[df_pct['序号'].isin(df['序号'])]

# #采用清洗版的回报率
# df_pct = df_pct.merge(df[['序号','三个月回报率', '六个月回报率', '7.1回报率']],on='序号',how='left')
#
# df_pct.info()
# #年化
# df_pct['三个月回报率_年化'] = df_pct['三个月回报率']*4
# df_pct['六个月回报率_年化'] = df_pct['六个月回报率']*2
# df_pct['7.1回报率_年化'] = df_pct['7.1回报率']*365/df_pct['距离7月1日天数']
# df_pct.head()
#
# df_pct.head()
#
# df_pct[['三个月回报率', '六个月回报率', '7.1回报率', '三个月回报率_年化', '六个月回报率_年化', '7.1回报率_年化']] = df_pct[['三个月回报率', '六个月回报率', '7.1回报率', '三个月回报率_年化', '六个月回报率_年化', '7.1回报率_年化']]/100#"""
df_pct = df.copy()

# def excess_return(group):
#     #对Z_score加权后排名赋分，最高100分，最低0分
#     group['超额回报率三个月年化'] = group['三个月回报率_年化'] - group['同行业三个月年化回报']
#     mean3 = group['超额回报率三个月年化'].mean()
#     std3 = group['超额回报率三个月年化'].std()
#     group['超额回报率三个月年化Z'] = ((group['超额回报率三个月年化']-mean3)/std3)*group['权重']
#     group['超额年化回报率三个月排名'] = group['超额回报率三个月年化Z'].rank(method='dense',ascending=False)
#     n = group['超额年化回报率三个月排名'].max()
#     group['超额年化回报率三个月得分'] = (100 - (group['超额年化回报率三个月排名']-1)*100/n).fillna(100)
#
#     group['超额回报率六个月年化'] = group['六个月回报率_年化'] - group['同行业六个月年化回报']
#     mean6 = group['超额回报率六个月年化'].mean()
#     std6 = group['超额回报率六个月年化'].std()
#     group['超额回报率六个月年化Z'] = ((group['超额回报率六个月年化']-mean6)/std6)*group['权重']
#     group['超额年化回报率六个月排名'] = group['超额回报率六个月年化Z'].rank(method='dense',ascending=False)
#     n = group['超额年化回报率六个月排名'].max()
#     group['超额年化回报率六个月得分'] = (100 - (group['超额年化回报率六个月排名']-1)*100/n).fillna(100)
#
#     group['超额回报率7.1年化'] = group['7.1回报率_年化'] - group['同行业7.1年化回报']
#     mean630 = group['超额回报率7.1年化'].mean()
#     std630 = group['超额回报率7.1年化'].std()
#     group['超额回报率7.1年化Z'] = ((group['超额回报率7.1年化']-mean630)/std630)*group['权重']
#     group['超额年化回报率7.1排名'] = group['超额回报率7.1年化Z'].rank(method='dense',ascending=False)
#     n = group['超额年化回报率7.1排名'].max()
#     group['超额年化回报率7.1得分'] = (100 - (group['超额年化回报率7.1排名']-1)*100/n).fillna(100)
#     return group


def excess_return(group):
    # 对Z_score加权后排名赋分，最高100分，最低0分
    # group['超额回报率三个月年化'] = group['三个月回报率_年化'] - group['同行业三个月年化回报']
    mean3 = group['超额回报率三个月'].mean()
    std3 = group['超额回报率三个月'].std()
    group['超额回报率三个月年化Z'] = ((group['超额回报率三个月'] - mean3) / std3) * group['权重']
    group['超额年化回报率三个月排名'] = group['超额回报率三个月年化Z'].rank(method='dense', ascending=False)
    n = group['超额年化回报率三个月排名'].max()
    group['超额年化回报率三个月得分'] = (100 - (group['超额年化回报率三个月排名'] - 1) * 100 / n).fillna(100)

    # group['超额回报率六个月年化'] = group['六个月回报率_年化'] - group['同行业六个月年化回报']
    mean6 = group['超额回报率六个月'].mean()
    std6 = group['超额回报率六个月'].std()
    group['超额回报率六个月年化Z'] = ((group['超额回报率六个月'] - mean6) / std6) * group['权重']
    group['超额年化回报率六个月排名'] = group['超额回报率六个月年化Z'].rank(method='dense', ascending=False)
    n = group['超额年化回报率六个月排名'].max()
    group['超额年化回报率六个月得分'] = (100 - (group['超额年化回报率六个月排名'] - 1) * 100 / n).fillna(100)

    # group['超额回报率7.1年化'] = group['7.1回报率_年化'] - group['同行业7.1年化回报']
    mean630 = group['超额回报率7.1'].mean()
    std630 = group['超额回报率7.1'].std()
    group['超额回报率7.1年化Z'] = ((group['超额回报率7.1'] - mean630) / std630) * group['权重']
    group['超额年化回报率7.1排名'] = group['超额回报率7.1年化Z'].rank(method='dense', ascending=False)
    n = group['超额年化回报率7.1排名'].max()
    group['超额年化回报率7.1得分'] = (100 - (group['超额年化回报率7.1排名'] - 1) * 100 / n).fillna(100)
    return group


#对于有目标价格的，6个月内最高价超过目标价格的给10分，未达到给0分
df_pct['是否达到目标价格'] = 0 
df_pct.loc[df_pct['推荐/目标价格'].isnull(),'是否达到目标价格'] = -1
df_pct.loc[(df_pct['推荐/目标价格']<=df_pct['6个月内最高价']),'是否达到目标价格'] = 10
"""df_pct['分数核算方式'] = 3
df_pct.loc[df_pct['距离7月1日天数']<182,'分数核算方式']=2
df_pct.loc[df_pct['距离7月1日天数']<90,'分数核算方式']=3#"""

df_pct['分组方式'] = 2
df_pct.loc[df_pct['距离7月1日天数']<182,'分组方式'] = 1#"""

if appe in ['-1']:#现在想来还是三者分别算，然后再分别按排名得分会比较公平，即不用分第二次组
    result_return = df_pct.groupby(['研究领域清洗版','分组方式']).apply(excess_return).reset_index(drop=True)
else:
    result_return = df_pct.groupby('研究领域清洗版').apply(excess_return).reset_index(drop=True)

#对于少于7.1的
result_return['超额回报率分数1'] = result_return['超额年化回报率三个月得分']*0.25 # 90~182天
result_return.loc[result_return['距离7月1日天数']>=182,'超额回报率分数1'] = result_return['超额年化回报率三个月得分']*0.12 + result_return['超额年化回报率六个月得分']*0.18 # 182天以上
result_return.loc[result_return['距离7月1日天数']<90,'超额回报率分数1'] = result_return['超额年化回报率7.1得分']*0.2 # 90天以下
result_return['超额回报率分数'] = result_return['超额回报率分数1']*4/3
result_return.loc[result_return['是否达到目标价格']!=-1,'超额回报率分数'] = result_return['是否达到目标价格'] + result_return['超额回报率分数1']
result_return['超额回报率分数'] = result_return['超额回报率分数'] + result_return['扣分']

print(result_return.head())

result_return = result_return[['序号', '推荐日期清洗版', '研究领域清洗版', '距离7月1日天数', '中期观察日', '长期观察日',
        '权重', '三个月回报率', '六个月回报率','7.1回报率','三个月回报率_年化', '同行业三个月年化回报','超额回报率三个月', '超额回报率三个月年化Z', '超额年化回报率三个月排名','超额年化回报率三个月得分',
       '六个月回报率_年化', '同行业六个月年化回报','超额回报率六个月', '超额回报率六个月年化Z', '超额年化回报率六个月排名', '超额年化回报率六个月得分',
       '7.1回报率_年化','同行业7.1年化回报','超额回报率7.1', '超额回报率7.1年化Z', '超额年化回报率7.1排名', '超额年化回报率7.1得分', '分组方式',
       '推荐/目标价格','6个月内最高价', '是否达到目标价格', '扣分', '超额回报率分数1', '超额回报率分数']]

result_return.to_excel(path_er,index=False)
#%%    3、合并主观评价分数（如果有）【if strat != '1':】
if isobject:
    result_subjective = result_subjective.rename(columns={'研究领域':'研究领域清洗版'})
    result_subjective.info()
    matching_subject = ['分析师姓名', '公司名称', '研究领域清洗版', '主观评价分数']
    match = ['分析师姓名', '公司名称', '研究领域清洗版']
    domain_4 = domain_4.merge(result_subjective[matching_subject], on=match, how='left')
    domain_4.rename(columns = {'研究领域清洗版':'研究领域'})
    domain_4['主观评价分数'].fillna(0, inplace=True)
    df = df.merge(result_subjective[matching_subject], on=match, how='left')
    no_stock = no_stock.merge(result_subjective[matching_subject], on=match, how='left')
    no_stock = no_stock[~no_stock['唯一标识'].isin(df['唯一标识'])]
    df = pd.concat([df,no_stock])
    df['主观评价分数'].fillna(0, inplace=True)
#%%    4、排名算分
#先将模板提出
df2 = df[['序号','分析师姓名','公司名称','主观评价分数','唯一标识','研究领域清洗版','标的股票代码','标的股票名称','推荐日期清洗版','推荐强度清洗版','看涨/看跌','权重','推荐/目标价格','6个月内最高价','推荐时收盘价','研报发布时标的价格','扣分','机构比例变化','股东数量变化率','头部公募基金变化']].copy()
print('df2')
print(df2.head())
time.sleep(4)
df_concat = []
df_group = df2.groupby('唯一标识')
for name,group in df_group:
    group['推荐股票总数'] = group['标的股票代码'].nunique()
    df_concat.append(group)
df2 = pd.concat(df_concat)
if isobject:
   df2 = pd.concat([df2,domain_4])

# ##修改--取消资格的名单
# check = ['刘乐文','竺劲']
# df2 = df2[~df2['分析师姓名'].isin(check)]

#买方反应
reaction = result_response[['序号','机构比例排名','机构比例变化_100','股东数量变化率排名','股东数量变化率_100','头部公募基金排名','头部公募基金变化_100','买方反应']]

#超额回报率
return0 = result_return[['序号','三个月回报率','三个月回报率_年化', '同行业三个月年化回报','超额回报率三个月', '超额回报率三个月年化Z', '超额年化回报率三个月排名', '超额年化回报率三个月得分',
                   '六个月回报率','六个月回报率_年化', '同行业六个月年化回报', '超额回报率六个月', '超额回报率六个月年化Z','超额年化回报率六个月排名','超额年化回报率六个月得分',
                   '7.1回报率','7.1回报率_年化', '同行业7.1年化回报','超额回报率7.1', '超额回报率7.1年化Z', '超额年化回报率7.1排名','超额年化回报率7.1得分','是否达到目标价格',
                   '超额回报率分数1','超额回报率分数','分组方式']] 

#合并数据
result = df2.merge(reaction,on='序号',how='left').merge(return0,on='序号',how='left')

#更名
rename ={'股东数量变化率':'股东数量变化率','股东数量变化率_100':'股东数量变化率_100','股东数量变化率排名':'股东数量变化率排名','超额回报率分数1':'超额回报率(未合并推荐分数)',
        '超额回报率分数':'超额回报率','研究领域清洗版':'研究领域','推荐日期清洗版':'推荐日期','推荐强度清洗版':'推荐强度','主观评价分数':'主观评价分数(原始)'}
result.rename(columns=rename,inplace=True)

result.info()

#计算客观评价分数（两个分数直接相加，因为已经加过权了）
result['客观评价分数']=(result['买方反应']+result['超额回报率']).fillna(0)

print(result['超额回报率'].describe(),result['买方反应'].describe())

result=result[['序号', '分析师姓名', '公司名称', '唯一标识', '研究领域', '标的股票代码', '标的股票名称','推荐日期', '推荐强度',
               '看涨/看跌', '权重','机构比例变化','机构比例排名','机构比例变化_100', '股东数量变化率','股东数量变化率排名','股东数量变化率_100','头部公募基金变化', '头部公募基金排名','头部公募基金变化_100', '买方反应',
               '三个月回报率', '三个月回报率_年化', '同行业三个月年化回报','超额回报率三个月', '超额回报率三个月年化Z', '超额年化回报率三个月排名', '超额年化回报率三个月得分',
               '六个月回报率', '六个月回报率_年化', '同行业六个月年化回报', '超额回报率六个月', '超额回报率六个月年化Z','超额年化回报率六个月排名','超额年化回报率六个月得分',
               '7.1回报率', '7.1回报率_年化', '同行业7.1年化回报','超额回报率7.1', '超额回报率7.1年化Z', '超额年化回报率7.1排名','超额年化回报率7.1得分','分组方式','超额回报率(未合并推荐分数)','推荐/目标价格', '6个月内最高价', '是否达到目标价格','推荐时收盘价','研报发布时标的价格','扣分',
               '超额回报率', '客观评价分数','主观评价分数(原始)','推荐股票总数']]


#主观评价分数加权，计算分析师平均分与各项评分
result['客观评价分数(平均)'] = result.groupby('唯一标识')['客观评价分数'].transform('mean')
result['主观评价分数(原始)'] = result['主观评价分数(原始)'].fillna(0)
result['主观评价分数'] = result['主观评价分数(原始)']*0.3
result['得分'] = result['客观评价分数'] + result['主观评价分数']
result['平均得分'] = result['客观评价分数(平均)'] + result['主观评价分数']
result['推荐日期'] = result['推荐日期'].dt.strftime('%Y-%m-%d')
result.loc[result['主观评价分数']==0,['主观评价分数', '主观评价分数(原始)']] = np.nan

print(result.head())

#先按照分析师的名字和机构组成的唯一标识与客观评价分数排列。
result = result.sort_values(by=['唯一标识','客观评价分数'],ascending=False)

result2 = result[['分析师姓名', '公司名称', '得分', '平均得分', '研究领域', '标的股票代码', '标的股票名称', '推荐日期',
       '推荐强度', '看涨/看跌', '权重', '机构比例变化', '股东数量变化率', '头部公募基金变化', '机构比例变化_100',
       '股东数量变化率_100', '头部公募基金变化_100', '买方反应', '三个月回报率', '六个月回报率',
       '超额年化回报率三个月得分', '超额年化回报率六个月得分', '推荐/目标价格', '6个月内最高价',
       '是否达到目标价格', '扣分', '超额回报率', '客观评价分数', '客观评价分数(平均)',
       '主观评价分数','推荐股票总数']]
    

#按照研究领域>平均得分>得分的顺序排列
if not isobject:
    result2 = result2.sort_values(by=['研究领域','客观评价分数(平均)','客观评价分数'],ascending=False)
else:
    result2 = result2.sort_values(by=['研究领域','平均得分','得分'],ascending=False)

#直接复制用于提取同行业回报率的字典代码

if yr != 2024:
    ind = [[] for _ in range(35)]
    industryID = {
        '宏观': ind[0], '策略': ind[1], '金融工程': ind[2], '固定收益': ind[3], '农林牧渔': ind[4],
        '基础化工': ind[5], '钢铁': ind[6], '有色金属': ind[7], '电子': ind[8], '汽车': ind[9],
        '家用电器': ind[10], '食品饮料': ind[11], '纺织服饰': ind[12], '轻工制造': ind[13],
        '医药生物': ind[14], '公用事业': ind[15], '交通运输': ind[16], '房地产': ind[17],
        '商贸零售': ind[18], '社会服务': ind[19], '银行': ind[20], '非银金融': ind[21],
        '综合': ind[22], '建筑材料': ind[23], '建筑装饰': ind[24], '电力设备': ind[25],
        '机械设备': ind[26], '国防军工': ind[27], '计算机': ind[28], '传媒': ind[29],
        '通信': ind[30], '煤炭': ind[31], '石油石化': ind[32], '环保': ind[33], '美容护理': ind[34]
    }

else:
    ind = [[] for _ in range(32)]
    


with pd.ExcelWriter(path_sum, engine='openpyxl') as writer:
    for ind in industryID.keys():
        result_ind = result2[result2['研究领域']==ind]
        result_ind = result_ind.sort_values(by=['平均得分'],ascending=False)
        result_ind.index = result_ind['平均得分'].rank(method='dense',ascending=False)
        result_ind.index.name='排名'
        result_ind.to_excel(writer,sheet_name=f'{ind}')

match = ['分析师姓名', '公司名称', '研究领域']
#暂且保留每位分析师的最高分股票到最终呈现的排名
result2 = result2.drop_duplicates(subset=match)#默认保留最上面那条，也就是分最高的那只股票

#稍微更改一下顺序
result3 = result2[['分析师姓名', '公司名称', '得分', '平均得分', '研究领域', '标的股票代码', '标的股票名称', '推荐日期',
       '推荐强度', '看涨/看跌', '权重', '机构比例变化', '股东数量变化率', '头部公募基金变化', '机构比例变化_100',
       '股东数量变化率_100', '头部公募基金变化_100', '买方反应', '三个月回报率', '六个月回报率',
       '超额年化回报率三个月得分', '超额年化回报率六个月得分', '推荐/目标价格', '6个月内最高价',
       '是否达到目标价格', '扣分', '超额回报率', '客观评价分数', '客观评价分数(平均)',
       '主观评价分数','推荐股票总数']]


check_with_retry(file_path_all)
with pd.ExcelWriter(file_path_all, engine='openpyxl') as writer:
    for ind in industryID.keys():
        result_ind = result3[result3['研究领域']==ind]
        if not isobject:
            if ind in ['固定收益','金融工程','宏观','策略']:
                result_ind = result_ind.sort_values(by='得分',ascending=False)
                result_ind = result_ind[['分析师姓名','公司名称','得分','研究领域']]
                result_ind['得分'] = result_ind['得分']*10/3
                result_ind.index = result_ind['得分'].rank(method='dense',ascending=False)
            else:
                if not switch_method: # 算术平均
                    result_ind = result_ind.sort_values(by=['客观评价分数(平均)'],ascending=False)
                    result_ind.index = result_ind['客观评价分数(平均)'].rank(method='dense',ascending=False)
                else:
                    result_ind = result_ind.sort_values(by=['客观评价分数'],ascending=False)
                    result_ind.index = result_ind['客观评价分数'].rank(method='dense',ascending=False)
        else:
            if ind in ['固定收益','金融工程','宏观','策略']:
                result_ind = result_ind.sort_values(by='得分',ascending=False)
                result_ind = result_ind[['分析师姓名','公司名称','得分','研究领域']]
                result_ind['得分'] = result_ind['得分']*10/3
                result_ind.index = result_ind['得分'].rank(method='dense',ascending=False)
            else:
                if not switch_method:
                    result_ind = result_ind.sort_values(by=['平均得分'],ascending=False)
                    result_ind.index = result_ind['平均得分'].rank(method='dense',ascending=False)
                else:
                    result_ind = result_ind.sort_values(by=['得分'],ascending=False)
                    result_ind.index = result_ind['得分'].rank(method='dense',ascending=False)
        result_ind.index.name= '排名'
        result_ind.to_excel(writer,sheet_name=f'{ind}')


"""import sys
if not isobject:#暂时不进行后续步骤
    end = time.time()
    print(f'花了{round(end-start,3)}秒')
    sys.exit(0)"""
#%%    5、更改文件呈现方式
from openpyxl import load_workbook  
from openpyxl.styles import PatternFill  

def set_column_number_format(sheet, column_letter, number_format='0.00'):  
    """设置指定列的数值格式"""  
    for cell in sheet[column_letter]:  
        if cell.value is not None:  
            cell.number_format = number_format  
            
def set_column_pct_format(sheet, column_letter, number_format='0.00%'):  
    """设置指定列的数值格式"""  
    for cell in sheet[column_letter]:  
        if cell.value is not None:  
            cell.number_format = number_format 

def estimate_column_width(sheet, column_letter, default_width=10):  
    """基于内容长度估算列宽（非常近似的方法）"""  
    max_length = 0  
    for cell in sheet[column_letter]:  
        if cell.value is not None:  
            cell_content = str(cell.value)  
            cell_length = len(cell_content)  
            if cell_length > max_length:  
                max_length = cell_length  
    # 返回一个稍微大于最长内容长度的宽度（加上额外的缓冲）  
    return max(default_width, (max_length + 2))  
  
def highlight_and_adjust_columns(file_path, highlight_columns, adjust_only_columns):  
    """  
    对多个特定列进行高亮和列宽调整。  
    highlight_columns: 需要高亮的列字母列表。  
    adjust_only_columns: 仅需要调整列宽的列字母列表，这些列不会被高亮。  
    """  
    wb = load_workbook(file_path)  
    hl_columns = highlight_columns
    ajo_columns = adjust_only_columns
    for sheet_name in wb.sheetnames:  
        sheet = wb[sheet_name]  
        if sheet_name not in ['固定收益','金融工程','宏观','策略']:
            adjust_only_columns = ajo_columns
            highlight_columns = hl_columns
            # 冻结窗格位置
            sheet.freeze_panes = 'F7'
        if sheet_name in ['固定收益','金融工程','宏观','策略']:
            highlight_columns = ['D']
            # 冻结窗格位置
            sheet.freeze_panes = 'E7'
        for column_letter in highlight_columns:  
            cell = sheet[f"{column_letter}1"]  
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  
            cell.fill = fill  
  
            set_column_number_format(sheet, column_letter)  
  
            estimated_width = estimate_column_width(sheet, column_letter)  
            sheet.column_dimensions[column_letter].width = estimated_width  
  
        for column_letter in adjust_only_columns:  
            estimated_width = estimate_column_width(sheet, column_letter)  
            sheet.column_dimensions[column_letter].width = estimated_width
            set_column_pct_format(sheet, column_letter)#设置百分比格式

  
    wb.save(file_path)  
    
file_paths = [file_path_all]
             
highlight_columns = ['S','E','D','AB','AC','AD','AE']  # 需要高亮的列
adjust_columns = ['M','N','T','U']  # 仅需要调整列宽，不需要高亮的列
style = {}

for element in ['A','B','D','E','F','G','H','S','AA','AB','AC','AD','AE']:
    style[element] = {'center':True}

for file_path in file_paths:
    highlight_and_adjust_columns(file_path, highlight_columns, adjust_columns)
    auto_fit_columns(file_path, file_path)
    format_excel_columns(file_path,
            {('A','B','D','E','F','G','H','S','AA','AB','AC','AD','AE'):{'center':True}})

end = time.time()
print(f'总共花费{round(end-start,3)}秒')

open_file(file_path_all)
