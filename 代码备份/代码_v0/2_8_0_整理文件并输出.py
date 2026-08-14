import os,shutil,pandas as pd

#%%去重、更改顺序，输出清洗完毕后文件，后可制作成模板文件
df3 = pd.read_excel(r'..\7.清洗股票代码\output_分析师已初步清洗股票代码.xlsx')
df3 = df3.sort_values(by=['机构编号','分析师姓名'])
df3['序号'] = None
df3['序号'].iloc[0] = 1
df3 = df3[['序号','分析师姓名','唯一标识','公司名称','研究领域','研究领域清洗版','标的股票代码(原)','标的股票代码(清洗)','标的股票代码','标的股票名称(原)','标的股票名称','标的股票名称(对比)','推荐日期','推荐日期清洗版','推荐强度','推荐强度清洗版','看涨/看跌','权重','机构编号','研报发布时标的价格(原)','研报推荐的目标价格(原)','研报推荐的目标价格','研报推荐的目标价格下限','上限','研报发布时标的价格']]
df3.info()
print(len(df3[df3.duplicated(keep=False)]),df3[df3.duplicated(keep=False)])
df3 = df3.sort_values(by=['机构编号','分析师姓名'])
df3.drop_duplicates(inplace=True)#对应情况1,共8条
df3.info()
df5 = df3[df3['研究领域清洗版'].isin(['金融工程','固定收益','宏观','策略'])]
df7 = df3[~df3['研究领域清洗版'].isin(['金融工程','固定收益','宏观','策略'])]#对应情况2，共105条
df5 = df5[['分析师姓名','研究领域清洗版','公司名称']]
df5.drop_duplicates(subset=['分析师姓名','公司名称'],inplace=True)

def toexcel(df,file_path):
    df.to_excel(file_path, index=False)
    os.startfile(file_path)

def make_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def copy_file(src, path):
    if not os.path.exists(path):
        shutil.copyfile(src, path)

path_sub = '..\..\必要文件\名单'
df5.to_excel(path_sub +r'\四大领域分析师.xlsx',index=False)

path3 = r'..\..\step.3 提取万得数据\1.数据准备\1.股票代码清洗与数据提取'
make_dir(path3)

df6 = df7[df7['标的股票代码(原)'] == '-']
df7 = df7[df7['标的股票代码(原)'] != '-']#对应情况3，共34条
df7.info()

copy_file(r'../说明/模板复制.xlsx', '模板复制.xlsx')
os.startfile('模板复制.xlsx')
copy_file(r'..\..\资料\代码\8.准备数据.py','3_准备数据.py')
copy_file(r'..\4.清洗研究领域\2.复数领域分析师.txt','复数领域分析师.txt')

file_path = '3_股票信息部分清洗版.xlsx'

print('success')
import time
time.sleep(1.2)

toexcel(df7, file_path)
toexcel(df6, path3 + r'\取消资格.xlsx')