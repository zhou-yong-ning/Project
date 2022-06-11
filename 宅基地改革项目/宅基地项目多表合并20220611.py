import pandas as pd
from win32com import client as wc
import os

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 遍历当前文件夹下的文件
file_list = os.listdir(Current_Folder_path)
# 提取后缀为xls的文件
docx_list = [a for a in file_list if a.endswith('.xls')]
# 将xls转换为xlsx
if not len(docx_list) == 0:
    try:
        word = wc.Dispatch("Excel.Application")
    except:
        word = wc.Dispatch("kwps.application")
    for A in docx_list:
        doc = word.Workbooks.Open(Current_Folder_path + "\\" + A)
        A = A.replace('.docx', '.doc')
        doc.SaveAs(Current_Folder_path + "\\" + A + "x", 51)
        doc.Close()
        print(A + "转换完成")
        os.remove(Current_Folder_path + "\\" + A)
    word.Quit()
# 提取文件内容
df0 = pd.read_excel(Current_Folder_path + "\\" + "宗地属性表.xlsx", dtype='object')
df1 = pd.read_excel(Current_Folder_path + "\\" + "房地一体.xlsx", dtype='object')
# 提取行索引所对应的内容 宗地码    土地坐落  实测面积    东     南     西     北     权利人名称 发证情况
df0 = df0.loc[:, ['ZDTYBM', 'TDZL', 'SCMJ', 'SZD', 'SZN', 'SZX', 'SZB', 'QLRMC', 'FZZT']]
df1 = df1.loc[:, ['宗地代码', '证件号码', '关系']]
# 筛选
df1 = df1.loc[(df1['关系'] == '户主')]
df1 = df1.loc[:, ['宗地代码', '证件号码']]
# 列重命名
df1.rename(columns={'宗地代码': 'ZDTYBM','证件号码':'身份证号码'}, inplace=True)
# 暴力重命名
# df1.columns = ['ZDTYBM','证件号码']
# merge函数匹配
df1to0 = pd.merge(df0, df1, on='ZDTYBM', how='left')
df0 = df1to0
df1 = pd.read_excel(Current_Folder_path + "\\" + "成员信息.xlsx", dtype='object')
mingchen = df1.loc[0,'Unnamed: 1']
daima = df1.loc[0,'Unnamed: 3']
df1 = pd.read_excel(Current_Folder_path + "\\" + "成员信息.xlsx", dtype='object',header=4)
df1 = df1.loc[:,['Unnamed: 6','户主联系电话']]
df1.rename(columns={'Unnamed: 6': '身份证号码','Unnamed: 13':'电话号码'}, inplace=True)
df1to0 = pd.merge(df0, df1, on='身份证号码', how='left')
df0 = df1to0
# df1 = pd.read_excel(Current_Folder_path + "\\" + "成员信息.xlsx", dtype='object')
# 新增一列固定值
df0['证件类型'] = '居民身份证'
df0['宅基地代码'] = daima
# 复制并在首列创建新列
df_1 = df0.TDZL
df0.insert(0,'坐落',df_1)
print(df0)
# df0 = df1to0
# print("表1匹配完成", end='..........\n')
# df2to1to0 = pd.merge(df1to0, df2, on='fwbh', how='left')
# print("表2匹配完成", end='..........\n')
# print(df2to1to0)
# df3to2to1to0 = pd.merge(df2to1to0, df3, on='fwbh', how='left')
# print("表3匹配完成", end='..........\n')
# df4to3to2to1to0 = pd.merge(df3to2to1to0, df4, on='fwbh', how='left')
# print("表4匹配完成", end='..........\n')
# df5to4to3to2to1to0 = pd.merge(df4to3to2to1to0, df5, on='fwbh', how='left')
# print("表5匹配完成", end='..........\n')
# df6to5to4to3to2to1to0 = pd.merge(df5to4to3to2to1to0, df6, on='fwbh', how='left')
# print("表6匹配完成", end='..........\n')
# df = df6to5to4to3to2to1to0
# df = df.loc[:, ['DKBM', 'QLRMC']]
# print("表数据提取完成", end='..........\n')
# df6to5to4to3to2to1to0.to_excel(Current_Folder_path + '\\汇总.xlsx', index=False)
# print("数据写入完成", end='..........\n')
# sleep(1)

# if 'a' in df0.columns:
#     print('zai')
# else:
#     print('fou')
# df0.to_excel(Current_Folder_path + '\\新建文件夹\\汇总表.xlsx', index=False)
# df = df2to1to0
# df = df0.loc[:, ['fwbh', '幢号', '姓名', '身份证号', '省（市、区）', '市（州、盟）', '县（市、区、旗）', '乡（镇、街道）', '村（社区）', '组']]
# df.to_excel(Current_Folder_path + '\\汇总.xlsx', index=False)
# print("数据写入完成", end='..........\n')
