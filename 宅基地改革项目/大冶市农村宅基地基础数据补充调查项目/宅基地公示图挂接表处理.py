import pandas as pd
import os
import warnings
from win32com import client as wc

# 忽略警告
warnings.filterwarnings('ignore')
# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 遍历当前文件夹下的文件
file_list = os.listdir(Current_Folder_path)
# 创建子文件夹
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
        os.mkdir(os.path.join(Current_Folder_path, 'NewFolder'))
save_path = os.path.join(Current_Folder_path, 'NewFolder')
# 提取后缀为xlsx的文件
xlsx_list = [a for a in file_list if a.endswith('.xlsx')]
# 利用推导式获取指定表数据
syqr = [i for i in xlsx_list if '使用权人表数据' in i]
fwxx = [i for i in xlsx_list if '房屋信息表数据' in i]
zdxx = [i for i in xlsx_list if '宗地信息表数据' in i]
dfsyqr = pd.read_excel(os.path.join(Current_Folder_path, syqr[0]), dtype='object')
print('使用权人表读取成功')
dfzdxx = pd.read_excel(os.path.join(Current_Folder_path, zdxx[0]), dtype='object')
print('宗地信息表读取成功')
dffwxx = pd.read_excel(os.path.join(Current_Folder_path, fwxx[0]), dtype='object')
print('房屋信息表读取成功')
# 使用权人表提取关键信息
dfsyqr = dfsyqr.loc[:,['宅基地代码','使用权人代表姓名']]
dfsyqr.rename(columns={'使用权人代表姓名':'NAME'},inplace=True)
dfsyqr.to_excel(os.path.join(Current_Folder_path,'NewFolder','使用权人表数据.xlsx'),sheet_name='Sheet2',index=False)
# 宗地信息表提取关键信息
dfzdxx = dfzdxx.loc[:,['宅基地代码','宗地面积','宅基地利用状况']]
dfsyqr.rename(columns={'宗地面积':'ZDMJ','宅基地利用状况':'ZJDLYZK'},inplace=True)
dfzdxx.to_excel(os.path.join(Current_Folder_path,'NewFolder','宗地信息表数据.xlsx'),sheet_name='Sheet2',index=False)
# 房屋信息表提取关键信息
dffwxx = dffwxx.loc[:,['农民房屋代码','利用状况']]
dffwxx.rename(columns={'利用状况':'FWLYZK'},inplace=True)
dffwxx.to_excel(os.path.join(Current_Folder_path,'NewFolder','房屋信息表数据.xlsx'),sheet_name='Sheet2',index=False)
# 文件类型转换
path= os.path.join(Current_Folder_path,'NewFolder')
file_list = os.listdir(path)
# 提取后缀为xlsx的文件
xlsx_list = [a for a in file_list if a.endswith('.xlsx')]
# 将xlsx转换为xls
if not len(xlsx_list) == 0:
    try:
        # 打开office应用程序
        file = wc.Dispatch("Excel.Application")
    except:
        # 打开wps应用程序
        file = wc.Dispatch("kwps.application")
    for A in xlsx_list:
        # 打开文件
        doc = file.Workbooks.Open(os.path.join(path, A))
        # 替换文件名
        B = A.replace('.xlsx', '.xls')
        # 文件另存
        doc.SaveAs(os.path.join(path, B), 56)
        # 关闭当前文档
        doc.Close()
        print(A + "转换完成")
        # 删除原数据
        os.remove(os.path.join(path,A))
    # 关闭应用程序
    file.Quit()
