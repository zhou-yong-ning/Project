import pandas as pd
import os
# 自动调整列宽
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
def reset_col(filename):
    wb = load_workbook(filename)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        df = pd.read_excel(filename, sheet).fillna('-')
        df.loc[len(df)] = list(df.columns)  # 把标题行附加到最后一行
        for col in df.columns:
            index = list(df.columns).index(col)  # 列序号
            letter = get_column_letter(index + 1)  # 列字母
            collen = df[col].apply(lambda x: len(str(x).encode())).max()  # 获取这一列长度的最大值 当然也可以用min获取最小值 mean获取平均值
            ws.column_dimensions[letter].width = collen * 1.2 + 2  # 也就是列宽为最大长度*1.2 可以自己调整
    wb.save(filename)
col = input('请输入列标题: ')
Current_Folder_path = os.getcwd()
# 遍历当前文件夹下的文件
file_list = os.listdir(Current_Folder_path)
if not os.path.exists(Current_Folder_path + '\\New Folder'):
    os.mkdir(Current_Folder_path + '\\New Folder')
docx_list = [a for a in file_list if a.endswith('.xlsx')]
print('正在读取文件')
df = pd.read_excel(Current_Folder_path + '\\' + docx_list[0], dtype='object')
print('文件读取成功')
# 获取某一列唯一值
listType = df[col].unique()
# 循环筛选导出
for i in listType:
    df0 = df.loc[(df[col] == i)]
    df0.to_excel(Current_Folder_path + '\\New Folder\\' + str(i) + '.xlsx', index=False)
    print(str(i) + '导出成功')
file_list = os.listdir(Current_Folder_path + '\\New Folder')
for j in file_list:
    reset_col(Current_Folder_path + '\\New Folder\\' + str(j))
    print(str(j) + '格式转换成功')
