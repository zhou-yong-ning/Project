import pandas as pd
import os

Current_Folder_path = os.getcwd()
# 遍历当前文件夹下的文件
file_list0 = os.listdir(Current_Folder_path)
# 提取后缀为pdf的文件
file_list = os.listdir(Current_Folder_path)
xlsx_list = [a for a in file_list if a.endswith('.xlsx')]

df = pd.read_excel(xlsx_list[0], dtype='object', index_col='ID')
id_list = list(set(df.index.tolist()))

DF_LIST = []
for i in id_list:
    # pandas条件查询
    idmerge_list = [i, 'POLYGON Z (('+','.join(df.loc[i, 'XY'].tolist())+'))']
    print(idmerge_list)
    DF_LIST.append(idmerge_list)
df_shuchu = pd.DataFrame(DF_LIST, columns=['ID', 'Geometry'])
df_shuchu.to_excel('坐标合并.xlsx')
