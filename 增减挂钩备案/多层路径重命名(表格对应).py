import os
import pandas as pd

rename_df = pd.read_excel('rename.xlsx', dtype='object', index_col='old')
# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 建立旧文件名路径列表
rename_list = []
for Folder_path, folder_path, file_path in os.walk(Current_Folder_path):
    # 文件夹路径   子文件夹列表    文件列表
    if len(folder_path) == 0:
        rename_list.append(Folder_path)
    rename_list = list(set(rename_list))
for old in rename_list:
    new_1 = str(old.split('\\')[-1])
    new_2 = '\\'.join(old.split('\\')[0:-1])
    new_3 = rename_df.loc[new_1, 'new']
    new = os.path.join(new_2, new_3)
    print(new)
    os.rename(old, new)
