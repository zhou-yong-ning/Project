import os
import pandas as pd

rename_df = pd.read_excel('rename.xlsx', dtype='object')
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
    # 旧文件夹名
    old_name = str(old.split('\\')[-1])
    # 子文件夹所在路径
    path_floder = '\\'.join(old.split('\\')[0:-1])
    # 新文件夹名
    new_name = rename_df.loc[rename_df['old'] == old_name, 'new'].tolist()[0]
    # 新文件夹路径
    new_name_path = os.path.join(path_floder, new_name)
    os.rename(old,new_name_path)
