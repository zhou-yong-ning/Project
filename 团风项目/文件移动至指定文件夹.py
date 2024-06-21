import pandas as pd
import os
import shutil

Current_Folder_path = os.getcwd()
df1 = pd.read_excel('文件移动.xlsx', dtype='object')
for i in df1.index.tolist():
    # pandas条件查询
    file = df1.loc[i, '文件名']
    file_Folder = df1.loc[i, '文件夹']
    file_Folder1 = os.path.join(Current_Folder_path,file_Folder)
    if not os.path.exists(file_Folder1):
        os.mkdir(file_Folder1)
    shutil.move(file, os.path.join(file_Folder1, file))
