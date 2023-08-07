import os
import pandas as pd

file_path_list = []
# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 循环遍历文件路径  文件夹列表   文件列表
for Folder_path, Folder_list, file_list in os.walk(Current_Folder_path):
    for i in file_list:
        file_path0 = os.path.join(Folder_path, i)
        print(file_path0)
        file_path_list.append(file_path0)

df = pd.DataFrame(file_path_list)
df.to_excel('文件路径提取.xlsx')
