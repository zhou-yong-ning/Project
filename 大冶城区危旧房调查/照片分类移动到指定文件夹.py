import os
import shutil

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
file_list = [i for i in file_list if i.endswith('.jpg')]
# 截取前17位字符串
filename0 = [i[0:17] for i in file_list]
# 去除列表中的重复项（顺序改变）
filename0 = list(set(filename0))
# 根据编号建立文件夹
for j in filename0:
    if not os.path.exists(os.path.join(Current_Folder_path, j)):
        os.mkdir(os.path.join(Current_Folder_path, j))
# 移动文件
for k in file_list:
    shutil.move(os.path.join(Current_Folder_path, k), os.path.join(Current_Folder_path, k[0:17], k))
