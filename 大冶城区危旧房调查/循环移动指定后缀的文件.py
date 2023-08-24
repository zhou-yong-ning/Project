import os
import shutil


# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 创建文件夹
if not os.path.exists('NewFolder'):
    os.mkdir(os.path.join('NewFolder'))
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取shp名字
filename = [i for i in file_list if i.endswith('(1).jpg') or i.endswith('(3).jpg')]
# 创建文件夹
for k in filename:
    # 循环移动文件
    # shutil.move(os.path.join(Current_Folder_path, k), os.path.join(Current_Folder_path, 'NewFolder', k))
    # 循环复制文件
    shutil.copy(os.path.join(Current_Folder_path, k), os.path.join(Current_Folder_path, 'NewFolder', k))
