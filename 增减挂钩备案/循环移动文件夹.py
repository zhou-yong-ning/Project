import os
import shutil

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取shp名字
filename0 = [i for i in file_list if '.' not in file_list]
# 去掉后缀
filename1 = [j.split('-')[0] for j in filename0]
filename1 = list(set(filename1))
# 创建文件夹
for k in filename1:
    # 创建文件夹
    if not os.path.exists(os.path.join(Current_Folder_path, k)):
        os.mkdir(os.path.join(Current_Folder_path, k))
    # 生成文件列表
    file = [i for i in file_list if i.split('-')[0] == k and ('.' not in i and '-'in i)]
    # 循环移动文件
    for a in file:
        shutil.move(os.path.join(Current_Folder_path, a), os.path.join(Current_Folder_path, k, a))
