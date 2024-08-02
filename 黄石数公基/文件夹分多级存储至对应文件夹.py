import os
import shutil


# 根据文件创建多级文件夹
def CeateFolder(Folder_path, folder_list):
    for i in folder_list:
        if not os.path.exists(os.path.join(Folder_path, i)):
            os.mkdir(os.path.join(Folder_path, i))
        Folder_path = os.path.join(Folder_path, i)
    return Folder_path


# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取shp名字
filename0 = [i for i in file_list if i.endswith('.tif')]
# 去掉后缀
filename1 = [j.split('.')[0] for j in filename0]
# 去除重复元素
filename1 = list(set(filename1))
# 创建文件夹
for k in filename1:
    # 建立分级文件夹
    k_list = k.split('_')
    # 创建多级文件夹
    move_path = CeateFolder(Current_Folder_path,k_list)
    # 生成文件列表
    file = [i for i in file_list if i.split('.')[0] == k]
    # 循环移动shp文件
    for a in file:
        shutil.move(os.path.join(Current_Folder_path, a),
                    os.path.join(move_path, a))