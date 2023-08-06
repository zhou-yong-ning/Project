import os
import shutil

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取shp文件列表
filename0 = [i for i in file_list if i.endswith('shp') and 'Z' not in i.split('.')[0]]
# 去掉后缀
filename1 = [j.split('.')[0] for j in filename0]
# 创建文件夹
for k in filename1:
    # 创建文件夹
    os.mkdir(os.path.join(Current_Folder_path, k))
    os.mkdir(os.path.join(Current_Folder_path, k, '宗地'))
    os.mkdir(os.path.join(Current_Folder_path, k, '自然幢'))
    os.mkdir(os.path.join(Current_Folder_path, k, '台账'))
    os.mkdir(os.path.join(Current_Folder_path, k, '利用现状图'))
    # 生成文件列表
    # 宅基地
    file = [i for i in file_list if i.split('.')[0] == k]
    # 自然幢
    fileZ = [j for j in file_list if j.startswith(k + 'Z')]
    # 台账
    fileTZ = [l for l in file_list if l.startswith(k) and l.endswith('.xlsx')]
    # 图件
    fileT = [m for m in file_list if m.startswith(k) and m.endswith('.pdf')]
    # 循环移动宗地shp文件
    for a in file:
        shutil.move(os.path.join(Current_Folder_path, a), os.path.join(Current_Folder_path, k, '宗地', a))
    # 循环移动幢shp文件
    for b in fileZ:
        shutil.move(os.path.join(Current_Folder_path, b), os.path.join(Current_Folder_path, k, '自然幢', b))
    # 循环移动台账
    for c in fileTZ:
        shutil.move(os.path.join(Current_Folder_path, c), os.path.join(Current_Folder_path, k, '台账', c))
    # 循环移动图件
    for d in fileT:
        shutil.move(os.path.join(Current_Folder_path, d), os.path.join(Current_Folder_path, k, '利用现状图', d))
