import os

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 建立旧文件名路径列表
rename_list = []
for Folder_path, folder_path, file_path in os.walk(Current_Folder_path):
    # 文件夹路径   子文件夹列表    文件列表
    # 提取所有子文件夹路径并去重
    if len(folder_path) == 0:
        rename_list.append(Folder_path)
    rename_list = list(set(rename_list))
for old in rename_list:
    new_name = str(old.split('-')[-1])
    # 文件夹路径
    path_folder = '\\'.join(old.split('\\')[0:-1])
    # 文件路径拼接
    new = os.path.join(path_folder, new_name)
    # 重命名
    os.rename(old, new)
