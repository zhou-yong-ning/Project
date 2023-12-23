import os

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 建立旧文件名路径列表
rename_list = []
for Folder_path, folder_path, file_path in os.walk(Current_Folder_path):
    # 文件夹路径   子文件夹列表    文件列表
    for i in file_path:
        old_name = os.path.join(Folder_path, i)
        # 提取所有文件路径
        rename_list.append(old_name)
for old in rename_list:
    new_name = str(old.split('-')[-1])
    # 文件夹路径
    path_folder = '\\'.join(old.split('\\')[0:-1])
    # 文件路径拼接
    new = os.path.join(path_folder, new_name)
    # 重命名
    os.rename(old, new)
