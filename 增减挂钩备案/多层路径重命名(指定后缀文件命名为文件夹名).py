import os

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 建立旧文件名路径列表
rename_list = []
for Folder_path, folder_path, file_path in os.walk(Current_Folder_path):
    # 文件夹路径   子文件夹列表    文件列表
    for i in file_path:
        old_name = os.path.join(Folder_path, i)
        # 提取后缀为.txt的数据
        if old_name.endswith('.txt'):
            rename_list.append(old_name)
for old in rename_list:
    new_1 = str(old.split('\\')[-2]) + '.txt'
    new_2 = '\\'.join(old.split('\\')[0:-1])
    new = os.path.join(new_2, new_1)
    os.rename(old, new)
