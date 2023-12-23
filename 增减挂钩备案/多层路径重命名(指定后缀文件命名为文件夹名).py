import os

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 建立旧文件名路径列表
rename_list = []
for Folder_path, folder_path, file_path in os.walk(Current_Folder_path):
    # 文件夹路径   子文件夹列表    文件列表
    for i in file_path:
        # 提取所有文件
        old_name = os.path.join(Folder_path, i)
        # 此处可用推导式提取指定文件
        if not old_name.endswith('.py'):
            rename_list.append(old_name)
for old in rename_list:
    # 提取文件夹所在文件夹路径
    old_folder_path = '\\'.join(old.split('\\')[0:-1])
    # 提取文件所在文件夹名
    folder_name = str(old.split('\\')[-2])
    # 提取文件后缀
    old_name1 = str(old.split('\\')[-1])
    houzhui = old_name1.split('.')[1:]
    # 新文件名
    new_name = '.'.join([folder_name] + houzhui)
    # 新文件名路径
    new_name_path = os.path.join(old_folder_path,new_name)
    # 文件重命名
    os.rename(old, new_name_path)
