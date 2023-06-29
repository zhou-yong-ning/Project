import zipfile
import os


def compress_files_to_zip(zip_path, file_paths):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for file_path in file_paths:
            if os.path.isfile(file_path):
                file_name = os.path.basename(file_path)
                zip_file.write(file_path, file_name)
            else:
                print(f"Warning: '{file_path}' is not a valid file. Skipping...")
    print(f"Files compressed successfully to '{zip_path}'")


# 示例用法
# file_paths = ['file1.txt', 'file2.txt', 'file3.txt']
# zip_path = 'compressed_files.zip'
# compress_files_to_zip(zip_path, file_paths)


# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(Current_Folder_path + '\\NewFolder')
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取shp名字
filename0 = [i for i in file_list if i.endswith('shp')]
# 去掉后缀
filename1 = [j.split('.')[0] for j in filename0]
for k in filename1:
    # 生成文件列表
    file = [i for i in file_list if i.startswith(k)]
    # 生成压缩包路径
    zip_file = os.path.join(Current_Folder_path, 'NewFolder', '.'.join([k, 'zip']))
    compress_files_to_zip(zip_file, file)
