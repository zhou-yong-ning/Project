import os
import shutil
import zipfile


# 压缩文件
def zipDir(dirPath, outFileName=None):
    '''
    :param dirPath: 目标文件或文件夹
    :param outFileName: 压缩文件名,默认和目标文件相同
    :return: none
    '''
    # 如果不校验文件也可以生成zip文件，但是压缩包内容为空，所以最好有限校验路径是否正确
    if not os.path.exists(dirPath):
        raise Exception("文件路径不存在：", dirPath)

    # zip文件名默认和被压缩的文件名相同，文件位置默认在待压缩文件的同级目录
    if outFileName == None:
        outFileName = dirPath + ".zip"

    if not outFileName.endswith("zip"):
        raise Exception("压缩文件名的扩展名不正确！因以.zip作为扩展名")

    zip = zipfile.ZipFile(outFileName, "w", zipfile.ZIP_DEFLATED)
    # 遍历文件或文件夹
    for path, dirnames, filenames in os.walk(dirPath):
        fpath = path.replace(dirPath, "")

        for filename in filenames:
            zip.write(os.path.join(path, filename), os.path.join(fpath, filename))
    zip.close()


# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取shp名字
filename0 = [i for i in file_list if i.endswith('shp') and 'Z' not in i.split('.')[0]]
# 去掉后缀
filename1 = [j.split('.')[0] for j in filename0]
# 创建文件夹
for k in filename1:
    # 创建文件夹
    os.mkdir(os.path.join(Current_Folder_path, k))
    os.mkdir(os.path.join(Current_Folder_path, k, '宗地'))
    os.mkdir(os.path.join(Current_Folder_path, k, '自然幢'))
    # 生成文件列表
    file = [i for i in file_list if i.split('.')[0] == k]
    fileZ = [j for j in file_list if j.startswith(k+'Z')]
    # 循环移动宗地shp文件
    for a in file:
        shutil.move(os.path.join(Current_Folder_path, a), os.path.join(Current_Folder_path, k, '宗地', a))
    # 循环移动幢shp文件
    for b in fileZ:
        shutil.move(os.path.join(Current_Folder_path, b), os.path.join(Current_Folder_path, k, '自然幢', b))

    zipDir(os.path.join(Current_Folder_path, k))

    
# 自然幢
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取shp名字
filename0 = [i for i in file_list if i.endswith('shp')]
# 去掉后缀
filename1 = [j.split('.')[0] for j in filename0]
# 创建文件夹
for k in filename1:
    # 创建文件夹
    os.mkdir(os.path.join(Current_Folder_path,k))
    os.mkdir(os.path.join(Current_Folder_path,k,'自然幢'))
    # 生成文件列表
    file = [i for i in file_list if i.startswith(k)]
    for j in file:
        shutil.move(os.path.join(Current_Folder_path,j),os.path.join(Current_Folder_path,k,'自然幢',j))
    zipDir(os.path.join(Current_Folder_path,k))
