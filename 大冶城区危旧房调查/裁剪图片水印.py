import os
from PIL import Image

# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
file_list = [i for i in file_list if i.endswith('.jpg')]
# 创建成果输出文件夹
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(os.path.join(Current_Folder_path, 'NewFolder'))
for i in file_list:
    img = Image.open(i)
    # 打印图片尺寸
    print(img.size)
    # 裁剪
    cropped = img.crop((0, 0, img.size[0], img.size[1]-50))  # (left, upper, right, lower)
    # 输出裁剪成果
    cropped.save(os.path.join(Current_Folder_path,"NewFolder","裁剪_"+i))