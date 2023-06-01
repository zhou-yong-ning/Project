from PIL import Image
import os
def rotate_image(image_path, degrees):
    # 打开图片
    image = Image.open(image_path)
    # 旋转图片
    rotated_image = image.rotate(degrees, expand=True)
    # 保存旋转后的图片
    rotated_image.save(image_path)
    print(image_path,'旋转完成')


# # 测试函数
# rotate_image('1_1-72.jpg', -90)  # 旋转角度与路径

Current_Folder_path = os.getcwd()
# 遍历当前文件夹下的文件
file_list0 = os.listdir(Current_Folder_path)
# 提取后缀为pdf的文件
file_list = os.listdir(Current_Folder_path)
jpg_list = [a for a in file_list if a.endswith('.jpg')]
for i in jpg_list:
    rotate_image(i,-90)

