import json
import os


# 读取json
def JSON(json_file):
    f = open(json_file, 'r', encoding='utf-8')
    content = f.read()
    a = json.loads(content)
    return a


# 字符串写入txt
def TO_TXT(content, name):
    with open(name, 'a', encoding='utf-8') as f:
        f.write(content)


Current_Folder_path = os.getcwd()
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(Current_Folder_path + '\\NewFolder')
# 遍历当前文件夹下的文件
file_list0 = os.listdir(Current_Folder_path)
# 提取后缀为json的文件
file_list = os.listdir(Current_Folder_path)
json_list = [a for a in file_list if a.endswith('.json')]
for i in json_list:
    # 提取文件名
    i_name = i.replace('.json', '')
    # 读取json文件
    data = JSON(i)
    # 提取文件列表
    a = data["log"]["entries"]
    # 遍历har文件子列表中的字典
    for j in range(len(a)):
        # 提取文本
        b = a[j]["response"]["content"]["text"]
        print(b)
        # 命名
        txt_name = i_name + '丨' + str(j) + '.txt'
        TO_TXT(b, txt_name)
