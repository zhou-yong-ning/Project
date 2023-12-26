import pandas as pd
import os


# txt转excel
def vertex_txt(file_paths, output_path):
    # 读取txt文件
    df = pd.read_csv(file_paths, header=None, dtype='object')
    df.to_excel(output_path)


def readtxt(txt_file):
    with open(txt_file, "r", encoding='utf-8') as f:  # 打开文本
        data = f.readlines()  # 读取文本(data变量为列表)
        if data[0] == '[属性描述]\n':
            data = data[6:]
        else:
            pass
    if '-' not in data[1]:
        Data_list = []
        for i in data:
            if '@\n' in i:
                pch = i.split(',')[3]
                print(pch)
            else:
                j = i.replace('\n', ',' + pch + '\n')
                Data_list.append(j)
    else:
        Data_list = data
    with open(txt_file, "w", encoding='utf-8') as f:  # 打开文本
        f.writelines(Data_list)  # 写入文本


# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 创建新建文件夹
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(Current_Folder_path + '\\NewFolder')
# 提取后缀为‘.txt’的文件
txt_files = [i for i in file_list if i.endswith('.txt')]
# 遍历txt列表
for j in txt_files:
    # 指定输出路径
    output_file = os.path.join('NewFolder', j.split('.')[0] + '.xlsx')
    # 调用方法
    readtxt(j)
    vertex_txt(j, output_file)
