import json
import os
import pandas as pd
import shutil


def TXT(txt_file):
    with open(file=txt_file, mode="r", encoding="utf-8") as f:
        data = f.read()
        a = json.loads(data)
        # 返回字典
        return a


Current_Folder_path = os.getcwd()
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(Current_Folder_path + '\\NewFolder')
# 遍历当前文件夹下的文件
file_list0 = os.listdir(Current_Folder_path)
# 提取后缀为txt的文件
file_list = os.listdir(Current_Folder_path)
txt_list = [a for a in file_list if a.endswith('.txt')]
# 创建空列表
DFLIST = []
for j in txt_list:
    # 页码
    try:
        data = TXT(j)
        WJ_fwbh = data['content']['dilapidated']['collectionSurveyNum']
        for i in data['content']['dilapidated']['dilapidatedPhoto']:
            list1 = [i['filePath'],WJ_fwbh+'现场查看情况']
            DFLIST.append(list1)
        for k in data['content']['dilapidated']['riskScenePhoto']:
            list2 = [k['filePath'], WJ_fwbh + '基本情况']
            DFLIST.append(list2)
    except:
        shutil.move(j, os.path.join('NewFolder', j))
# 二维列表转df
df = pd.DataFrame(DFLIST)
df.to_excel(os.path.join(Current_Folder_path, 'NewFolder\\文件提取.xlsx'))
