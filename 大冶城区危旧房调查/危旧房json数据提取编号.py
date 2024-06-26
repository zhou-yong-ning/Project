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
    data = TXT(j)
    # 页码
    try:
        qsh = data['content']['pageNum']
        for i in data['content']['list']:
            list0 = [i["collectionSurveyNum"], i["townName"], i["villageName"], i["investHouseno"], i["buildName"],
                     i["upFloors"], i["buildHeight"], qsh]
            DFLIST.append(list0)
    except:
        shutil.move(j, os.path.join('NewFolder', j))
# 二维列表转df
df = pd.DataFrame(DFLIST)
df.to_excel(os.path.join(Current_Folder_path, 'NewFolder\\文件提取.xlsx'))
