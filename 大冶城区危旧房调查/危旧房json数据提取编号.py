import json
import os
import pandas as pd

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
for i in txt_list:
    data = TXT(i)
    for i in data['content']['list']:
        list0 = [i["collectionSurveyNum"], i["townName"], i["villageName"], i["investHouseno"]]
        print(list0)
        DFLIST.append(list0)
# 二维列表转df
df = pd.DataFrame(DFLIST)
df.to_excel(os.path.join(Current_Folder_path,'NewFolder\\文件提取.xlsx'))
