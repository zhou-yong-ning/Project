import requests
import time
import pandas as pd
import os

headers = {
    'UserAgent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'Accept': '*/*',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,ja;q=0.7',
    'Appversion': 'pc',
    'Authorization': 'Bearer 2701a323-b32e-42c0-a09c-24ff43a1e25f',
    'Connection': 'keep-alive',
    'Dnt': '1',
    'Host': '59.208.147.83:7090',
    'Origin': 'http://59.208.147.83:18089'
}

Current_Folder_path = os.getcwd()
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(Current_Folder_path + '\\NewFolder')
file_list = os.listdir(Current_Folder_path)
excel_list = [a for a in file_list if a.endswith('.xlsx')]
rename_df = pd.read_excel(excel_list[0], dtype='object')
for i in rename_df.index.tolist():
    try:
        # pandas条件查询
        src_url = rename_df.loc[i, 'url1']
        url = 'http://59.208.147.83:7090/risk-census/building/findHouseByFwbh?fwbh='+src_url
        # data_text = requests.post(url = url, data=headers, json=data) # post请求
        data_text = requests.get(url=url, headers=headers).text
        with open('./NewFolder/'+str(i)+'.txt', 'w', encoding='utf-8') as f:
            f.write(data_text)
        print(src_url + '下载成功')
        time.sleep(5)
    except:
        print(i + '下载失败')
