import requests
from fake_useragent import UserAgent
import pandas as pd

rename_df = pd.read_excel('图片爬取.xlsx', dtype='object')
for i in rename_df.index.tolist():
    # pandas条件查询
    src_url = rename_df.loc[i, 'url']
    name = rename_df.loc[i, 'name']
    src_headers = {
        'User-Agent': UserAgent().chrome
    }
    src_resp = requests.get(url=src_url, headers=src_headers).content
    with open('./新建文件夹/' + name, 'wb') as f:
        f.write(src_resp)
    print(name + '下载成功')
