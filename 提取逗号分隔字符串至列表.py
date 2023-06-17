
import pandas as pd
with open("新建文本文档.txt", "r", encoding='utf-8') as f:  #打开文本
    data = f.read()   #读取文本
data = data.split(',')

df = pd.DataFrame(data,columns=['文件'])
df.to_excel('文件提取.xlsx',index=None)
print(data)

