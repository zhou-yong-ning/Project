import random
import pandas as pd


def generate_random_string(length):
    # 使用random模块生成随机字符串
    random_string = ''.join(random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789') for _ in range(length))
    return random_string


LIST = []
# 调用方法生成随机字符串
for j in range(10000):
    random_str = generate_random_string(17)
    random_str = '01XL'+ random_str
    LIST.append(random_str)
    print(random_str)
df = pd.DataFrame(LIST)
df.to_excel('ID.xlsx')
