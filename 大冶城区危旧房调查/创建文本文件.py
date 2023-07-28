import os

# 定义要创建的文件数量
num_files = 5

# 定义要创建的文件名前缀
file_prefix = "file"

# 循环创建文件
for i in range(num_files):
    # 拼接文件名
    file_name = f"{file_prefix}{i+1}.txt"
    # 使用 open 函数创建空文本文件
    with open(file_name, "w") as file:
        pass
