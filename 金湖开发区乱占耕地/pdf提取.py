import pandas as pd
import pdfplumber
import os

def extract_pdf(doc_path):
    file = pdfplumber.open(doc_path)
    # 指定页面
    first_page = file.pages[0]
    # 提取该页面表格类
    col_values1 = first_page.extract_tables()[0]
    # 提取页面文本类
    txt_values1 = first_page.extract_text()
    # 提取字符串
    doc_values = [col_values1[0][1],
                  col_values1[0][3],
                  col_values1[1][1],
                  col_values1[2][1],
                  col_values1[3][1]]
    return txt_values1.split('\n')[1:2]+doc_values
Current_Folder_path = os.getcwd()
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(Current_Folder_path + '\\NewFolder')
# 遍历当前文件夹下的文件
file_list0 = os.listdir(Current_Folder_path)
# 提取后缀为pdf的文件
file_list = os.listdir(Current_Folder_path)
pdf_list = [a for a in file_list if a.endswith('.pdf')]
# 创建空列表
DF = []
for pdf_file in pdf_list:
    # 调用pdf数据读取函数
    pdf = extract_pdf(os.path.join(Current_Folder_path,pdf_file))
    # 单个列表中插入文件名
    pdf.insert(0,pdf_file)
    # 添加子列表
    DF.append(pdf)
    print(pdf)
df = pd.DataFrame(DF,columns=['文件名','图斑编号','建设或占有使用主体姓名（单位）','证件号码','主体类型','房屋性质','房屋坐落'])
print('文件写入中......')
df.to_excel(os.path.join(Current_Folder_path,'NewFolder\\文件提取.xlsx'))
print('文件写入完成')
