from PyPDF2 import PdfMerger
import os


# pdf合并方法
def merge_pdfs(file_paths, output_path):
    # 创建一个PdfMerger对象
    merger = PdfMerger()
    for path in file_paths:
        # 使用append()方法依次将每个PDF文件添加到合并器中
        merger.append(path)
        print(path, '合并成功！')
    # 使用write()方法将合并后的PDF写入指定的输出路径
    merger.write(output_path)
    # 使用close()方法关闭合并器
    merger.close()


# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
if not os.path.exists(Current_Folder_path + '\\NewFolder'):
    os.mkdir(Current_Folder_path + '\\NewFolder')
# 获取路径下的所有文件
file_list = os.listdir(Current_Folder_path)
# 提取后缀为‘.pdf’的文件
pdf_files = [i for i in file_list if i.endswith('.pdf')]
# 合并'_'前相同pdf
pdf_files_BH = [j.split('_')[0] for j in pdf_files]
# 相同前缀列表去重
pdf_files_BH = list(set(pdf_files_BH))
for k in pdf_files_BH:
    # 指定输出路径
    output_file = os.path.join(Current_Folder_path,'NewFolder',k + '.pdf')
    # 相同前缀创建列表
    pdf_files_list = [i for i in file_list if i.startswith(k)]
    # print(pdf_files_list)
    # 调用方法
    merge_pdfs(pdf_files_list, output_file)
