import pandas as pd
import os
from multiprocessing import Process


def dxcfor(dfbiao, list):
    Current_Folder_path = os.getcwd()
    for i in list:
        df0 = dfbiao.loc[(dfbiao['FBFBM'] == i)]
        col_list = ['CBFBM', 'CBFMC', 'DKBM', 'YHTMJM', '原合同面积统计', 'HTMJM', '发证面积统计']
        df0 = df0[col_list]
        df0.rename(columns={
            'CBFBM': '承包方编码',
            'CBFMC': '承包方名称',
            'DKBM': '地块名称',
            'YHTMJM': '原合同面积',
            'HTMJM': '发证面积'
        }, inplace=True)
        # 根据列进行升排列
        df0.sort_values(by=['承包方编码'], ascending=True, inplace=True)
        df0.to_excel(os.path.join(Current_Folder_path, 'NewFolder', i + '.xlsx'), index=False)
        print(i + ' 表格导出成功')


if __name__ == '__main__':
    # col = input('请输入列标题: ')
    col = 'FBFBM'
    Current_Folder_path = os.getcwd()
    # 遍历当前文件夹下的文件
    file_list0 = os.listdir(Current_Folder_path)
    if not os.path.exists(Current_Folder_path + '\\NewFolder'):
        os.mkdir(os.path.join(Current_Folder_path, 'NewFolder'))
    file_list1 = [a for a in file_list0 if a.endswith('.xlsx')]
    print('正在读取文件')
    df = pd.read_excel(os.path.join(Current_Folder_path, file_list1[0]), dtype='object')
    print('文件读取成功')
    # 获取某一列唯一值
    listType = df[col].unique()
    # # 循环筛选导出
    # dxcfor(df, listType)
    listnum = 60
    Arr = []
    for j in range(0, len(listType), listnum):
        if j + listnum <= len(listType):
            p = Process(target=dxcfor, args=(df, listType[j:j + listnum]))
            p.start()
            Arr.append(p)
        else:
            p = Process(target=dxcfor, args=(df, listType[j:len(listType) + 1]))
            p.start()
            Arr.append(p)
    print(Arr)
    for k in Arr:
        k.join()
