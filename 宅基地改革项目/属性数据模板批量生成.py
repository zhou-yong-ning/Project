import pandas as pd
from win32com import client as wc
import os
import openpyxl as vb


def huzhu(guanxi):
    if guanxi == '本人':
        return "是"
    else:
        return "否"


# 获取当前文件夹路径
Current_Folder_path = os.getcwd()
# 遍历当前文件夹下的文件
file_list = os.listdir(Current_Folder_path)
# 提取后缀为xls的文件
xls_list = [a for a in file_list if a.endswith('.xls')]
# 将xls转换为xlsx
if not len(xls_list) == 0:
    try:
        word = wc.Dispatch("Excel.Application")
    except:
        word = wc.Dispatch("kwps.application")
    for A in xls_list:
        doc = word.Workbooks.Open(Current_Folder_path + "\\" + A)
        A = A.replace('.docx', '.doc')
        doc.SaveAs(Current_Folder_path + "\\" + A + "x", 51)
        doc.Close()
        print(A + "转换完成")
        os.remove(Current_Folder_path + "\\" + A)
    word.Quit()

xlsx_list = [a for a in file_list if a.endswith('.xlsx')]
# 利用推导式获取“.*成员信息.xlsx”数据表
# 或者
cyxx = [i for i in xlsx_list if '成员信息' in i]
# 提取文件内容
df0 = pd.read_excel(os.path.join(Current_Folder_path, cyxx[0]), dtype='object')
mingchen = df0.loc[0, 'Unnamed: 1']
daima = df0.loc[0, 'Unnamed: 3']
cpy = df0.loc[1, 'Unnamed: 2']
qydm = df0.loc[1, 'Unnamed: 3']
# 删除指定行
df0 = df0.drop(df0.index[[0, 1, 2, 3]])
# 两列数据合并
df0['通讯地址'] = df0['Unnamed: 12'] + df0['Unnamed: 2']
df0 = df0.loc[:, ['Unnamed: 1', 'Unnamed: 3',
                  'Unnamed: 4', 'Unnamed: 6', 'Unnamed: 7',
                  'Unnamed: 13', '通讯地址']]
df0.rename(columns={'Unnamed: 1': '农户预编码',
                    'Unnamed: 3': '姓名',
                    'Unnamed: 4': '性别',
                    'Unnamed: 6': '证件号码',
                    'Unnamed: 7': '与户主关系',
                    'Unnamed: 13': '联系电话',
                    }, inplace=True)
# 替换某一列信息并强制转换值类型
df0['是否户主'] = df0.apply(lambda x: huzhu(x.与户主关系), axis=1)
df0['是否五保户'] = '否'
df0['是否贫困户'] = '否'
df0['是否本集体经济组织成员'] = '是'
df0['资格权保障情况'] = ''
df0['证件类型'] = '身份证'
df0['户口类型'] = '农业户口'
df0['户籍是否在本村'] = '是'
df0['户籍不在本村原因'] = ''
df0['户籍不在本村原因其他'] = ''
df0['成员备注'] = ''
df0['成员备注说明'] = ''
df0['农户代码'] = ''
df0['备注'] = ''
# cpy = 'QS'
# qydm = '420281007022'

# cpy = input('请输入村名拼音简写： ')
# qydm = input('请输入区域代码： ')
# 跟据现有信息新增列
df0['农户预编码'] = df0.apply(lambda x: cpy + x.农户预编码, axis=1)
# 调整列顺序
df0 = df0[['农户预编码', '通讯地址', '是否户主',
           '是否五保户', '是否贫困户', '资格权保障情况',
           '姓名', '证件类型', '证件号码', '性别', '与户主关系',
           '户口类型', '联系电话', '户籍是否在本村', '户籍不在本村原因', '户籍不在本村原因其他',
           '是否本集体经济组织成员', '成员备注', '成员备注说明', '备注', '农户代码']]
# 创建不同的工作表
write = pd.ExcelWriter(os.path.join(Current_Folder_path, '属性数据模板.xlsx'))
# df0.to_excel(write, sheet_name='农户及户内成员', index=False)

df1 = df0.loc[(df0['与户主关系'] == '本人')]
df1 = df1.drop(columns='与户主关系')
df1['所有权人代码'] = daima
df1['宗地预编码'] = ''
df1['农民房屋预编码'] = df1['农户预编码']
df1['农户预编码'] = ''
df1['使用权人代码'] = ''
df1['使用权人代表姓名'] = df1['姓名']
df1['不动产单元号'] = ''
df1['不动产权证号'] = ''
df1['权证印刷序列号'] = ''
df1['发证机关'] = ''
df1['所属行业'] = ''
df1['国家/地区'] = '中国'
df1['户籍所在省市'] = ''
df1['地址'] = df1['通讯地址']
df1['是否使用权人之间共有'] = '否'
df1['分摊宗地面积'] = ''
df1['权利人类型'] = '个人'
df1['是否本农村集体经济组织成员'] = '是'
df1['户口类型'] = '农业户口'
df1['备注'] = ''
df1['电话'] = df1['联系电话']
df1['使用权人代表证件类型'] = '身份证'
df1['使用权人代表证件号码'] = df1['证件号码']
df1['区域代码'] = qydm
# 调整列顺序
df1 = df1[['所有权人代码', '宗地预编码', '农民房屋预编码', '农户预编码', '使用权人代码', '使用权人代表姓名', '使用权人代表证件类型', '使用权人代表证件号码',
           '不动产单元号', '不动产权证号', '权证印刷序列号', '发证机关', '所属行业', '国家/地区', '户籍所在省市', '性别', '电话', '地址',
           '是否使用权人之间共有', '分摊宗地面积', '权利人类型', '是否本农村集体经济组织成员', '户口类型', '备注', '区域代码']]
# df1.to_excel(write, sheet_name='使用权人', index=False)

df2 = pd.DataFrame([['所有权人代码', '所有权人名称', '所有权性质', '所有权确权登记面积', '代表人姓名',
                     '代表人证件类型', '代表人证件号码', '代表人联系电话', '代表人通讯地址', '代表人邮政编码',
                     '代理人姓名', '代理人证件类型', '代理人证件号码', '代理人联系电话', '代理人通讯地址',
                     '代理人邮政编码', '是否成立农村集体经济组织', '区域代码', '发包方代码']])
df2.to_excel(write, sheet_name='所有权人', header=False, index=False)
df1.to_excel(write, sheet_name='使用权人', index=False)
df0.to_excel(write, sheet_name='农户及户内成员', index=False)
# 保村并关闭属性表
write.save()
write.close()
# 读取附码表的值
fmdj = [j for j in xlsx_list if '赋码' in j]
workbook0 = vb.load_workbook(os.path.join(Current_Folder_path, fmdj[0]))
worksheet0 = workbook0['信息采集表']
xingming = worksheet0['B4'].value
sfzhm = worksheet0['E4'].value
dbryzbm = worksheet0['E7'].value
dbrtxdz = worksheet0['B6'].value
# 打开属性表
workbook1 = vb.load_workbook(os.path.join(Current_Folder_path, '属性数据模板.xlsx'))
worksheet1 = workbook1['所有权人']
worksheet1['A2'] = daima
worksheet1['B2'] = mingchen
worksheet1['C2'] = '村集体经济组织'
worksheet1['E2'] = xingming
worksheet1['G2'] = sfzhm
worksheet1['I2'] = dbrtxdz
worksheet1['J2'] = dbryzbm
worksheet1['Q2'] = '是'
worksheet1['R2'] = qydm
workbook1.save(os.path.join(Current_Folder_path, '属性数据模板.xlsx'))
print('提取成功')
