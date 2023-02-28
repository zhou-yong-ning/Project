import pandas as pd
import matplotlib.pyplot as plt
import warnings


# 统计图方法（传入数据透视表，根据表格生成统计图）
def Statistical_chart(DF,name):
    # 遇到数据中有中文的时候，一定要先设置中文字体
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 用黑体显示中文
    # 解决坐标轴负号问题
    plt.rcParams['axes.unicode_minus'] = False
    # 图片展示(调整图片大小与分辨率)
    plt.figure(dpi=200, figsize=(6, 4))
    # 画柱状图：x轴是姓名，y轴是分类，颜色是红色
    plt.bar(DF.index, DF.数量, label="数量", color="deepskyblue", alpha=1, width=0.5, )
    # 显示图例
    plt.legend()
    # 设置X与Y轴的标题
    plt.xlabel("宅基地利用状况",
               labelpad=-12,  # 调整x轴标签与x轴距离
               x=0,  # 调整x轴标签的左右位置
               fontsize=9,
               bbox={'facecolor': 'silver', 'pad': 2, 'alpha': 0.2})
    plt.ylabel("数量")
    # 刻度标签及文字旋转
    plt.xticks(DF.index, rotation=30)
    # y轴的刻度范围
    plt.ylim([0, DF['数量'].max() + DF['数量'].max()//10])  # 获取数量最大值
    # 设置图表的标题、字号、粗体
    plt.title(name + "宅基地利用状况分类统计图", fontsize=12, fontweight='bold')
    # 添加数据标签
    for x, y in enumerate(DF.数量):  # X代表下标，Y代表值
        plt.text(x, y + 2, str(y), ha='center', va='bottom', fontsize=10, rotation=0)
    # 调整图片上下左右边距，
    plt.subplots_adjust(left=None, bottom=0.3, right=None, top=0.9, wspace=None, hspace=0.4)
    # 图片展示
    # plt.show()
    return plt


# 数据透视方法（根据数据生成数据透视表）
def Pivot_ZD(Df):
    Df1 = pd.pivot_table(Df, index='宅基地利用状况', values='宅基地代码', aggfunc=len)
    Df1.sort_values(by='宅基地利用状况', inplace=True, ascending=False)
    Df1.rename(columns={'宅基地代码': '数量'}, inplace=True)  # 在原数据上修改
    return Df1

# 分表方法
def dxcfor(dfbiao, list, col):
    for i in list:
        df0 = dfbiao.loc[(dfbiao[col] == i)]
        # 数据透视表
        df = Pivot_ZD(df0)
        # 调用图片生成统计图表函数
        try:  # 防止出现空值报错
            tp = Statistical_chart(df,i)
            tp.savefig(i + '.jpg')
            print(i + '.jpg', '导出成功')
        except:
            print(i, '存在问题')



if __name__ == '__main__':
    warnings.filterwarnings('ignore')      # 忽略警告
    col = '所属区划'      # 指定分割列名称
    print('正在读取文件')
    df = pd.read_table('宗地信息表数据.txt', sep='\t', encoding='UTF-8', dtype='object')      # 读取数据
    print('文件读取成功')
    listType = df[col].unique()      # 获取某一列唯一值
    dxcfor(df, listType, col)       # 分表
