import pandas as pd
import matplotlib.pyplot as plt
import warnings


# 统计图方法（传入数据透视表，根据表格生成统计图）
def Statistical_chart(DF, name):
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
    plt.xlabel("利用状况",
               labelpad=2,  # 调整x轴标签与x轴距离
               x=0,  # 调整x轴标签的左右位置
               fontsize=9,
               bbox={'facecolor': 'silver', 'pad': 2, 'alpha': 0.2})
    plt.ylabel("数量")
    # 刻度标签及文字旋转
    plt.xticks(DF.index, rotation=0)
    # y轴的刻度范围
    plt.ylim([0, DF['数量'].max() + DF['数量'].max() // 10])  # 获取数量最大值
    # 设置图表的标题、字号、粗体
    plt.title(name + "房屋利用状况分类统计图", fontsize=12, fontweight='bold')
    # 添加数据标签
    for x, y in enumerate(DF.数量):  # X代表下标，Y代表值
        plt.text(x, y + 2, str(y), ha='center', va='bottom', fontsize=10, rotation=0)
    # 调整图片上下左右边距，
    plt.subplots_adjust(left=None, bottom=0.15, right=None, top=0.9, wspace=None, hspace=0.4)
    # 图片展示
    # plt.show()
    return plt


# 数据透视方法（根据数据生成数据透视表）
def Pivot_ZD(Df):
    Df1 = pd.pivot_table(Df, index='利用状况', values='农民房屋代码', aggfunc=len)
    Df1['id'] = Df1.index
    Df1['level'] = Df1.apply(lambda x: getlevel(x['id']), axis=1)
    Df1.sort_values(by='level', inplace=True, ascending=True)
    Df1.rename(columns={'农民房屋代码': '数量'}, inplace=True)  # 在原数据上修改
    return Df1


# 分表方法
def dxcfor(dfbiao, list, col):
    for i in list:
        df0 = dfbiao.loc[(dfbiao[col] == i)]
        try:  # 防止出现空值报错
            df = Pivot_ZD(df0)  # 数据透视表
            tp = Statistical_chart(df, i)  # 调用图片生成统计图表函数
            tp.savefig(i + '.jpg')
            # print(i + '.jpg', '导出成功')
        except:
            print(i, '存在问题')


# 创建排序方法
def getlevel(score):
    tinydict = {'自住': 1, '出租居住': 2, '自营': 3, '出租经营': 4,
                '合作（入股）': 5, '闲置': 6, '其他': 7}
    try:
        return tinydict[score]
    except:
        return 100


if __name__ == '__main__':
    warnings.filterwarnings('ignore')  # 忽略警告
    col = '所属区划'  # 指定分割列名称
    print('正在读取文件')
    df = pd.read_table('房屋信息表数据总.txt', encoding='UTF-8', dtype='object')  # 读取数据
    print('文件读取成功')
    listType = df[col].unique()  # 获取某一列唯一值
    dxcfor(df, listType, col)  # 分表
