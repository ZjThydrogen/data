import xlrd
import numpy as np
import data.Z_V_fitforlcjjsj as zvfit


"""
本方法实现耗水率上游水位拟合拟合
x-upz：m
y-upv:e4m
"""

#读取数据
excel = xlrd.open_workbook('D:\pyprject\PBUC_20190807\dataforpeakloadusingip\data.xlsx')

#读取起始库容
initv_table = excel.sheets()[1]
initv = initv_table.col_values(3, start_rowx=1, end_rowx=14)
#转化为起始水位
initz = [zvfit.upz_upv[i](initv[i])for i in range(13)]
print(initz)

#对水位耗水率拟合时只在起始水位的上下两米的范围内进行拟合
hslupz_table = excel.sheets()[5]

def nihe_function(x, y, n):
    coe = np.polyfit(x, y, n)
    return np.poly1d(coe),coe

def get_mse(records_real, records_predict):
    if len(records_real) == len(records_predict):
        return sum([(x - y) ** 2 for x, y in zip(records_real, records_predict)]) / len(records_real)
    else:
        return None

def get_maxindex(a, x):
    for k in range(len(x)):
        if x[k] < a:
            continue
        else:
           return k

#对拟合数据集阶段
def get_haoshuilv_upz_tofit(z, upz, hsl):
    maxfitx = z+2
    minfitx = z-2

    minindex = get_maxindex(minfitx, upz)
    maxindex = get_maxindex(maxfitx, upz)

    return upz[minindex:maxindex],hsl[minindex:maxindex]


startcol = 0

hsl_upz =[]
hsl_upz_coe = []
#[a,b] a*x+b

for i in range(13):
    print(i)
    end_row = int(hslupz_table.col_values(startcol+1, start_rowx=0, end_rowx=1)[0])
    upz = hslupz_table.col_values(startcol, start_rowx=1, end_rowx=end_row+1)
    hsl = hslupz_table.col_values(startcol+1, start_rowx=1, end_rowx=end_row+1)
    startcol += 2

    upz, hsl = get_haoshuilv_upz_tofit(initz[i], upz, hsl)
    print(initz[i])
    print(upz)
    print(hsl)

    #拟合
    fithslupz, fithslupz_coe = nihe_function(upz, hsl, 1)
    hsl_upz.append(fithslupz)
    hsl_upz_coe.append(fithslupz_coe)
    print(get_mse(hsl, fithslupz(upz)))

    # plt.plot(upz, fithslupz(upz))
    # plt.scatter(upz, hsl, c='r',marker='o')
    # plt.show()
    # plt.close()

print(hsl_upz[0](1400.5))
print(hsl_upz_coe[0][0]*1400.5+hsl_upz_coe[0][1])





