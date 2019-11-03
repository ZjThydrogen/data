import xlrd
import numpy as np
"""
本方法实现上游水位库容拟合
x-upz：m
y-upv:e4m
"""

#读取数据
excel = xlrd.open_workbook('D:\pyprject\PBUC_20190807\dataforpeakloadusingip\data.xlsx')
upzv_table = excel.sheets()[0]

def nihe_function(x, y, n):
    coe = np.polyfit(x, y, n)
    return np.poly1d(coe),coe
def get_mse(records_real, records_predict):
    if len(records_real) == len(records_predict):
        return sum([(x - y) ** 2 for x, y in zip(records_real, records_predict)]) / len(records_real)
    else:
        return None

upz_upv = []
upz_upv_coe = []

startcol = 0
for i in range(13):
    end_row = int(upzv_table.col_values(startcol + 1, start_rowx=0, end_rowx=1)[0])
    upz = upzv_table.col_values(startcol, start_rowx=1, end_rowx=end_row+1)
    upv = upzv_table.col_values(startcol+1, start_rowx=1, end_rowx=end_row + 1)

    startcol += 2

    fitupzv, fitupzv_coe = nihe_function(upv, upz, 2)
    upz_upv.append(fitupzv)
    upz_upv_coe.append(fitupzv_coe)

    print(get_mse(upz, fitupzv(upv)))

# testv = 163079.6292919862*0.5+163043.44991856642*0.5
# print(upz_upv_coe[12][0]*testv**2+upz_upv_coe[12][1]*testv+upz_upv_coe[12][2])
# print(upz_upv[12](testv))
