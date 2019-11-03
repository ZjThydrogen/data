import xlrd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib import cm
from matplotlib.ticker import LinearLocator, FormatStrFormatter
import numpy as np
import data.basedata as bd
import math

from scipy.optimize import curve_fit

#读取数据
excel = xlrd.open_workbook('D:\pyprject\PBUC_20190807\data.xlsx')

nld_table = excel.sheets()[3]

plant_upz = nld_table.col_values(0, start_rowx=1, end_rowx=27)
plant_upv = nld_table.col_values(1, start_rowx=1, end_rowx=27)

plant_dwz = nld_table.col_values(2, start_rowx=1, end_rowx=6)
plant_Q = nld_table.col_values(3, start_rowx=1, end_rowx=6)


unit_H = nld_table.col_values(4, start_rowx=1, end_rowx=None)
unit_N = nld_table.col_values(5, start_rowx=1, end_rowx=None)
unit_Q = nld_table.col_values(6, start_rowx=1, end_rowx=None)


def nihe_function(x, y, n):
    coe = np.polyfit(x, y, n)
    return np.poly1d(coe), coe


#4次多项式拟合
UPZ_UPV_fit, UPZ_UPV_fit_coe = nihe_function(np.array(plant_upv), np.array(plant_upz), 4)
DWZ_Q_fit, DWZ_Q_fit_coe = nihe_function(np.array(plant_Q), np.array(plant_dwz), 2)


#NHQ拟合非线性关系为：P = b0+b1*h+b2*q+b3*h*q+b4*h*h+b5*q*q 从matlab读取系数
N_HQ_fit_coe = nld_table.col_values(8, start_rowx=0, end_rowx=6)


def N_HQ_fit(H, Q):
    return N_HQ_fit_coe[0]+N_HQ_fit_coe[1]*H+N_HQ_fit_coe[2]*Q+N_HQ_fit_coe[3]*H*Q+N_HQ_fit_coe[4]*H*H+N_HQ_fit_coe[5]*Q*Q


def get_mse(records_real, records_predict):
    if len(records_real) == len(records_predict):
        return sum([(x - y) ** 2 for x, y in zip(records_real, records_predict)]) / len(records_real)
    else:
        return None


#拟合误差指标
MSE_ZV_FIT = get_mse(plant_upz, UPZ_UPV_fit(plant_upv))
MSE_ZQ_FIT = get_mse(plant_dwz, DWZ_Q_fit(plant_Q))
predictN = []
for i in range(len(unit_N)):
    predictN.append(N_HQ_fit(unit_H[i], unit_Q[i]))
MSE_NHQ_FIT = get_mse(unit_N, predictN)


#zv线性化
#曲线分割
vmin = 10281
vmax = 23685
number_gap_1 = 16

V = np.linspace(vmin, vmax, number_gap_1)
ZU = UPZ_UPV_fit(V)
index_gap_1 = list(range(0, number_gap_1))

Mu = list(range(1, number_gap_1))
V_gap = dict(zip(index_gap_1, V.tolist()))
ZU_gap = dict(zip(index_gap_1, ZU.tolist()))
print("*********ZV曲线线性分割完成***********")
print('起始：%f,结束：%f,共分割成%d段' %(vmin, vmax, number_gap_1-1))
print(Mu)
print(V_gap)
print(ZU_gap)


qmin = 0
qmax = 260
number_gap_2 = 16

U = np.linspace(qmin, qmax, number_gap_2)
ZD = DWZ_Q_fit(U)
index_gap_2 = list(range(0, number_gap_2))

Md = list(range(1, number_gap_2))
U_gap = dict(zip(index_gap_2, U.tolist()))
ZD_gap = dict(zip(index_gap_2, ZD.tolist()))
print("*********ZQ曲线线性分割完成***********")
print('起始：%f,结束：%f,共分割成%d段' %(qmin,qmax,number_gap_2-1))
print(Md)
print(U_gap)
print(ZD_gap)

#NHQ分析

#获得水头最大最小值
print("*********NHQ曲线线性分割离散***********")
H_max = math.ceil((bd.zu_max - DWZ_Q_fit(sum(bd.qunit_min)))[0])
H_min = math.floor((bd.zu_min - DWZ_Q_fit(sum(bd.qunit_max)))[0])

#水头离散
number_gap_h = 5
H = np.linspace(H_min,H_max,number_gap_h)
index_gap_h = list(range(0, number_gap_h))
H_gap = dict(zip(index_gap_h, H.tolist()))
Mh = list(range(1, number_gap_h))
print("*********水头离散完成***********")
print('水头离散成了%d段' % (number_gap_h-1))
print(H_gap)
print(Mh)
#发电流量离散
Q_max = bd.qunit_max[0]
Q_min = bd.qunit_min[0]

number_gap_q = 5
Q = np.linspace(Q_min, Q_max, number_gap_q)
index_gap_q = list(range(0, number_gap_q))
Q_gap = dict(zip(index_gap_q, Q.tolist()))
Mq = list(range(1, number_gap_q))
print("*********发电流量离散完成***********")
print('发电流量离散成了%d段' % (number_gap_q-1))
print(Q_gap)
print(Mq)
#水头区间发电函数离散
h_avg = [(H_gap[i]+H_gap[i+1])/2 for i in range(0, number_gap_h-1)]
print("*********%d条典型发电曲线离散完成***********" % (number_gap_h-1))
P_gap = {(m, n): N_HQ_fit(h_avg[m-1], Q_gap[n])for m in range(1, number_gap_h)for n in range(0, number_gap_q)}
print(P_gap)

# print(UPZ_UPV_fit_coe)
# coe_zv = UPZ_UPV_fit_coe
# v= 10281.0
# res = coe_zv[0]*v**4+coe_zv[1]*v**3+coe_zv[2]*v**2+coe_zv[3]*v**1+coe_zv[4]
# print(res)
# print(UPZ_UPV_fit(10281.0))
#
# print(DWZ_Q_fit_coe)
# coe_zq = DWZ_Q_fit_coe
# q= 52.0
# res = coe_zq[0]*q**2+coe_zq[1]*q+coe_zq[2]
# print(res)
# print(DWZ_Q_fit(52.0))




































