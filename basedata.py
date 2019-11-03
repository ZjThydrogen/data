import xlrd
import itertools


#电站个数
N = 1
NN = list(range(1, N+1))
print("*********电站索引生成***********")
print(NN)

#机组总个数
I = 2

#时段hour
T = 24
K = 4
NT = list(range(1, T+1))
NK = list(range(1, T*K+1))
NRTK = list(range(1, K+1))


#[(1,1),(1,2),(1,3),(1,4),...,(24,4)]
index_NK = []
for index_T in range(0, T):
    for index_K in range(0, K):
        res = (index_T+1, index_K+1)
        index_NK.append(res)
print("*********时间类索引生成***********")
print(NT)
print(NRTK)

excel = xlrd.open_workbook('D:\pyprject\PBUC_20190807\data.xlsx')

#-------------------------读取电站表--------------------------------------
pd_table = excel.sheets()[0]
end_col = N+1

#库水位上下限m
zu_min = pd_table.row_values(1, start_colx=1, end_colx=end_col)
zu_max = pd_table.row_values(2, start_colx=1, end_colx=end_col)
#电站编码作为索引
ZU_min = dict(zip(NN, zu_min))
ZU_max = dict(zip(NN, zu_max))
print("*********库水位上下限获得***********")
print(ZU_min)
print(ZU_max)
#库容上下限万m3
v_min = pd_table.row_values(3, start_colx=1, end_colx=end_col)
v_max = pd_table.row_values(4, start_colx=1, end_colx=end_col)

V_min = dict(zip(NN, v_min))
V_max = dict(zip(NN, v_max))
print("*********库容上下限获得***********")
print(V_min)
print(V_max)
#泄量上下限m3/s
q_min = pd_table.row_values(5, start_colx=1, end_colx=end_col)
q_max = pd_table.row_values(6, start_colx=1, end_colx=end_col)
Q_min = dict(zip(NN, q_min))
Q_max = dict(zip(NN, q_max))
print("*********出库流量上下限获得***********")
print(Q_min)
print(Q_max)
#弃水流量上下限m3/s
s_max = pd_table.row_values(7, start_colx=1, end_colx=end_col)
S_max = dict(zip(NN, s_max))
print("*********弃水流量上限获得***********")
print(S_max)
#机组台数
n_unit = pd_table.row_values(8, start_colx=1, end_colx=end_col)
#生成电站-机组索引
index_plant_unit = []
for index_plant in range(0, N):
    unit = int(n_unit[index_plant])
    for index_unit in range(0, unit):
        res = (index_plant+1, index_unit+1)
        index_plant_unit.append(res)
#生成电站包含机组的二维列表
plant_unit = {}
for index_plant in range(0, N):
    unit = int(n_unit[index_plant])
    u = list(range(1, unit+1))
    plant_unit[1] = u
print("*********电站-机组索引获得***********")
print(index_plant_unit)
print(plant_unit)

v_start = pd_table.row_values(9, start_colx=1, end_colx=end_col)
V_start = dict(zip(NN, v_start))
print("*********初始库容获得***********")
print(V_start)

#-------------------------读取机组表--------------------------------------
ud_table = excel.sheets()[1]
end_col = I+1

#机组出力上下限
ph_min = ud_table.row_values(1, start_colx=1, end_colx=end_col)
ph_max = ud_table.row_values(2, start_colx=1, end_colx=end_col)
#索引为电站-机组
PH_min = dict(zip(index_plant_unit, ph_min))
PH_max = dict(zip(index_plant_unit, ph_max))
print("*********机组出力上下限获得***********")
print(PH_min)
print(PH_max)

#机组发电流量上下限
qunit_min = ud_table.row_values(3, start_colx=1, end_colx=end_col)
qunit_max = ud_table.row_values(4, start_colx=1, end_colx=end_col)
QUNIT_min = dict(zip(index_plant_unit, qunit_min))
QUNIT_max = dict(zip(index_plant_unit, qunit_max))
print("*********机组发电流量上下限获得***********")
print(QUNIT_min)
print(QUNIT_max)

#-------------------------读取时间序列表--------------------------------------
td_table = excel.sheets()[2]
end_row = T+1

#风电区间
pw_L = td_table.col_values(1, start_rowx=1, end_rowx=end_row)
pw_R = td_table.col_values(2, start_rowx=1, end_rowx=end_row)
PW_L = dict(zip(NT, pw_L))
PW_R = dict(zip(NT, pw_R))
print("*********风电区间获得***********")
print(PW_L)
print(PW_R)

#光伏区间
ps_L = td_table.col_values(3, start_rowx=1, end_rowx=end_row)
ps_R = td_table.col_values(4, start_rowx=1, end_rowx=end_row)
PS_L = dict(zip(NT, ps_L))
PS_R = dict(zip(NT, ps_R))
print("*********光伏区间获得***********")
print(PS_L)
print(PS_R)
#日前电价区间
da_price_L = td_table.col_values(5, start_rowx=1, end_rowx=end_row)
da_price_R = td_table.col_values(6, start_rowx=1, end_rowx=end_row)
DA_price_L = dict(zip(NT, da_price_L))
DA_price_R = dict(zip(NT, da_price_R))
print("*********日前电价区间获得***********")
print(DA_price_L)
print(DA_price_R)

#日内实时电价
end_row = T*K+1
rt_price_L = td_table.col_values(8, start_rowx=1, end_rowx=end_row)
rt_price_R = td_table.col_values(9, start_rowx=1, end_rowx=end_row)
# RT_price_L = dict(zip(NK, rt_price_L))
# RT_price_R = dict(zip(NK, rt_price_R))
RT_price_L = dict(zip(index_NK, rt_price_L))
RT_price_R = dict(zip(index_NK, rt_price_R))
print("*********实时电价区间获得***********")
print(RT_price_L)
print(RT_price_R)

#平衡电价
BA_price = td_table.cell_value(1, 10)
print('平衡电价：%d' % BA_price)
#区间流量
inflow = []
for i in range(11, 11+N):
    values = td_table.col_values(i, start_rowx=1, end_rowx=T+1)
    inflow += values
#生成电站-时间索引
index_plant_t = []
for index_plant in range(0, N):
    for index_t in range(0, T):
        res = (index_plant+1, index_t+1)
        index_plant_t.append(res)
INFLOW = dict(zip(index_plant_t, inflow))
print("*********区间流量获得***********")
print(INFLOW)
print("-----------------------------------------------基础数据更新完毕-------------------------------------------------")











