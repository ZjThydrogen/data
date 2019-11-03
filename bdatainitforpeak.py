import xlrd
import numpy as np

#电站个数
N = 13
index_plant = list(range(1, N+1))
print("*********电站编号生成***********")
print(index_plant)

#补偿电站生成
#index_splant = [8,9,10,11,12,13]#金沙江流域
index_splant = [1,2,3,4,5,6,7]#澜沧江流域
# index_splant = [4, 5]
# index_splant=index_plant
print("*********补偿电站编号生成***********")
print(index_splant)

#时段hour
T = 96
index_t = list(range(1,T+1))
print("*********时段编号生成***********")
print(index_t)

T_2 = 95
index_t2 = list(range(1,T_2+1))
print("*********时段编号生成***********")
print(index_t2)

#[电站,时间]索引
index_plant_t = []
for index_N in range(0, N):
    for index_T in range(0, T):
        res = (index_N+1, index_T+1)
        index_plant_t.append(res)
print("*********电站时段-索引生成***********")
print(index_plant_t)

#读取excel
excel = xlrd.open_workbook('D:\pyprject\PBUC_20190807\dataforpeakloadusingip\data.xlsx')
basedata_table = excel.sheets()[1]

plant_name = basedata_table.col_values(0, start_rowx=1, end_rowx=N+3)
print(plant_name)

up_plant = basedata_table.col_values(2, start_rowx=1, end_rowx=N+1)
for i in range(0,len(up_plant)):
    up_plant[i] = int(up_plant[i])

#上游电站
plant_up_plant = dict(zip(index_plant, up_plant))
print("*********电站-上游电站***********")
print(plant_up_plant)

#起始库容 e4m3
init_v = basedata_table.col_values(3, start_rowx=1, end_rowx=N+1)
plant_init_v = dict(zip(index_plant, init_v))
print("*********电站-起始库容***********")
print(plant_init_v)

#末库容e4m3
end_v = basedata_table.col_values(4, start_rowx=1, end_rowx=N+1)
plant_end_v = dict(zip(index_plant, end_v))
print("*********电站-末库容***********")
print(plant_end_v)

#库容上限e4m3
max_v = basedata_table.col_values(5, start_rowx=1, end_rowx=N+1)
plant_max_v = dict(zip(index_plant, max_v))
print("*********电站-库容上限***********")
print(plant_max_v)

#库容下限e4m3
min_v = basedata_table.col_values(6, start_rowx=1, end_rowx=N+1)
plant_min_v = dict(zip(index_plant, min_v))
print("*********电站-库容下限***********")
print(plant_min_v)

#出力上限 MW
max_n = basedata_table.col_values(7, start_rowx=1, end_rowx=N+1)#出力上限变为
plant_max_n = dict(zip(index_plant, max_n))
print("*********电站-出力上限***********")
print(plant_max_n)

#出力下限 MW
min_n = basedata_table.col_values(8, start_rowx=1, end_rowx=N+1)
plant_min_n = dict(zip(index_plant, min_n))
print("*********电站-出力下限***********")
print(plant_min_n)

#爬坡约束 MW
ramp_n = basedata_table.col_values(9, start_rowx=1, end_rowx=N+1)
plant_ramp_n = dict(zip(index_plant, ramp_n))
print("*********电站-爬坡约束***********")
print(plant_ramp_n)

#耗水率
water_con = basedata_table.col_values(10, start_rowx=1, end_rowx=N+1)
plant_water_con = dict(zip(index_plant, water_con))
print("*********电站-耗水率***********")
print(plant_water_con)

#滞时 时段数
wait_time = basedata_table.col_values(11, start_rowx=1, end_rowx=N+1)

for i in range(0, len(wait_time)):
    wait_time[i] = int(wait_time[i])

plant_wait_time = dict(zip(index_plant, wait_time))
print("*********电站-滞时***********")
print(plant_wait_time)

#发电流量上限 m3/s
max_q = basedata_table.col_values(12, start_rowx=1, end_rowx=N+1)
plant_max_q = dict(zip(index_plant, max_q))
print("*********电站-发电流量上限***********")
print(plant_max_q)

#发电流量下限 m3/s
min_q = basedata_table.col_values(13, start_rowx=1, end_rowx=N+1)
plant_min_q = dict(zip(index_plant, min_q))
print("*********电站-发电流量下限***********")
print(plant_min_q)

#出库流量上限 m3/s
max_o = basedata_table.col_values(14, start_rowx=1, end_rowx=N+1)
plant_max_o = dict(zip(index_plant, max_o))
print("*********电站-发电流量上限***********")
print(plant_max_o)

#出库流量下限 m3/s
min_o = basedata_table.col_values(15, start_rowx=1, end_rowx=N+1)
plant_min_o = dict(zip(index_plant, min_o))
print("*********电站-发电流量下限***********")
print(plant_min_o)


inflow_table = excel.sheets()[2]
#区间流量m3/s
inflow = []
for i in range(0, N):
    inflow_i = inflow_table.col_values(i, start_rowx=1, end_rowx=97)
    inflow.extend(inflow_i)
plant_t_inflow = dict(zip(index_plant_t, inflow))
print("*********电站-区间流量***********")
print(plant_t_inflow)

#昨日出库流量m3/s
inflow_yest = []
for i in range(0, N):
    inflow_i_yes = inflow_table.col_values(14+i, start_rowx=1, end_rowx=97)
    inflow_yest.extend(inflow_i_yes)
plant_t_inflow_yest = dict(zip(index_plant_t, inflow_yest))
print("*********电站-区间流量***********")
print(plant_t_inflow_yest)

load_table = excel.sheets()[3]
#系统负荷MW
load = load_table.col_values(3, start_rowx=1, end_rowx=97)
t_load = dict(zip(index_t, load))
print("*********系统负荷***********")
print(t_load)
windfore = load_table.col_values(1, start_rowx=1, end_rowx=97)
t_windfore = dict(zip(index_t, windfore))
print("*********风电预测值***********")
print(t_windfore)
solarfore = load_table.col_values(2, start_rowx=1, end_rowx=97)
t_solarfore = dict(zip(index_t, solarfore))
print("*********光伏预测值***********")
print(t_solarfore)

#60%的置信区间-6
# wind_l = load_table.col_values(10-6, start_rowx=1, end_rowx=97)
# wind_r = load_table.col_values(10+6, start_rowx=1, end_rowx=97)

newenergyload_table = excel.sheets()[4]
wind_l = newenergyload_table.col_values(0, start_rowx=1, end_rowx=97)
wind_r = newenergyload_table.col_values(1, start_rowx=1, end_rowx=97)
t_wind_l = dict(zip(index_t, wind_l))
t_wind_r = dict(zip(index_t, wind_r))
print("*********风电区间***********")
print(t_wind_l)
print(t_wind_r)

#60%的置信区间-6
# solar_l = load_table.col_values(31-6, start_rowx=1, end_rowx=97)
# solar_r = load_table.col_values(31+6, start_rowx=1, end_rowx=97)
solar_l = newenergyload_table.col_values(2, start_rowx=1, end_rowx=97)
solar_r = newenergyload_table.col_values(3, start_rowx=1, end_rowx=97)
t_solar_l = dict(zip(index_t, solar_l))
t_solar_r = dict(zip(index_t, solar_r))
print("*********光伏区间***********")
print(t_solar_l)
print(t_solar_r)
