# モジュール組込
import minimalmodbus    # Modbusモジュール組込
import serial           # シリアル通信モジュール組込
import csv              # CSVファイルモジュール組込
import datetime         # 時計モジュール組込
from time import sleep  # タイマー組込

# 通信設定
modbus=minimalmodbus.Instrument('COM5',1)
modbus.debug=False
modbus.handle_local_echo=False
close_port_after_echo_call=True
modbus.mode=minimalmodbus.MODE_RTU
modbus.serial.baudrate=9600
modbus.serial.bytesize=8
modbus.serial.parity=serial.PARITY_NONE
modbus.serial.stopbits=1
modbus.serial.timeout=1
# 設定パラメーターセット
# ----------設定
sys_cnt=[10,13]
sys_add=[0x2000,0x2200]
sys_name=[["相線式","ID","CT","CT比","ローカット電流"
           ,"簡易計測","簡易計測電圧","簡易計測力率"
           ,"パルス端子割当","電圧割当"],
         ["プロトコル","通信速度","データ長","ストップビット"
           ,"パリティ","送信待ち時間","パルス出力単位","VT比"
           ,"換算レート","換算単位","LCD消灯時間","表示桁固定"
           ,"警告ON/OFF"]]
sys_max=[[4,99,5,99999,199,1,99999,100,2,2],
           [3,5,1,1,2,99,7,99999,99999,2,3,2,1]]
sys_type=[[10,0,11,1,1,7,1,2,12,13],
         [14,15,8,9,16,0,17,2,3,4,18,19,7]]
# ----------設定値リスト
sys_list=[["OFF","ON"],["7bit","8bit"],["1bit","2bit"]
          ,["1P2W","1P3W","3P3W","1P2W2","1P3W2"]
          ,["5A","50A","100A","225A","400A","600A"]
          ,["OFF","OUT1","OUT2"],["R-N","T-N","R-T"]
          ,["CompoWay/F","Modbus","BACnetMS/TP","KM20"]
          ,["1.2kbps","2.4kbps","4.8kbps","9.6kbps"
            ,"19.2kbps","38.4kbps"],["なし","奇数","偶数"]
          ,["1Wh","10Wh","100Wh","1kWh","5kWh","10kWh","50kWh","100kWh"]
          ,["OFF","1分","5分","10分"],["OFF","kWh固定","MWh固定"]]
# ----------計測パラメーター
ctrl_cnt=[10,5,5]
ctrl_add=[0x0000,0x0200,0x220]
ctrl_name=[["電圧1","電圧2","電圧3","電流1","電流2","電流3"
            ,"力率","周波数","有効電力","無効電力"],
           ["積算有効電力量","積算回生電力量","積算進み無効電力量"
            ,"積算遅れ無効電力量","積算総合無効電力量"],
           ["積算有効電力量","積算回生電力量","積算進み無効電力量"
            ,"積算遅れ無効電力量","積算総合無効電力量"]]
ctrl_unit=[["V","V","V","A","A","A","","Hz","W","Var"]
           ,["Wh","Wh","Varh","Varh","Varh"]
           ,["kWh","kWh","kVarh","kVarh","kVarh"]]
ctrl_type=[[1,1,1,3,3,3,2,1,1,1],[0,0,0,0,0],[0,0,0,0,0]]

def timer_sub(sel_timer):   # タイマーセット
    if sel_timer=="m":sel=1 # １分間
    if sel_timer=="h":sel=2 # １時間
    if sel_timer=="d":sel=3 # １日間
    set_timer=[dt_now.second,dt_now.minute,dt_now.hour,dt_now.day]
    return set_timer[sel],set_timer[sel-1]

def read_data(date_data,par,sl_cnt,sl_add,sl_type):
    read_list=[]
    for a in range(2): # ID loop
        id=a+1
        data_list1=[date_data,id]
        data_list2=[date_data,id]
        modbus=minimalmodbus.Instrument('COM5',id)
        for sys in range(par): # parameter loop
            data_type=sl_type[sys]
            for add in range(sl_cnt[sys]):
                reg_add=sl_add[sys]+add*2
                data=modbus.read_long(registeraddress=reg_add,
                                        functioncode=3,
                                        number_of_registers=2,
                                        signed=True)
                if data_type[add]<7: # 単位判定
                    chk_list=[data,data/10,data/100,data/1000,str(data)]
                    chk_data=chk_list[data_type[add]]
                else: # 設定値判定
                    chk_list=sys_list[data_type[add]-7]
                    chk_data=chk_list[data]
                data_list1.append(data)
                data_list2.append(chk_data)
        read_list.append(data_list2)
    return read_list

# 設定内容読込
file_name='omron_setup'
dt_now = datetime.datetime.now() # 日時を取得
file_time=dt_now.strftime('_20%y_%m_%d_%H%M')
file_name=file_name+file_time+'.csv'
print("ファイル名:",file_name)
with open(file_name, 'w', newline='') as file: # CSVファイルを作成
    writer = csv.writer(file)
    name_data=["","ID"]+sys_name[0]+sys_name[1]
    max_data=["",""]+sys_max[0]+sys_max[1]
    writer.writerow(name_data) # ヘッダー１を書込
    writer.writerow(max_data) # ヘッダー２を書込
    read_list=read_data("",2,sys_cnt,sys_add,sys_type)
    print("設定データ",read_list[0])
    print("計測データ",read_list[1])
    writer.writerow(read_list[0])
    writer.writerow(read_list[1])
    
file_name='omron_running'
file_time=dt_now.strftime('_20%y_%m_%d_%H%M')
file_name=file_name+file_time+'.csv'
print("ファイル名:",file_name)
with open(file_name, 'w', newline='') as file: # CSVファイルを作成
    writer = csv.writer(file)
    name_data=["日時","ID"]+ctrl_name[0]+ctrl_name[1]+ctrl_name[2]
    unit_data=["年月日時分秒",""]+ctrl_unit[0]+ctrl_unit[1]+ctrl_unit[2]
    writer.writerow(name_data) # ヘッダー１を書込
    writer.writerow(unit_data) # ヘッダー２を書込

    sel_timer="m" # ループタイマー:m=1分,h=1時間,d=1日
    interval1=8.7 # 計測間隔（秒）+1.3秒:8.7->10秒
    interval2=1
    set_timer1,set_timer2=timer_sub(sel_timer)
    loop_timer1=set_timer1
    loop_timer2=set_timer2
    
    while loop_timer1<=set_timer1: # ループ設定
        dt_now = datetime.datetime.now() # 日時を取得
        date_time=dt_now.strftime('%y/%m/%d %H:%M:%S')
        read_list=read_data(date_time,3,ctrl_cnt,ctrl_add,ctrl_type)
        print(read_list[0])
        print(read_list[1])
        writer.writerow(read_list[0])
        writer.writerow(read_list[1])
        set_timer1,set_timer2=timer_sub(sel_timer) 
        #print(loop_timer1,set_timer1,loop_timer2,set_timer2)
        if loop_timer1<set_timer1: # ループ終了判定
            if loop_timer2<set_timer2:
                break
            if set_timer1-loop_timer1>interval2:
                break
        sleep(interval1)
        

