# ----------初期設定
from pymodbus.client import ModbusSerialClient as ModbusClient  # Modbus組込
import csv                                                      # CSVファイルモジュール組込
import datetime                                                 # 時計モジュール組込
from time import sleep                                          # タイマー組込
import tkinter as tk                                            # GUIモジュール組込
import threading                                                # スレッド組込

interval=5.9                                                      # 待ち時間（秒）
# ----------パラメーター設定 [address,byte,name,type,unit,data...]
para_data=[[0x0213,1,"系統電圧",1,"V(AC)"],[0x0214,1,"系統電流",1,"A(AC)"],
           [0x0215,1,"系統周波数",2,"Hz"],[0x021e,1,"系統充電電流",1,"A(AC)"],
           
           [0x0210,1,"機器状態",8,"","起動","待機","初期化","省電力","商用出力",
            "インバーター出力","系統出力","混合出力","-","-","停止","故障"],
           [0x0216,1,"出力電圧",1,"V(AC)"],[0x0219,1,"出力電流",1,"A(AC)"],
           [0x0218,1,"出力周波数",2,"Hz"],[0x021b,1,"負荷有効電力",0,"W"],
           [0x021c,1,"負荷皮相電力",0,"W"],[0x021f,1,"負荷率",0,"%"],
           
           [0x010b,1,"充電状態",8,"","未充電","定電流(CC)充電","定電圧(CV)充電","-",
            "浮遊充電","-","充電中1","充電中2"],[0x0100,1,"蓄電池SOC",0,"%"],
           [0x0101,1,"蓄電池電圧",1,"V(DC)"],[0x0102,1,"蓄電池電流",1,"A(DC)"],
           [0x010e,1,"充電電力",0,"W"],[0x0217,1,"INV電流",1,"A(DC)"],
           
           [0x0107,1,"PV入力電圧",1,"V(DC)"],[0x0108,1,"PV入力電流",1,"A(DC)"],
           [0x0109,1,"PV入力電力",0,"W"],[0x0224,1,"PV降圧電流",1,"A(DC)"],
           
           [0x0212,1,"DCバス電圧",1,"V(DC)"],[0x0225,1,"降圧電流",1,"A(DC)"],
           [0x0220,1,"PVHT温度",1,"℃"],[0x0221,1,"INVHT温度",1,"℃"],
           [0x0222,1,"Tr温度",1,"℃"],[0x0223,1,"内部温度",1,"℃"],
           
           [0xf02d,1,"本日充電量",0,"Ah"],[0xf02e,1,"本日放電量",0,"Ah"],
           [0xf02f,1,"本日発電量",1,"kWh"],[0xf030,1,"本日消費量",1,"kWh"],
           [0xf03c,1,"本日商用充電量",0,"Ah"],[0xf03d,1,"本日商用電力消費量",1,"kWh"],
           [0xf03e,1,"本日インバーター稼働時間",0,"時間"],[0xf03f,1,"本日バイパス稼働時間",0,"時間"],

           [0xf034,2,"累積充電量",0,"Ah"],[0xf036,2,"累積放電量",0,"Ah"],
           [0xf038,2,"累積発電量",0,"kWh"],[0xf03a,2,"累積負荷積算電力量",0,"kWh"],
           [0xf046,2,"累積商用充電量",0,"kWh"],[0xf048,2,"累積商用負荷電力消費量",0,"kWh"],
           [0xf04a,1,"累積インバーター稼働時間",0,"時間"],[0xf04b,1,"累積バイパス稼働時間",0,"時間"]]
para_xy=[[0,1],[0,2],[0,3],[0,4],
         [0,6],[0,7],[0,8],[0,9],[0,10],[0,11],[0,12],
         
         [0,14],[0,15],[0,16],[0,17],[0,18],[0,19],
         [0,21],[0,22],[0,23],[0,24],
         [0,26],[0,27],[0,28],[0,29],[0,30],[0,31],
         
         [5,1],[5,2],[5,3],[5,4],[5,5],[5,6],[5,7],[5,8],
         [5,10],[5,11],[5,12],[5,13],[5,14],[5,15],[5,16],[5,17]]
para_list=43
m1_data=[["AC入力側",0,1],["AC出力側",0,6],["総合データ",10,1],
          ["蓄電池側",0,14],["PV入力側",0,21],["内部システム",0,26],
          ["当日積算データ",5,1],["累積積算データ",5,10]]
m1_list=8
m2_data=[["蓄電池総合電流",10,9,1,1,"A"],["蓄電池総合電力",10,10,2,1,"kW"],
         ["PV発電総合電力",10,11,3,3,"kW"],["AC入力総合電力",10,12,4,3,"kVA"],
         ["AC出力総合電力",10,13,5,3,"kW"],["機器消費電力",10,14,6,0,"W"],
         
         ["総合充電量",10,2,7,0,"Ah"],["総合放電量",10,3,8,0,"Ah"],
         ["総合発電量",10,4,9,1,"kWh"],["総合消費量",10,5,10,1,"kWh"],
         ["総合商用充電量",10,6,11,0,"Ah"],["総合商用消費量",10,7,12,1,"kWh"]]
m2_list=12

pra_type=[
          ["ユーザー設定","密閉型鉛","開放型鉛","ゲル型鉛","LFPx14","LEPx15","LFPx16"
           ,"LFPx7","LFPx8","LFPx8","NCAx7","NCAx8","NCAx13","NCAx14"],
          ["OFF","BMS制御","INV制御"],
          ["OFF","AC入力PV出力（逆潮流）","AC出力PV出力（順潮流）"],
          ["単相","並列","2P0","2P1","2P2","3P1","3P2","3P3"],
          ["PV優先","系統優先","蓄電池優先"],
          ["APL(90-280V)","UPS(90-140V)"],
          ["PV優先","系統優先","ハイブリッド","PV専用"],
          ["OFF","RS485-BMS","CAN-BMS"],
          ["Pace","Rata","Allgrand","Oliter","PCT","Sunwoda","Dyness","SRNE"
           ,"Pylontech","","","","","","","","WS Technicals","Uz Energy"],
          ["0°","120°","180°"]]
sys_name=[
           ["発電量1日前","発電量2日前","発電量3日前","発電量4日前"
           ,"発電量5日前","発電量6日前","発電量7日前"
           ,"充電量1日前","充電量2日前","充電量3日前","充電量4日前"
           ,"充電量5日前","充電量6日前","充電量7日前"
           ,"放電量1日前","放電量2日前","放電量3日前","放電量4日前"
           ,"放電量5日前","放電量6日前","放電量7日前"
           ,"商用充電量1日前","商用充電量2日前","商用充電量3日前","商用充電量4日前"
           ,"商用充電量5日前","商用充電量6日前","商用充電量7日前"],
           ["消費電力量1日前","消費電力量2日前","消費電力量3日前","消費電力量4日前"
           ,"消費電力量5日前","消費電力量6日前","消費電力量7日前"
           ,"商用消費量1日前","商用消費量2日前","商用消費量3日前","商用消費量4日前"
           ,"商用消費量5日前","商用消費量6日前","商用消費量7日前"]]
sys_unit=[
           ["Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah"
           ,"Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah","Ah"],
           ["kWh","kWh","kWh","kWh","kWh","kWh","kWh","kWh","kWh","kWh","kWh","kWh","kWh","kWh"]]

prt_name=[["PV充電最大電流","蓄電池容量","システム電圧","蓄電池タイプ","過電圧停止電圧"
           ,"充電上限電圧","均等充電電圧","CC充電電圧","CV充電電圧","充電再開電圧"
           ,"低電圧復帰電圧","低電圧警告電圧","過放電遅延オフ電圧","放電停止電圧","放電停止SOC"
           ,"過放電遅延オフ時間","均等充電時間","CC充電遅延時間","均等充電間隔","温度補償係数"
           ,"充電上限温度","充電下限温度","放電上限温度","放電下限温度","加熱開始温度"
           ,"加熱停止温度","バイパス切替電圧","充電停止電流","充電停止SOC","過放電警報SOC"
           ,"バイパス切替SOC"],
           ["INV切替SOC","","INV切替電圧","均等充電遅延","リチウム起動電流","充電制御"
           ,"充電開始時間１","充電終了時間１","充電開始時間２","充電終了時間２","充電開始時間３"
           ,"充電終了時間３","充電時間設定","放電開始時間１","放電終了時間１","放電開始時間２"
           ,"放電終了時間２","放電開始時間３","放電終了時間３","放電時間設定",""
           ,"","","PV系統出力","漏電検知機能",""],
           ["","","","","","","","","","","","","","","","","","",""
           ,"","","","","","定格出力","","","","","","","PV最大電圧"],
           ["PV最大充電電流","","","","","","","","","最大出力電流","","","","","","","",""],
           ["RS485アドレス","並列運転モード","","","出力優先","AC充電最大電流","均等充電有無"
           ,"N相出力","AC出力電圧","AC出力周波数","最大充電電流","AC入力電圧範囲"
           ,"省エネモード","過負荷停止再起動","高温停止再起動","充電モード","警報音","電源切替警報"
           ,"バイパス出力","障害記録","分相変圧器","BMS通信","","","","","","BMSプロトコル"]]
prt_unit=[["A","Ah","V","","V","V","V","V","V","V","V","V","V","V","%","s","min"
           ,"min","day","","℃","℃","℃","℃","℃","℃","V","A","%","%","%"],
           ["%","","V","min","A","","時分","時分","時分","時分","時分","時分" ,""
           ,"時分","時分","時分","時分","時分","時分","","","","","","",""],
           ["","","","","","","","","","","","","","","","","","",""
           ,"","","","","","kW","","","","","","","V"],
           ["A","","","","","","","","","A","","","","","","","",""],
           ["ID","","","","","A","","未実装","V","Hz","A","","",""
           ,"","","","","","","","","","","","","",""]]
prt_type=[[1,0,0,10,4,4,4,4,4,4,4,4,4,4,0,0,0,0,0,0,0,0,0,0,0,0,4,1,0,0,0],
          [0,0,4,0,1,11,7,7,7,7,7,7,6,7,7,7,7,7,7,6,0,0,0,0,12,6],
          [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0],
          [1,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0],
          [0,13,0,0,14,1,6,0,1,2,1,15,6,6,6,16,6,6,6,6,6,17,0,0,0,0,0,18]]

# ----------Modbus接続設定
client1=ModbusClient(framer="rtu",port="COM5",                  # USBポート5（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=1)
client2=ModbusClient(framer="rtu",port="COM6",                  # USBポート6（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=1)

# ----------Modbusデータ読出
def modbus_read(slave_add,slave_count,slave_id):                # ファンクション03h
#    client1.connect()                                           # ポート接続
#    client2.connect()
#    read_data1=client1.read_holding_registers(
#        address=slave_add,count=slave_count,slave=slave_id)
#    read_data2=client2.read_holding_registers(
#        address=slave_add,count=slave_count,slave=slave_id+1)
#    client1.close()                                             # ポート切断
#    client2.close()
#    return read_data1.registers,read_data2.registers
    test_data=[ [0x010b,[1],[2]],[0x0100,[53],[54]],[0x0101,[524],[525]],[0x0102,[10],[10]],[0x010e,[241],[242]],
                [0x0107,[1727],[1728]],[0x0108,[14],[15]],[0x0109,[241],[242]],[0x0224,[4],[5]],
                [0x0210,[5],[3]],[0x0216,[999],[1000]],[0x0217,[23],[24]],[0x0218,[5999],[6000]],[0x0219,[25],[25]],
                [0x021b,[741],[742]],[0x021c,[761],[762]],[0x021f,[5],[6]],[0x0225,[20],[21]],
                [0x0212,[3930],[3931]],[0x0213,[0],[1000]],[0x0214,[0],[500]],[0x0215,[0],[6001]],[0x021e,[0],[30]],
                [0x0220,[410],[411]],[0x0221,[404],[405]],[0x0222,[537],[538]],[0x0223,[549],[560]],
                [0xf02d,[28],[29]],[0xf02e,[11],[12]],[0xf02f,[17],[18]],[0xf030,[16],[17]],
                [0xf03c,[30],[31]],[0xf03d,[16],[17]],[0xf03e,[6],[7]],[0xf03f,[0],[0]],
                [0xf034,[4079,0],[4080,0]],[0xf036,[2488,0],[2499,0]],[0xf038,[3061,0],[3090,0]],[0xf03a,[2022,0],[2080,0]],
                [0xf046,[991,0],[1000,0]],[0xf048,[825,0],[900,0]],[0xf04a,[752],[753]],[0xf04b,[83],[84]]]
    test_add=[0xf034,0xf036,0xf038,0xf03a,0xf046,0xf048,]
    for v in range(para_list):
        
        if slave_add==test_data[v][0]:
            read_data1=test_data[v][1]
            read_data2=test_data[v][2]
            
        if slave_add==0xe003:
            read_data1=[48]
            read_data2=read_data1
        if slave_add==0x0035:
                read_data1=[50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50]
                read_data2=read_data1
    return read_data1,read_data2

# ----------CSVファイル設定
slave_id=1
sys_volt,sys_volt2=modbus_read(0xe003,1,slave_id)                         # システム電圧読込
sys_id,sys_id2=modbus_read(0x0035,20,slave_id)                          # プロダクトID読込                                                      # 機器ID
dt_now = datetime.datetime.now()                                # 日時を取得
file_time=dt_now.strftime('_20%y_%m_%d_%H%M')
id_name=''
for a in range(20):id_name=id_name+chr(sys_id[a])
#file_name1=id_name+'_setting.csv'                               # 設定ファイル名作成
file_name2=id_name+file_time+'.csv'                             # ログファイル名作成
#print("設定ファイル名:",file_name1)
print("ログファイル名:",file_name2)
name_data1=[file_time]
unit_data1=[""]
for i in range(para_list):
    name_data1.append(para_data[i][2])
    unit_data1.append(para_data[i][4])

# ----------データ変換
def change_minus(rdd):return rdd-65536 if rdd>32767 else rdd
def change_type(data_type,rdd,sys_volt,d_type,d_add):                 # データタイプ変換
    if data_type<5:
        if data_type<4:                                         # 1/10,1/100,1/1000判定
            if d_type==1 :                                      # 16bit
                rdd=change_minus(rdd)
            if d_type==2 :                                      # 32byte
                if rdd>2147483647 :rdd=rdd-4294967295
            data_list=[rdd,rdd/10,rdd/100,rdd/1000]
            d=data_list[data_type]
        if data_type==4:                                        # 電圧データ判定24V/48V
            d=rdd/10
            if sys_volt[0]==24:d=d*2  
            if sys_volt[0]==48:d=d*4   
        d=str(d)
    if data_type==5:d=chr(rdd)                                  # 文字データ判定　
    if data_type==6:d=data_list["OFF","ON"][rdd]                # ON/OFF判定  
    if data_type==7:                                            # 時刻判定
        ym=hex(rdd)[2:].zfill(4)
        d=str(int(ym[0:2], 16)).zfill(2)+":"+str(int(ym[2:4], 16)).zfill(2)
    if data_type>=8:
        d=para_data[d_add][rdd+5]                          # 拡張判定
    return d

def data_read():
    dt_now=datetime.datetime.now()                              # 日時を取得
    date_time=dt_now.strftime('%y/%m/%d %H:%M:%S')
    p_type=0
    writer_data1=[]
    writer_data2=[]
    reader_data1=[]
    reader_data2=[]
    for a in range(para_list):
        d_type=para_data[a][1]
        data_type=para_data[a][3]
        read_data1,read_data2=modbus_read(para_data[a][0],d_type,slave_id)
        if d_type==1:
            data1=change_type(data_type,read_data1[0],sys_volt[0],d_type,a)# 16bitデータ変換
            data2=change_type(data_type,read_data2[0],sys_volt[0],d_type,a)
        else:
            data1=change_type(data_type,                           # 32bitデータ変換
                int(str(int(hex(read_data1[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(read_data1[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt,d_type,a)
            data2=change_type(data_type,
                int(str(int(hex(read_data2[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(read_data2[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt,d_type,a)
        writer_data1.append(data1)
        writer_data2.append(data2)
        reader_data1.append(read_data1)
        reader_data2.append(read_data2)
    #if p_type==0:prt="logging:"                  
    #else:prt="setting:"
    csv_data1=[date_time]
    csv_data2=[date_time]
    for i in range(para_list):
        csv_data1.append(writer_data1[i])
        csv_data2.append(writer_data2[i])
    writer.writerow(csv_data1)                                  # CSVデータ書込
    writer.writerow(csv_data2)
    orign_data=[]
    q=reader_data1[14][0]
    r=reader_data2[14][0]
    if q>32767 :q=q-65536
    if r>32767 :r=r-65536
    orign_data.append((q+r)/10)
    calc_list=[round(((reader_data1[13][0]*q+reader_data2[13][0]*r)/100)/1000,3),
               (reader_data1[19][0]+reader_data2[19][0])/1000,
               ((reader_data1[0][0]*reader_data1[1][0])+(reader_data2[0][0]*reader_data2[1][0]))/100000,
               (reader_data1[8][0]+reader_data2[8][0])/1000,
               round(round(((reader_data1[13][0]*q+reader_data2[13][0]*r)/100)/1000,3)
               +((reader_data1[8][0]+reader_data2[8][0])/1000
                 -reader_data1[19][0]+reader_data2[19][0])/1000
               -((reader_data1[0][0]*reader_data1[1][0])+(reader_data2[0][0]*reader_data2[1][0]))/100000,3),
    
               reader_data1[27][0]+reader_data2[27][0],
               reader_data1[28][0]+reader_data2[28][0],
               (reader_data1[29][0]+reader_data2[29][0])/10,
               (reader_data1[30][0]+reader_data2[30][0])/10,
               reader_data1[31][0]+reader_data2[31][0],
               (reader_data1[32][0]+reader_data2[32][0])/10]
    for s in range(m2_list-1):orign_data.append(calc_list[s])
    #print(prt,date_time)      
    return date_time,writer_data1,writer_data2,orign_data

# ----------モニター画面
def create_gui():                                               # GUI作成
    root=tk.Tk()
    root.geometry('1100x720')                                   # ウインドウサイズ
    root.title("HYP4850U100-H")                                 # ウインドウタイトル
    frame=tk.Frame(root)
    frame.grid(row=0,column=0,sticky=tk.NSEW,padx=5,pady=10)
    wid=[19,13,5,13,5]
    label=tk.Label(frame,width=wid[4],text="日時",anchor=tk.W)
    label.grid(column=4,row=0)
    label0=tk.Label(frame,width=wid[0],text=date_time,anchor=tk.W,borderwidth=1)
    label0.grid(column=5,row=0)
    for j in range(2):
        labels=[tk.Label(frame,width=wid[i],text=["","ID1","","ID2",""][i]
                        ,anchor=tk.W)for i in range(5)]
        [labels[h].grid(column=h+j*5,row=1)for h in range(5)]
    col1='#0000ff'                                              #データ文字色
    col2='#cccccc'                                              #データ背景色
    labels=[tk.Label(frame,width=wid[0],text=para_data[i][2]
                     ,anchor=tk.W)for i in range(para_list)]
    [labels[h].grid(column=para_xy[h][0],row=para_xy[h][1]+1)for h in range(para_list)]
    labels1=[tk.Label(frame,width=wid[1],text=writer_data1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(para_list)]
    [labels1[h].grid(column=para_xy[h][0]+1,row=para_xy[h][1]+1)for h in range(para_list)]
    labels2=[tk.Label(frame,width=wid[3],text=writer_data1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(para_list)]
    [labels2[h].grid(column=para_xy[h][0]+3,row=para_xy[h][1]+1)for h in range(para_list)]
    for j in range(2):
        labels=[tk.Label(frame,width=wid[4],text=para_data[i][4]
                        ,anchor=tk.W)for i in range(para_list)]
        [labels[h].grid(column=para_xy[h][0]+2+2*j,row=para_xy[h][1]+1)for h in range(para_list)]
    labels=[tk.Label(frame,width=wid[0],text=m1_data[i][0],font=("Atari",9,"bold")
                        ,anchor=tk.W)for i in range(m1_list)]
    [labels[h].grid(column=m1_data[h][1],row=m1_data[h][2])for h in range(m1_list)]
    labels=[tk.Label(frame,width=wid[0],text=m2_data[i][0]
                        ,anchor=tk.W)for i in range(m2_list)]
    [labels[h].grid(column=m2_data[h][1],row=m2_data[h][2])for h in range(m2_list)]
    
    labels3=[tk.Label(frame,width=wid[1],text=orign_data[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(m2_list)]
    [labels3[h].grid(column=m2_data[h][1]+1,row=m2_data[h][2])for h in range(m2_list)]
    
    labels=[tk.Label(frame,width=wid[4],text=m2_data[i][5]
                        ,anchor=tk.W)for i in range(m2_list)]
    [labels[h].grid(column=m2_data[h][1]+2,row=m2_data[h][2])for h in range(m2_list)]
    
    button=tk.Button(frame,text="計測終了",command=root.destroy) # ループ終了
    button.grid(column=0,row=0)

    thread=threading.Thread(target=update_data,args=([label0,
                                    [labels1[x]for x in range(para_list)],
                                    [labels2[x]for x in range(para_list)],
                                    [labels3[x]for x in range(m2_list)]]))
    thread.daemon=True                                          # メインウィンドウが閉じたらスレッドも終了
    thread.start()                                              # スレッド処理開始
    root.mainloop()                                             # メインループ開始

# ----------データ更新処理
def update_data(label0,labels1,labels2,labels3):
        while True:
            date_time,writer_data1,writer_data2,orign_data=data_read()
            label0.config(text=f"{date_time}")                  # GUIデータ更新
            [labels1[x].config(text=f"{writer_data1[x]}")for x in range(para_list)]
            [labels2[x].config(text=f"{writer_data2[x]}")for x in range(para_list)]
            [labels3[x].config(text=f"{orign_data[x]}")for x in range(m2_list)]
            sleep(interval)                                     # インターバルタイマー

# ----------実行
with open(file_name2,'w', newline='') as file:                  # CSVファイルオープン
        writer=csv.writer(file)
        w_data=[name_data1,unit_data1]
        [writer.writerow(w_data[i])for i in range(2)]           # ヘッダー書込
        date_time,writer_data1,writer_data2,orign_data=data_read()
        if __name__ == "__main__":
            create_gui()
            
# 終了
