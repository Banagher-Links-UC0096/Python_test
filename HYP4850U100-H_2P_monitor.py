# ----------初期設定・・・・HYP4850U100-H　2並列単相3線仕様専用
from pymodbus.client import ModbusSerialClient as ModbusClient  # Modbus組込
import csv                                                      # CSVファイルモジュール組込
import datetime                                                 # 時計モジュール組込
from time import sleep                                          # タイマー組込
import tkinter as tk                                            # GUIモジュール組込
import threading                                                # スレッド組込

interval=5                                                      # 待ち時間（秒）+5秒程度
# ----------パラメーター設定 [address,byte,name,type,unit,data...]
para_data=[[0x0220,1,"PVHT温度",1,"℃"],[0x0221,1,"INVHT温度",1,"℃"],
           [0x0222,1,"Tr温度",1,"℃"],[0x0223,1,"内部温度",1,"℃"],
           
           [0x0210,1,"機器状態",8,"","起動","待機","初期化","省電力","商用出力",
            "インバーター出力","系統出力","混合出力","-","-","停止","故障"],
           [0x0216,1,"出力電圧",1,"V(AC)"],[0x0219,1,"出力電流",1,"A(AC)"],
           [0x0218,1,"出力周波数",2,"Hz"],[0x021b,1,"負荷有効電力",0,"W"],
           [0x021c,1,"負荷皮相電力",0,"W"],[0x021f,1,"負荷率",0,"%"],
           [0x0225,1,"降圧電流",1,"A(AC)"],
           
           [0x010b,1,"充電状態",8,"","未充電","定電流充電","定電圧充電","-",
            "浮遊充電","-","充電中1","充電中2"],[0x0100,1,"蓄電池SOC",0,"%"],
           [0x0101,1,"蓄電池電圧",1,"V(DC)"],[0x0102,1,"蓄電池電流",1,"A(DC)"],
           [0x010e,1,"充電電力",0,"W"],[0x0217,1,"INV電流",1,"A(DC)"],
           
           [0x0107,1,"PV電圧",1,"V(DC)"],[0x0108,1,"PV電流",1,"A(DC)"],
           [0x0109,1,"PV電力",0,"W"],[0x0224,1,"PV降圧電流",1,"A(DC)"],
           
           [0x0212,1,"DCバス電圧",1,"V(DC)"],[0x0213,1,"系統電圧",1,"V(AC)"],
           [0x0214,1,"系統電流",1,"A(AC)"],[0x0215,1,"系統周波数",2,"Hz"],
           [0x021e,1,"系統充電電流",1,"A(AC)"],
           
           [0xf02d,1,"今日充電量",0,"Ah"],[0xf02e,1,"今日放電量",0,"Ah"],
           [0xf02f,1,"今日発電量",1,"kWh"],[0xf030,1,"今日消費量",1,"kWh"],
           [0xf03c,1,"今日商用充電量",0,"Ah"],[0xf03d,1,"今日商用電力消費量",1,"kWh"],
           [0xf03e,1,"今日インバーター稼働時間",0,"時間"],[0xf03f,1,"今日バイパス稼働時間",0,"時間"],

           [0xf034,2,"累積充電量",0,"Ah"],[0xf036,2,"累積放電量",0,"Ah"],
           [0xf038,2,"累積発電量",0,"kWh"],[0xf03a,2,"累積負荷積算電力量",0,"kWh"],
           [0xf046,2,"累積商用充電量",0,"kWh"],[0xf048,2,"累積商用負荷電力消費量",0,"kWh"],
           [0xf04a,1,"累積インバーター稼働時間",0,"時間"],[0xf04b,1,"累積バイパス稼働時間",0,"時間"]]
para_xy=[[0,1],[0,2],[0,3],[0,4],
         [0,6],[0,7],[0,8],[0,9],[0,10],[0,11],[0,12],[0,13],
         
         [5,1],[5,2],[5,3],[5,4],[5,5],[5,6],
         [5,8],[5,9],[5,10],[5,11],
         [5,13],[5,14],[5,15],[5,16],[5,17],
         
         [10,1],[10,2],[10,3],[10,4],[10,5],[10,6],[10,7],[10,8],
         [10,10],[10,11],[10,12],[10,13],[10,14],[10,15],[10,16],[10,17]]
sys_list=43

# ----------Modbus接続設定・・・・ID1:ポート5、ID2:ポート6  
client1=ModbusClient(framer="rtu",port="COM5",                  # USBポート5（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=1)
client2=ModbusClient(framer="rtu",port="COM6",                  # USBポート6（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=1)

# ----------Modbusデータ読出
def modbus_read(slave_add,slave_count,slave_id):                # ファンクション03h
    client1.connect()                                           # ポート接続
    client2.connect()
    read_data1=client1.read_holding_registers(
        address=slave_add,count=slave_count,slave=slave_id)
    read_data2=client2.read_holding_registers(
        address=slave_add,count=slave_count,slave=slave_id+1)
    client1.close()                                             # ポート切断
    client2.close()
    return read_data1.registers,read_data2.registers

# ----------CSVファイル設定・・・・ファイル名,ID1：シリアルナンバー
slave_id=1
sys_volt,sys_volt2=modbus_read(0xe003,1,slave_id)               # システム電圧読込
sys_id,sys_id2=modbus_read(0x0035,20,slave_id)                  # プロダクトID読込
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
for i in range(sys_list):
    name_data1.append(para_data[i][2])
    unit_data1.append(para_data[i][4])


# ----------データ変換
def change_type(data_type,rdd,sys_volt,d_type,d_add):           # データタイプ変換
    if data_type<5:
        if data_type<4:                                         # 1/10,1/100,1/1000判定
            if d_type==1 :                                      # 16bit
                if rdd>32767 :rdd=rdd-65536
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
        ym=hex(rdd)[2:].zfill(2)
        d=str(int(ym[0:2], 16)).zfill(2)+":"+str(int(ym[2:4], 16)).zfill(2)
    if data_type>=8:
        d=para_data[d_add][rdd+5]                               # 拡張判定
    return d

def data_read():
    dt_now=datetime.datetime.now()                              # 日時を取得
    date_time=dt_now.strftime('%y/%m/%d %H:%M:%S')
    p_type=0
    writer_data1=[]
    writer_data2=[]
    for a in range(sys_list):
        d_type=para_data[a][1]
        data_type=para_data[a][3]
        read_data1,read_data2=modbus_read(para_data[a][0],d_type,slave_id)
        if d_type==1:                                           # 16bitデータ変換
            data1=change_type(data_type,read_data1[0],sys_volt[0],d_type,a)
            data2=change_type(data_type,read_data2[0],sys_volt[0],d_type,a)
        else:
            data1=change_type(data_type,                        # 32bitデータ変換
                int(str(int(hex(read_data1[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(read_data1[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt,d_type,a)
            data2=change_type(data_type,
                int(str(int(hex(read_data2[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(read_data2[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt,d_type,a)
        writer_data1.append(data1)
        writer_data2.append(data2)
    #if p_type==0:prt="logging:"                  
    #else:prt="setting:"
    csv_data1=[date_time]
    csv_data2=[date_time]
    for i in range(sys_list):
        csv_data1.append(writer_data1[i])
        csv_data2.append(writer_data2[i])
    writer.writerow(csv_data1)                                  # CSVデータ書込
    writer.writerow(csv_data2)
    #print(prt,date_time)      
    return date_time,writer_data1,writer_data2

# ----------モニター画面
def create_gui():                                               # GUI作成
    root=tk.Tk()
    root.geometry('1250x450')                                   # ウインドウサイズ
    root.title("HYP4850U100-H")                                 # ウインドウタイトル
    frame=tk.Frame(root)
    frame.grid(row=0,column=0,sticky=tk.NSEW,padx=5,pady=10)
    wid=[20,13,1,13,8]
    label=tk.Label(frame,width=wid[0],text="日時",anchor=tk.W)
    label.grid(column=0,row=0)
    label0=tk.Label(frame,width=wid[1],text=date_time,anchor=tk.E,borderwidth=1)
    label0.grid(column=1,row=0)
    for j in range(3):
        labels=[tk.Label(frame,width=wid[i],text=["","ID1","","ID2",""][i]
                        ,anchor=tk.W)for i in range(5)]
        [labels[h].grid(column=h+j*5,row=1)for h in range(5)]
    col1='#0000ff'
    col2='#aaaaaa'
    labels=[tk.Label(frame,width=wid[0],text=para_data[i][2]
                     ,anchor=tk.W)for i in range(sys_list)]
    [labels[h].grid(column=para_xy[h][0],row=para_xy[h][1]+1)for h in range(sys_list)]
    
    labels1=[tk.Label(frame,width=wid[1],text=writer_data1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(sys_list)]
    [labels1[h].grid(column=para_xy[h][0]+1,row=para_xy[h][1]+1)for h in range(sys_list)]
    
    labels2=[tk.Label(frame,width=wid[3],text=writer_data1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(sys_list)]
    [labels2[h].grid(column=para_xy[h][0]+3,row=para_xy[h][1]+1)for h in range(sys_list)]
    
    labels=[tk.Label(frame,width=wid[4],text=para_data[i][4]
                     ,anchor=tk.W)for i in range(sys_list)]
    [labels[h].grid(column=para_xy[h][0]+4,row=para_xy[h][1]+1)for h in range(sys_list)]
    
    button=tk.Button(frame,text="計測終了",command=root.destroy) # ループ終了
    button.grid(column=13,row=0)

    thread=threading.Thread(target=update_data
                            ,args=([label0,[labels1[x]for x in range(sys_list)],
                                    [labels2[x]for x in range(sys_list)],]))
    thread.daemon=True                                          # メインウィンドウが閉じたらスレッド終了
    thread.start()                                              # スレッド処理開始
    root.mainloop()                                             # メインループ開始

# ----------データ更新処理
def update_data(label0,labels1,labels2):
        while True:
            date_time,writer_data1,writer_data2=data_read()
            label0.config(text=f"{date_time}")                  # GUIデータ更新
            [labels1[x].config(text=f"{writer_data1[x]}")for x in range(sys_list)]
            [labels2[x].config(text=f"{writer_data2[x]}")for x in range(sys_list)]
            sleep(interval)                                     # インターバルタイマー

# ----------実行
with open(file_name2,'w', newline='') as file:                  # CSVファイルオープン
        writer=csv.writer(file)
        w_data=[name_data1,unit_data1]
        [writer.writerow(w_data[i])for i in range(2)]           # ヘッダー書込
        date_time,writer_data1,writer_data2=data_read()
        if __name__ == "__main__":
            create_gui()
            
# 終了
