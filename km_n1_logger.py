# ----------初期設定
from pymodbus.client import ModbusSerialClient as ModbusClient # Modbusモジュール組込
import csv                                              # CSVファイルモジュール組込
import datetime                                         # 時計モジュール組込
from time import sleep                                  # タイマー組込
import tkinter as tk                                    # GUIモジュール組込
import threading                                        # スレッド組込

machine=2                                               # 計測機器台数　設定
interval=10                                             # 計測間隔（秒）設定
# ----------接続設定
client = ModbusClient(framer="rtu",port="COM5",         # USBポート（Windows）
    stopbits=2,bytesize=8,parity='N',baudrate=9600,timeout=1)

# ----------パラメーターデータ
pra_type=[]
ctrl_loop=3
ctrl_name=[["電圧1","電圧2","電圧3","電流1","電流2","電流3","力率","周波数"
           ,"有効電力","無効電力"]
           ,["積算有効電力量","積算回生電力量","積算進み無効電力量","積算遅れ無効電力量"
           ,"積算総合無効電力量"]
           ,["積算有効電力量","積算回生電力量","積算進み無効電力量","積算遅れ無効電力量"
           ,"積算総合無効電力量"]]
ctrl_unit=[["V","V","V","A","A","A","","Hz","W","Var"]
           ,["Wh","Wh","Varh","Varh","Varh"]
           ,["kWh","kWh","kVarh","kVarh","kVarh"]]
ctrl_type=[[1,1,1,3,3,3,2,1,1,1],[0,0,0,0,0],[0,0,0,0,0]]
ctrl_count=[10,5,5]
ctrl_add=[0x0000,0x0200,0x0220]

# ----------CSVファイル設定
file_name='omron'
dt_now = datetime.datetime.now()                        # 日時を取得
file_time=dt_now.strftime('_20%y_%m_%d_%H%M')
file_name=file_name+file_time+'.csv'
print("ファイル名:",file_name)
id_data=["","入力側"]
for i in range(19):id_data.append("")
id_data.append("出力側")
name_data=["日時"]
unit_data=["年月日時分秒"]
for a in range(machine):
    for b in range(ctrl_loop):
        data1=ctrl_name[b]
        data2=ctrl_unit[b]
        data3=ctrl_count[b]
        for c in range(data3):
            name_data.append(data1[c])
            unit_data.append(data2[c])

# ----------データ変換
def change_type(data_type,rdd,sys_volt,cnt_count):      # データタイプ変換
    if data_type<5:
        if data_type<4:                                 # 1/10,1/100,1/1000判定
            if cnt_count==1 :                           # 1byte
                if rdd>32767 :rdd=rdd-65536
            if cnt_count==2 :                           # 2byte
                if rdd>2147483647 :rdd=rdd-4294967295
            data_list=[rdd,rdd/10,rdd/100,rdd/1000]
            d=data_list[data_type]
        if data_type==4:                                # 電圧データ判定24V/48V
            d=rdd/10
            if sys_volt[0]==24:d=d*2  
            if sys_volt[0]==48:d=d*4   
        d=str(d)
    if data_type==5:d=chr(rdd)                          # 文字データ判定　
    if data_type==6:                                    # ON/OFF判定
        data_list=["OFF","ON"]
        d=data_list[rdd]
    if data_type==7:                                    # 時刻判定
        ym=hex(rdd)[2:].zfill(2)
        d=str(int(ym[0:2], 16)).zfill(2)+":"+str(int(ym[2:4], 16)).zfill(2)
    if data_type>=8:                                    # 拡張判定
        d_type=pra_type[data_type-8]
        d=d_type[rdd]
    return d

# ----------Modbus読出
def data_set(date_time):                                # データ読出
    set_data=[]
    client.connect()                                    # ポート接続開始
    for a in range(machine):
        for b in range(ctrl_loop):
            ctrl_data=client.read_holding_registers(    # Modbus ファンクション03H
                address=ctrl_add[b],count=ctrl_count[b]*2,slave=a+1)  
            cnt_rdd=ctrl_data.registers
#            cnt_rdd=test(ctrl_add[b],a)                # 未接続テスト用
            cnt_type=ctrl_type[b]
            cnt_data=[]
            for c in range(ctrl_count[b]):              # 一括データ変換
                add=c*2
                cnt_data.append(change_type(
                    cnt_type[c],int(str(int(hex(cnt_rdd[add])[2:].zfill(2)+
                        hex(cnt_rdd[add+1])[2:].zfill(2),16)).zfill(2)),0,2))
            set_data=set_data+cnt_data
    client.close()    
    write_data=[date_time]+set_data
    #print("Modbus :",write_data)                        # コンソール画面に出力  
    return write_data

# ----------未接続テスト用
def test(sl_add,a):                                     # テストパラメーター
    testdata=[[0,966,0,1012,0,1977,0,5943,0,3466,0,3208,0,93,0,600,0,8552,0,6329]
              ,[0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0]
              ,[0,957,0,1014,0,1971,0,5399,0,3161,0,2942,0,65444,0,600,0,57826,0,3200]
              ,[0,0,0,0,0,0,0,0,0,0],[0,0,0,0,0,0,0,0,0,0]]
    if sl_add==0x0000:
        if a==0:ctrl_data=testdata[0]
        else:ctrl_data=testdata[3]
    if sl_add==0x0200:
        if a==0:ctrl_data=testdata[1]
        else:ctrl_data=testdata[4]
    if sl_add==0x0220:
        if a==0:ctrl_data=testdata[2]
        else:ctrl_data=testdata[5]
    return ctrl_data

# ----------GUI作成
def create_gui():                                       # GUI作成
    root=tk.Tk()
    root.geometry('1000x300')                           # ウインドウサイズ
    root.title("OMRON KM-N1")                           # ウインドウタイトル
    frame=tk.Frame(root)
    frame.grid(row=0,column=0,sticky=tk.NSEW,padx=5,pady=10)
    wid=[15,10,5]
    txt=[name_data[0],"ID 1(入力側)","ID 2(出力側)"]
    loc_x=[0,0,6]
    loc_y=[0,1,1]
    labels=[tk.Label(frame,width=wid[0],text=txt[i],anchor=tk.W)for i in range(3)]
    [labels[h].grid(column=loc_x[h],row=loc_y[h])for h in range(3)]
    label0=tk.Label(frame,width=wid[0],text="",anchor=tk.E,borderwidth=1)
    label0.grid(column=1,row=0)
    col1='#0000ff'
    col2='#aaaaaa'
    txt1=ctrl_name[0]+ctrl_name[1]+ctrl_name[2]+ctrl_name[0]+ctrl_name[1]+ctrl_name[2]
    txt3=ctrl_unit[0]+ctrl_unit[1]+ctrl_unit[2]+ctrl_unit[0]+ctrl_unit[1]+ctrl_unit[2]
    loc_x=[]
    loc_y=[]
    for x in range(4):
        for y in range(10):
            loc_y.append(y+2)
            loc_x.append(x*3)
    print(loc_x,loc_y)
    labels=[tk.Label(frame,width=wid[0],text=txt1[i],anchor=tk.W)for i in range(40)]
    [labels[h].grid(column=loc_x[h],row=loc_y[h])for h in range(40)]
    labels1=[tk.Label(frame,width=wid[1],text="",anchor=tk.E,relief=tk.SOLID,
                            borderwidth=1,foreground=col1,background=col2)for i in range(40)]
    [labels1[h].grid(column=loc_x[h]+1,row=loc_y[h])for h in range(40)]
    labels=[tk.Label(frame,width=wid[2],text=txt3[i],anchor=tk.W)for i in range(40)]
    [labels[h].grid(column=loc_x[h]+2,row=loc_y[h])for h in range(40)]
    
    button=tk.Button(frame,text="計測終了",command=root.destroy) # ループ終了
    button.grid(column=11,row=0)
    
    thread=threading.Thread(target=update_data
                            ,args=([label0,[labels1[x]for x in range(40)],]))
    thread.daemon=True                                  # メインウィンドウが閉じたらスレッドも終了
    thread.start()                                      # スレッド処理開始
    root.mainloop()                                     # メインループ開始

# ----------データ更新処理
def update_data(label0,labels1):
        while True:
            dt_now=datetime.datetime.now()              # 日時を取得
            date_time=dt_now.strftime('%y/%m/%d %H:%M:%S')
            writer_data=data_set(date_time)             # Modbusデータ読込
            writer.writerow(writer_data)                # CSVデータ書込
            label0.config(text=f"{writer_data[0]}")     # GUIデータ更新
            [labels1[x].config(text=f"{writer_data[x+1]}")for x in range(40)]
            sleep(interval)                             # インターバルタイマー

# ----------スタート
with open(file_name, 'w', newline='') as file:          # CSVファイルオープン
        writer = csv.writer(file)
        w_data=[id_data,name_data,unit_data]
        [writer.writerow(w_data[i])for i in range(3)]   # ヘッダー書込
       
        if __name__ == "__main__":
            create_gui()
            
# 終了
