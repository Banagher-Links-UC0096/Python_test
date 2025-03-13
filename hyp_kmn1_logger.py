# HYP4850U100-H_parallel_KM-N1_Logger
# ----------初期設定
from pymodbus.client import ModbusSerialClient as ModbusClient  # Modbus組込
import csv                                                      # CSVファイルモジュール組込
#from openpyxl import load_workbook                              # xlsxファイルモジュール組込
import datetime                                                 # 時計モジュール組込
from time import sleep                                          # タイマー組込
import tkinter as tk                                            # GUIモジュール組込
import threading                                                # スレッド組込
import matplotlib.pyplot as plt                                 # グラフ作成
import matplotlib.ticker as ticker                              # グラフ補助
import matplotlib.animation as animation                        # グラフ補助
import matplotlib.font_manager as fm                            # グラフ補助
import pandas as pd                                             # ファイルモジュール組込

import numpy as np
import seaborn as sns
import glob
import os
#import pydrive                                                  # GoogleDrive組込


interval1=5.9                                                   # 待ち時間（秒）

# ----------ログパラメーター [0address,1byte,2x,3y,4type,5name,6unit,7len,8min,9max,10set,11data...]
p_data=[[0x0213,1,0,2,1,"系統電圧","V(AC)"],
        [0x0214,1,0,3,1,"系統電流","A(AC)"],
        [0x0215,1,0,4,2,"系統周波数","Hz"],
        [0x021e,1,0,5,1,"系統充電電流","A(AC)"],
        [0x0210,1,0,7,8,"機器状態","",0,0,0,"起動","待機","初期化","省電力","商用出力",
        "インバーター出力","系統出力","混合出力","-","-","停止","故障"],
        [0x0216,1,0,8,1,"出力電圧","V(AC)"],
        [0x0219,1,0,9,1,"出力電流","A(AC)"],
        [0x0218,1,0,10,2,"出力周波数","Hz"],
        [0x021b,1,0,11,0,"負荷有効電力","W"],
        [0x021c,1,0,12,0,"負荷皮相電力","W"],
        [0x021f,1,0,13,0,"負荷率","%"],
        
        [0x010b,1,0,15,8,"充電状態","",0,0,0,"未充電","定電流(CC)充電","定電圧(CV)充電","-",
        "浮遊充電","-","充電中1","充電中2"],
        [0x0100,1,0,16,0,"蓄電池SOC","%"],
        [0x0101,1,0,17,1,"蓄電池電圧","V(DC)"],
        [0x0102,1,0,18,1,"蓄電池電流","A(DC)"],
        [0x010e,1,0,19,0,"充電電力","W"],
        [0x0217,1,0,20,1,"INV電流 ","A(DC)"],
        
        [0x0107,1,0,22,1,"PV入力電圧","V(DC)"],
        [0x0108,1,0,23,1,"PV入力電流","A(DC)"],
        [0x0109,1,0,24,0,"PV入力電力","W"],
        [0x0224,1,0,25,1,"PV降圧電流","A(DC)"],
           
        [0x0212,1,0,27,1,"DCバス電圧","V(DC)"],
        [0x0225,1,0,28,1,"降圧電流","A(DC)"],
        [0x0220,1,0,29,1,"PVHT温度","℃"],
        [0x0221,1,0,30,1,"INVHT温度","℃"],
        [0x0222,1,0,31,1,"Tr温度","℃"],
        [0x0223,1,0,32,1,"内部温度","℃"],
           
        [0xf02d,1,5,2,0,"本日充電量","Ah"],
        [0xf02e,1,5,3,0,"本日放電量","Ah"],
        [0xf02f,1,5,4,1,"本日発電量","kWh"],
        [0xf030,1,5,5,1,"本日消費量","kWh"],
        [0xf03c,1,5,6,0,"本日商用充電量","Ah"],
        [0xf03d,1,5,7,1,"本日商用電力消費量","kWh"],
        [0xf03e,1,5,8,0,"本日インバーター稼働時間","時間"],
        [0xf03f,1,5,9,0,"本日バイパス稼働時間","時間"],

        [0xf034,2,5,11,0,"累積充電量","Ah"],
        [0xf036,2,5,12,0,"累積放電量","Ah"],
        [0xf038,2,5,13,0,"累積発電量","kWh"],
        [0xf03a,2,5,14,0,"累積負荷積算電力量","kWh"],
        [0xf046,2,5,15,0,"累積商用充電量","kWh"],
        [0xf048,2,5,16,0,"累積商用負荷電力消費量","kWh"],
        [0xf04a,1,5,17,0,"累積インバーター稼働時間","時間"],
        [0xf04b,1,5,18,0,"累積バイパス稼働時間","時間"]]
r_data=[["蓄電池総合電流",5,20,1,1,"A"],["蓄電池総合電力",5,21,2,1,"kW"],
        ["PV発電総合電力",5,22,3,3,"kW"],["AC入力総合電力",5,23,4,3,"kVA"],
        ["AC出力総合電力",5,24,5,3,"kW"],["機器消費電力",5,25,6,0,"W"],
         
        ["総合充電量",5,27,7,0,"Ah"],["総合放電量",5,28,8,0,"Ah"],
        ["総合発電量",5,29,9,1,"kWh"],["総合消費量",5,30,10,1,"kWh"],
        ["総合商用充電量",5,31,11,0,"Ah"],["総合商用消費量",5,32,12,1,"kWh"]]
t_data=[["日時",4,0,4],["L1側",1,1,1],["L2側",3,1,1],["L1側",6,1,1],
        ["L2側",8,1,1],["AC入力側",0,1,0],["AC出力側",0,6,0],
        ["総合データ",5,19,0],["蓄電池側",0,14,0],["PV入力側",0,21,0],
        ["内部システム",0,26,0],["当日積算データ",5,1,0],
        ["累積積算データ",5,10,0],["KM-N1データ",0,35,0],["入力側",1,35,1],
        ["出力側",3,35,1],["入力側",6,35,1],["出力側",8,35,1]]
k_data=[[0x0000,2,0,36,1,"電圧1","V"],
        [0x0002,2,0,37,1,"電圧2","V"],
        [0x0004,2,0,38,1,"電圧3","V"],
        [0x0006,2,0,39,3,"電流1","A"],
        [0x0008,2,0,40,3,"電流2","A"],
        [0x000a,2,0,41,3,"電流3","A"],
        [0x000c,2,0,42,2,"力率",""],
        [0x000e,2,0,43,1,"周波数","Hz"],
        [0x0010,2,0,44,1,"有効電力","W"],
        [0x0012,2,0,45,1,"無効電力","Var"],
        [0x0200,2,5,36,0,"積算有効電力量","Wh"],
        [0x0202,2,5,37,0,"積算回生電力量","Wh"],
        [0x0204,2,5,38,0,"積算進み無効電力量","Varh"],
        [0x0206,2,5,39,0,"積算遅れ無効電力量","Varh"],
        [0x0208,2,5,40,0,"積算総合無効電力量","Varh"],
        [0x0220,2,5,41,0,"積算有効電力量","kWh"],
        [0x0222,2,5,42,0,"積算回生電力量","kWh"],
        [0x0224,2,5,43,0,"積算進み無効電力量","kVarh"],
        [0x0226,2,5,44,0,"積算遅れ無効電力量","kVarh"],
        [0x0228,2,5,45,0,"積算総合無効電力量","kVarh"]]
# ----------設定パラメーター [0address,1byte,2x,3y,4type,5name,6unit,7len,8min,9max,10set,11data...]
s_data=[[0xe004,1,0,2,8,"蓄電池タイプ","",14,0,13,"設定08","ユーザー設定","密閉型鉛","開放型鉛","ゲル型鉛",
          "LFPx14","LEPx15","LFPx16","LFPx7","LFPx8","LFPx8","NCAx7","NCAx8","NCAx13","NCAx14"],#0
        [0xe20f,1,0,3,8,"充電モード","",4,0,3,"設定06","PV優先","系統優先","ハイブリッド","PV専用"],#1
        [0xe002,1,0,4,1,"蓄電池容量","Ah",900,0,900,""],#2
        [0xe005,1,0,5,4,"過電圧停止電圧","V",50,100,150,""],#3
        [0xe006,1,0,6,4,"充電上限電圧","V",50,100,150,""],#4
        [0xe009,1,0,7,4,"CV充電電圧","V",50,100,150,"設定11"],#5
        [0xe01c,1,0,8,1,"充電停止電流","A",100,0,100,"設定57"],#6
        [0xe008,1,0,9,4,"CC充電電圧","V",50,100,150,"設定09"],#7
        [0xe20a,1,0,10,1,"最大充電電流","A",600,0,600,"設定07"],#8
        [0xe012,1,0,11,0,"CC充電遅延時間","min",120,0,120,"設定10"],#9
        [0xe022,1,0,12,4,"INV切替電圧","V",50,100,150,"設定05"],#10
        [0xe00a,1,0,13,4,"充電再開電圧","V",50,100,150,"設定37"],#11
        [0xe00b,1,0,14,4,"低電圧復帰電圧","V",50,100,150,"設定35"],#12
        [0xe00c,1,0,15,4,"低電圧警告電圧","V",50,100,150,"設定14"],#13
        [0xe01b,1,0,16,4,"バイパス切替電圧","V",50,100,150,"設定04"],#14
        [0xe00d,1,0,17,4,"過放電遅延オフ電圧","V",50,100,150,"設定12"],#15
        [0xe010,1,0,18,0,"過放電遅延オフ時間","s",60,0,60,"設定13"],#16
        [0xe00e,1,0,19,4,"放電停止電圧","V",50,100,150,"設定15"],#17

        [0xe206,1,0,20,6,"均等充電有無","",2,0,1,"設定16"],#18
        [0xe007,1,0,21,4,"均等充電電圧","V",50,100,150,"設定17"],#19
        [0xe011,1,0,22,0,"均等充電時間","min",120,0,120,"設定18"],#20
        [0xe023,1,0,23,0,"均等充電遅延","min",120,0,120,"設定19"],#21
        [0xe013,1,0,24,0,"均等充電間隔","day",7,0,7,"設定20"],#22
        [0xdf0d,1,0,25,6,"均等充電有無","",2,0,1,"設定21"],#23
          
        [0xe215,1,5,2,8,"BMS通信","",3,0,2,"設定32","OFF","RS485-BMS","CAN-BMS"],#24
        [0xe21b,1,5,3,8,"BMSプロトコル","",18,0,17,"設定33","Pace","Rata","Allgrand","Oliter","PCT",
          "Sunwoda","Dyness","SRNE","Pylontech","","","","","","","","WS Technicals","Uz Energy"],#25
        [0xe025,1,5,4,8,"充電制御","",3,0,2,"設定39","OFF","BMS制御","INV制御"],#26
        [0xe01e,1,5,5,0,"過放電警報SOC","%",100,0,100,"設定58"],#27
        [0xe00f,1,5,6,0,"放電停止SOC","%",100,0,100,"設定59"],#28
        [0xe01d,1,5,7,0,"充電停止SOC","%",100,0,100,"設定60"],#29
        [0xe01f,1,5,8,0,"バイパス切替SOC","%",100,0,100,"設定61"],#30
        [0xe020,1,5,9,0,"INV切替SOC","%",100,1,100,"設定62"],#31
          
        [0xe11f,1,5,10,0,"PV最大電圧","V",500,0,500,""],#32
        [0xe001,1,5,11,1,"PV最大充電電流","A",800,0,800,"設定36"],#33
        [0xe120,1,5,12,1,"最大充電電流","A",1000,0,1000,""],#34
          
        [0xe204,1,5,13,8,"出力優先","",3,0,2,"設定01","PV優先","系統優先","蓄電池優先"],#35
        [0xe209,1,5,14,2,"AC出力周波数","Hz",2,5000,6000,"設定02"],#36
        [0xe208,1,5,15,1,"AC出力電圧","V",20,100,120,"設定38"],#37
        [0xe129,1,5,16,1,"最大出力電流","A",420,0,420,""],#38
        [0xe118,1,5,17,1,"定格出力","kW",100,0,100,""],#39
                 
        [0xe20b,1,5,18,8,"AC入力電圧範囲","",2,0,1,"設定03","APL(90-280V)","UPS(90-140V)"],#40
        [0xe201,1,5,19,8,"並列運転モード","",8,0,7,"設定31","単相","並列","2P0","2P1","2P2","3P1","3P2","3P3"],#41
        [0xe205,1,5,20,1,"AC充電最大電流","A",400,0,400,"設定28"],#42
        [0xe037,1,5,21,8,"PV出力先","",3,0,2,"設定34","OFF","AC入力PV出力（逆潮流）","AC出力PV出力（順潮流）"],#43
          
        [0xe20c,1,5,22,6,"省エネモード","",2,0,1,"設定22"],#44
        [0xe212,1,5,23,6,"バイパス出力有無","",2,0,1,"設定27"],#45
        [0xe200,1,5,24,0,"RS485アドレス","",240,1,240,"設定30"],#46
        [0xe214,1,5,25,6,"分相変圧器","",2,0,1,"設定29"],#47
        [0xe207,1,5,26,6,"N相出力","",2,0,1,"設定63"],#48
        [0xe038,1,5,27,6,"漏電検知機能","",2,0,1,"設定56"],#49

        [0xe213,1,5,28,6,"障害記録","",2,0,1,""],#50
        [0xe210,1,5,29,6,"警報音有無","",2,0,1,"設定25"],#51
        [0xe211,1,5,30,6,"電源切替警報有無","",2,0,1,"設定26"],#52
        [0xe20d,1,5,31,6,"過負荷停止再起動","",2,0,1,"設定23"],#53
        [0xe20e,1,5,32,6,"高温停止再起動","",2,0,1,"設定24"], #54

        #[0xe,5,23,"設定54,55","",9,"",0,0,0],
        [0xe026,1,10,2,7,"充電開始時間１","",0,0,0,"設定40"],
        [0xe027,1,10,3,7,"充電終了時間１","",0,0,0,"設定41"],
        [0xe028,1,10,4,7,"充電開始時間２","",0,0,0,"設定42"],
        [0xe029,1,10,5,7,"充電終了時間２","",0,0,0,"設定43"],
        [0xe02a,1,10,6,7,"充電開始時間３","",0,0,0,"設定44"],
        [0xe02b,1,10,7,7,"充電終了時間３","",0,0,0,"設定45"],
        [0xe02c,1,10,8,6,"充電時間設定","",2,0,1,"設定46"],
        [0xe02d,1,10,9,7,"放電開始時間１","",0,0,0,"設定47"],
        [0xe02e,1,10,10,7,"放電終了時間１","",0,0,0,"設定48"],
        [0xe02f,1,10,11,7,"放電開始時間２","",0,0,0,"設定49"],
        [0xe030,1,10,12,7,"放電終了時間２","",0,0,0,"設定50"],
        [0xe031,1,10,13,7,"放電開始時間３","",0,0,0,"設定51"],
        [0xe032,1,10,14,7,"放電終了時間３","",0,0,0,"設定52"],
        [0xe033,1,10,15,6,"放電時間設定","",0,0,0,"設定53"]]

# ----------Modbus接続設定
client1=ModbusClient(framer="rtu",port="COM7",                  # Hybrid1（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=1)
client2=ModbusClient(framer="rtu",port="COM8",                  # Hyblid2（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=1)
client3=ModbusClient(framer="rtu",port="COM9",                  # KM-N112（Windows）
                       baudrate=9600,bytesize=8,stopbits=2,parity='N',timeout=1)
# ----------ハイブリッドインバーターデータ取得
def hybrid_modbus_read(slave_add,slave_count):                  # ファンクション03h
    client1.connect()                                           # ポート接続
    client2.connect()
    read_data1=client1.read_holding_registers(
        address=slave_add,count=slave_count,slave=0)
    read_data2=client2.read_holding_registers(
        address=slave_add,count=slave_count,slave=1)
    client1.close()                                             # ポート切断
    client2.close()
    return read_data1.registers,read_data2.registers
# ----------未接続時テスト用
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
    #print(hex(slave_add))
    for v in range(len(test_data)):
        if slave_add==test_data[v][0]:
            read_data1=test_data[v][1]
            read_data2=test_data[v][2]
        if slave_add==0xe003:
            read_data1=[48]
            read_data2=read_data1
        if slave_add==0x0035:
                read_data1=[51,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50]
                read_data2=[50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50,50]
    return read_data1,read_data2

# ----------KM-N1データ取得
def kmn1_modbus_read(slave_add,slave_count):                    # ファンクション03h
    client3.connect()                                           # ポート接続
    read_data3=client3.read_holding_registers(
        address=slave_add,count=slave_count,slave=0)
    read_data4=client3.read_holding_registers(
        address=slave_add,count=slave_count,slave=1)
    client3.close()                                             # ポート切断
    return read_data3.registers,read_data4.registers
# ----------未接続時テスト用
    test_data=[[0x0000,[0,0],[0,966]],[0x0002,[0,0],[0,1012]],[0x0004,[0,0],[0,1977]],
               [0x0006,[0,0],[0,5943]],[0x0008,[0,0],[0,3466]],[0x000a,[0,0],[0,3208]],
               [0x000a,[0,0],[0,93]],[0x000c,[0,0],[0,600]],[0x000e,[0,0],[0,0]],
               [0x0010,[0,0],[0,8552]],[0x0012,[0,0],[0,6329]],
               [0x0200,[0,0],[0,0]],[0x0202,[0,0],[0,0]],[0x0204,[0,0],[0,0]],
               [0x0206,[0,0],[0,0]],[0x0208,[0,0],[0,0]],
               [0x0220,[0,0],[0,0]],[0x0222,[0,0],[0,0]],[0x0224,[0,0],[0,0]],
               [0x0226,[0,0],[0,0]],[0x0228,[0,0],[0,0]]]
    #print(hex(slave_add))
    for v in range(len(test_data)):
        if slave_add==test_data[v][0]:
            read_data3=test_data[v][1]
            read_data4=test_data[v][2]
    return read_data3,read_data4

# ----------データ変換    
def change_minus(rdd):                                          # 負変換
    if rdd>32767 :rdd=rdd-65536
    return rdd

def change_type(d_type,rdd,sys_volt,byte,d_add,n_data):         # データタイプ変換
    if d_type<5:
        if d_type<4:                                            # 1/10,1/100,1/1000判定
            if byte==1 :                                        # 16bit
                rdd=change_minus(rdd)
            if byte==2 :                                        # 32byte
                if rdd>2147483647 :rdd=rdd-4294967295
            data_list=[rdd,rdd/10,rdd/100,rdd/1000]
            d=data_list[d_type]
        if d_type==4:                                           # 電圧データ判定24V/48V
            d=rdd/10
            if sys_volt==24:d=d*2  
            if sys_volt==48:d=d*4   
        d=str(d)
    if d_type==5:d=chr(rdd)                                     # 文字データ判定　
    if d_type==6:d=["OFF","ON"][rdd]                            # ON/OFF判定  
    if d_type==7:                                               # 時刻判定
        ym=hex(rdd)[4:].zfill(4)
        d=str(int(ym[0:2],16)).zfill(2)+":"+str(int(ym[2:4],16)).zfill(2)
    if d_type==8:d=n_data[d_add][rdd+11]                        # 拡張判定 p:ps=7,s:ps=10
    if d_type==9:print(rdd)
        
    return d

def data_read(n_data):                                          # データ処理
    dt_now=datetime.datetime.now()                              # 日時を取得
    date_time=dt_now.strftime('%y/%m/%d %H:%M:%S')
    hiwd1,hiwd2,hird1,hird2,o_data=[],[],[],[],[]
    kmwd1,kmwd2,kmrd1,kmrd2=[],[],[],[]
    n_list=len(n_data)
    for a in range(n_list):
        byte=n_data[a][1]
        d_type=n_data[a][4]
        r_data1,r_data2=[],[]
        r_data1,r_data2=hybrid_modbus_read(n_data[a][0],byte)
        if byte==1:                                             # 16bitデータ変換
            data1=change_type(d_type,r_data1[0],sys_volt1[0],byte,a,n_data)
            data2=change_type(d_type,r_data2[0],sys_volt1[0],byte,a,n_data)
        else:                                                   # 32bitデータ変換
            data1=change_type(d_type,
                int(str(int(hex(r_data1[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(r_data1[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt1,byte,a,n_data)
            data2=change_type(d_type,
                int(str(int(hex(r_data2[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(r_data2[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt1,byte,a,n_data)
        hiwd1.append(data1)
        hiwd2.append(data2)
        hird1.append(r_data1)
        hird2.append(r_data2)
    for a in range(len(k_data)):
        r_data1,r_data2=kmn1_modbus_read(k_data[a][0],2,)
        data1=change_type(d_type,
                int(str(int(hex(r_data1[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(r_data1[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt1,byte,a,n_data)
        data2=change_type(d_type,
                int(str(int(hex(r_data2[1])[2:].zfill(2)[0:4],16))
                +str(int(hex(r_data2[0])[2:].zfill(2)[0:4],16)))
                ,sys_volt1,byte,a,n_data)
        kmwd1.append(data1)
        kmwd2.append(data2)
        kmrd1.append(r_data1)
        kmrd2.append(r_data2)
    csv_data1=[date_time]
    csv_data2=[date_time]
    csv_data1=csv_data1+hiwd1
    csv_data2=csv_data2+hiwd2
    
    rdd=hird1[14][0]
    q=change_minus(rdd)
    rdd=hird2[14][0]
    r=change_minus(rdd)
    i_calc=((hird1[0][0]*hird1[1][0])+(hird2[0][0]*hird2[1][0]))/100000 # AC入力合算
    o_calc=(hird1[8][0]+hird2[8][0])/1000                               # AC出力合算
    p_calc=(hird1[19][0]+hird2[19][0])/1000                             # PV入力合算
    b_calc=round(((hird1[13][0]*q+hird2[13][0]*r)/100)/1000,3)          # バッテリー電力合算
    b_curr=((q+r)/10)                                                   # バッテリー電流合算
    calc_list=[b_curr,b_calc,p_calc,i_calc,o_calc
               ,round(b_calc+o_calc-p_calc-i_calc,3),
               hird1[27][0]+hird2[27][0],hird1[28][0]+hird2[28][0],
               (hird1[29][0]+hird2[29][0])/10,(hird1[30][0]+hird2[30][0])/10,
               hird1[31][0]+hird2[31][0],(hird1[32][0]+hird2[32][0])/10]
    for s in range(len(r_data)):o_data.append(calc_list[s])
    return date_time,hiwd1,hiwd2,kmwd1,kmwd2,o_data,csv_data1,csv_data2

# ----------モニター画面
def create_gui():                                               # GUI作成
    root=tk.Tk()
    root.geometry('830x950+20+20')                              # ウインドウサイズ
    root.title("HYP4850U100-H 並列モニター")                     # ウインドウタイトル
    frame=tk.Frame(root)
    frame.grid(row=0,column=0,sticky=tk.NSEW,padx=5,pady=10)
    wid=[19,13,5,13,5]
    col1='#0000ff'                                              #データ文字色
    col2='#cccccc'                                              #データ背景色
    p_list,t_list,r_list,k_list=len(p_data),len(t_data),len(r_data),len(k_data)
    
    labels=[tk.Label(frame,width=wid[t_data[i][3]],text=t_data[i][0],font=("Atari",9,"bold")
                        ,anchor=tk.W)for i in range(t_list)]
    [labels[h].grid(column=t_data[h][1],row=t_data[h][2])for h in range(t_list)]
    label0=tk.Label(frame,width=wid[0],text=date_time,anchor=tk.W,borderwidth=1)
    label0.grid(column=5,row=0)
    labels=[tk.Label(frame,width=wid[0],text=p_data[i][5]
                     ,anchor=tk.W)for i in range(p_list)]
    [labels[h].grid(column=p_data[h][2],row=p_data[h][3])for h in range(p_list)]
    labels1=[tk.Label(frame,width=wid[1],text=hiwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(p_list)]
    [labels1[h].grid(column=p_data[h][2]+1,row=p_data[h][3])for h in range(p_list)]
    labels2=[tk.Label(frame,width=wid[3],text=hiwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(p_list)]
    [labels2[h].grid(column=p_data[h][2]+3,row=p_data[h][3])for h in range(p_list)]
    for j in range(2):
        labels=[tk.Label(frame,width=wid[4],text=p_data[i][4]
                        ,anchor=tk.W)for i in range(p_list)]
        [labels[h].grid(column=p_data[h][2]+2+2*j,row=p_data[h][3])for h in range(p_list)]
    labels=[tk.Label(frame,width=wid[0],text=r_data[i][0]
                        ,anchor=tk.W)for i in range(r_list)]
    [labels[h].grid(column=r_data[h][1],row=r_data[h][2])for h in range(r_list)]
    labels3=[tk.Label(frame,width=wid[1],text=o_data[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(r_list)]
    [labels3[h].grid(column=r_data[h][1]+1,row=r_data[h][2])for h in range(r_list)]
    labels=[tk.Label(frame,width=wid[4],text=r_data[i][5]
                        ,anchor=tk.W)for i in range(r_list)]
    [labels[h].grid(column=r_data[h][1]+2,row=r_data[h][2])for h in range(r_list)]
    
    labels=[tk.Label(frame,width=wid[0],text=k_data[i][5]
                        ,anchor=tk.W)for i in range(k_list)]
    [labels[h].grid(column=k_data[h][2],row=k_data[h][3])for h in range(k_list)]
    labels4=[tk.Label(frame,width=wid[1],text=kmwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(k_list)]
    [labels4[h].grid(column=k_data[h][2]+1,row=k_data[h][3])for h in range(k_list)]
    labels5=[tk.Label(frame,width=wid[3],text=kmwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(k_list)]
    [labels5[h].grid(column=k_data[h][2]+3,row=k_data[h][3])for h in range(k_list)]
    
    button1=tk.Button(frame,text="計測終了",command=root.destroy)    # ループ終了
    button1.grid(column=8,row=0)
    button2=tk.Button(frame,text="機器機器設定",command=root.destroy) # ループ終了
    button2.grid(column=6,row=0)
    button3=tk.Button(frame,text="チャート表示",command=root.destroy) # チャート移動
    button3.grid(column=3,row=0)

    thread1=threading.Thread(target=update_data,args=([label0,
                                    [labels1[x]for x in range(p_list)],
                                    [labels2[x]for x in range(p_list)],
                                    [labels3[x]for x in range(r_list)],
                                    [labels4[x]for x in range(k_list)],
                                    [labels5[x]for x in range(k_list)],]))
    
    thread1.daemon=True                                         # スレッド終了
    thread1.start()                                             # スレッド処理開始
    root.mainloop()                                             # メインループ開始

# ----------データ更新処理
def update_data(label0,labels1,labels2,labels3,labels4,labels5):
        while True:
            p_list,k_list=len(p_data),len(k_data)
            date_time,hiwd1,hiwd2,kmwd1,kmwd2,o_data,csv_data1,csv_data2=data_read(p_data)
            writer1.writerow(csv_data1)
            writer2.writerow(csv_data2)
            label0.config(text=f"{date_time}")                  # GUIデータ更新
            [labels1[x].config(text=f"{hiwd1[x]}")for x in range(p_list)]
            [labels2[x].config(text=f"{hiwd2[x]}")for x in range(p_list)]
            [labels3[x].config(text=f"{o_data[x]}")for x in range(len(r_data))]
            [labels4[x].config(text=f"{kmwd1[x]}")for x in range(k_list)]
            [labels5[x].config(text=f"{kmwd2[x]}")for x in range(k_list)]
            sleep(interval1)                                    # インターバルタイマー

# ----------CSVファイル設定
sys_volt1,sys_volt2=hybrid_modbus_read(0xe003,1)                # システム電圧読込
sys_id1,sys_id2=hybrid_modbus_read(0x0035,20)                   # プロダクトID読込
dt_now = datetime.datetime.now()                                # 日時を取得
file_time=dt_now.strftime('_20%y_%m_%d_%H%M')
id_name1,id_name2,id_name3='','','omuron'
for a in range(20):
    id_name1=id_name1+chr(sys_id1[a])
    id_name2=id_name2+chr(sys_id2[a])
file_name='today_logfile.csv'
file_name1=id_name1+file_time+'.csv'                            # ID1ファイル名作成
file_name2=id_name2+file_time+'.csv'                            # ID2ファイル名作成
file_name3=id_name3+file_time+'.csv'                            # KM-N1ファイル名作成
print("ID1 ログファイル名:",file_name1)
print("ID2 ログファイル名:",file_name2)
print("電力計ログファイル名:",file_name3)

for a in range(3):
    m_data1=[p_data,p_data,k_data]
    n_data,u_data=[file_time],["日付"]
    for b in range(len(m_data1[a])):
        n_data.append(m_data1[a][b][5])
        u_data.append(m_data1[a][b][6])
    if a==0:n_data1,u_data1=n_data,u_data
    if a==1:n_data2,u_data2=n_data,u_data
    if a==2:n_data3,u_data3=n_data,u_data

# ----------実行
with open(file_name,'w', newline='') as file:                  # CSVファイルオープン
        writer=csv.writer(file)
        with open(file_name1,'w',newline='')as file1:
            writer1=csv.writer(file1)
            [writer1.writerow([n_data1,u_data1][i])for i in range(2)]# ヘッダー書込
            with open(file_name2,'w',newline='')as file2:
                writer2=csv.writer(file2)
                [writer2.writerow([n_data1,u_data1][i])for i in range(2)]
                date_time,hiwd1,hiwd2,kmwd1,kmwd2,o_data,csv_data1,csv_data2=data_read(p_data)
                writer1.writerow(csv_data1)
                writer2.writerow(csv_data2)
                if __name__ == "__main__":
                    create_gui()
            
# 終了

