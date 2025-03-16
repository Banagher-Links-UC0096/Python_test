# HYP4850U100-H_parallel_KM-N1_Logger
#   Install Pakages
#       pymodbus.py
#       pyserial.py
#       matplotlib.py
#       pandas.py
#       openpyxl.py
#       seaborn.py
# ----------初期設定
from pymodbus.client import ModbusSerialClient as ModbusClient  # Modbus組込
import csv                                                      # CSVファイルモジュール組込
import datetime                                                 # 時計モジュール組込
from time import sleep                                          # タイマー組込
import tkinter as tk                                            # GUIモジュール組込
from tkinter import filedialog
import threading                                                # スレッド組込
import matplotlib.pyplot as plt                                 # グラフ作成
import matplotlib.ticker as ticker                              # グラフ補助
import matplotlib.animation as animation                        # グラフ補助
import matplotlib.font_manager as fm                            # グラフ補助
import pandas as pd                                             # ファイルモジュール組込
import seaborn as sns


interval1=0                                                     # 待ち時間（秒）+10秒
USB1='COM10'                                                    # HYP ID1 USBポート
USB2='COM11'                                                    # HYP ID2 USBポート
USB3='COM12'                                                    # KM-N1 USBポート


# ----------ハイブリッドインバーターデータパラメーター
# [0:address,1:byte,2:x,3:y,4:type,5:name,6:unit,7:len,8:min,9:max,10:set,11:data...]
p_data=[[0x0213,1, 0, 2,1,"系統電圧"    ,"V(AC)"],#0
        [0x0214,1, 0, 3,1,"系統電流"    ,"A(AC)"],#1
        [0x0215,1, 0, 4,2,"系統周波数"  ,"Hz"],#2
        [0x021e,1, 0, 5,1,"系統充電電流","A(DC)"],#3
        [0x0210,1, 0, 7,8,"機器状態"    ,"",0,0,0,"","起動","待機","初期化","省電力",
        "商用出力","インバーター出力","系統出力","混合出力","-","-","停止","故障"],#4
        [0x0216,1, 0, 8,1,"出力電圧"    ,"V(AC)"],#5
        [0x0219,1, 0, 9,1,"出力電流"    ,"A(AC)"],#6
        [0x0218,1, 0,10,2,"出力周波数"  ,"Hz"],#7
        [0x021b,1, 0,11,0,"負荷有効電力","W"],#8
        [0x021c,1, 0,12,0,"負荷皮相電力","W"],#9
        [0x021f,1, 0,13,0,"負荷率"      ,"%"],#10
        
        [0x010b,1, 0,15,8,"充電状態"    ,"",0,0,0,"","未充電","定電流(CC)充電",
        "定電圧(CV)充電","-","浮遊充電","-","充電中1","充電中2"],#11
        [0x0100,1, 0,16,0,"蓄電池SOC"   ,"%"],#12
        [0x0101,1, 0,17,1,"蓄電池電圧"  ,"V(DC)"],#13
        [0x0102,1, 0,18,1,"蓄電池電流"  ,"A(DC)"],#14
        [0x010e,1, 0,19,0,"充電電力"    ,"W"],#15
        [0x0217,1, 0,20,1,"INV電流 "    ,"A(AC)"],#16
        
        [0x0107,1, 0,22,1,"PV入力電圧"  ,"V(DC)"],#17
        [0x0108,1, 0,23,1,"PV入力電流"  ,"A(DC)"],#18
        [0x0109,1, 0,24,0,"PV入力電力"  ,"W"],#19
        [0x0224,1, 0,25,1,"PV降圧電流"  ,"A(DC)"],#20
           
        [0x0212,1, 0,27,1,"DCバス電圧"  ,"V(DC)"],#21
        [0x0225,1, 0,28,1,"降圧電流"    ,"A(DC)"],#22
        [0x0220,1, 0,29,1,"PVHT温度"    ,"℃"],#23
        [0x0221,1, 0,30,1,"INVHT温度"   ,"℃"],#24
        [0x0222,1, 0,31,1,"Tr温度"      ,"℃"],#25
        [0x0223,1, 0,32,1,"内部温度"    ,"℃"],#26
           
        [0xf02d,1, 5, 2,0,"本日充電量"  ,"Ah"],#27
        [0xf02e,1, 5, 3,0,"本日放電量"  ,"Ah"],#28
        [0xf02f,1, 5, 4,1,"本日発電量"  ,"kWh"],#29
        [0xf030,1, 5, 5,1,"本日消費量"  ,"kWh"],#30
        [0xf03c,1, 5, 6,0,"本日商用充電量","Ah"],#31
        [0xf03d,1, 5, 7,1,"本日商用電力消費量","kWh"],#32
        [0xf03e,1, 5, 8,0,"本日インバーター稼働時間","時間"],#33
        [0xf03f,1, 5, 9,0,"本日バイパス稼働時間","時間"],#34

        [0xf034,2, 5,11,0,"累積充電量"  ,"Ah"],#35
        [0xf036,2, 5,12,0,"累積放電量"  ,"Ah"],#36
        [0xf038,2, 5,13,0,"累積発電量"  ,"kWh"],#37
        [0xf03a,2, 5,14,0,"累積負荷積算電力量","kWh"],#38
        [0xf046,2, 5,15,0,"累積商用充電量","kWh"],#39
        [0xf048,2, 5,16,0,"累積商用負荷電力消費量","kWh"],#40
        [0xf04a,1, 5,17,0,"累積インバーター稼働時間","時間"],#41
        [0xf04b,1, 5,18,0,"累積バイパス稼働時間","時間"]]#42
# ----------オリジナルデータパラメーター
# [0:name,1:x,2:y,3:no,4:type,5:unit]
r_data=[["蓄電池総合電流",5,20,1,1,"A"],["蓄電池総合電力",5,21,2,1,"W"],
        ["PV発電総合電力",5,22,3,3,"W"],["AC入力総合電力",5,23,4,3,"VA"],
        ["AC出力総合電力",5,24,5,3,"W"],["損失電力",5,25,6,0,"W"],
         
        ["総合充電量",5,27,7,0,"Ah"],["総合放電量",5,28,8,0,"Ah"],
        ["総合発電量",5,29,9,1,"kWh"],["総合消費量",5,30,10,1,"kWh"],
        ["総合商用充電量",5,31,11,0,"Ah"],["総合商用消費量",5,32,12,1,"kWh"]]
# ----------表示項目データパラメーター
# [0:name,1:x,2:y,3:size]
t_data=[["日時",4,0,4],["L1側",1,1,1],["L2側",3,1,1],["L1側",6,1,1],
        ["L2側",8,1,1],["AC入力側",0,1,0],["AC出力側",0,6,0],
        ["総合データ",5,19,0],["蓄電池側",0,14,0],["PV入力側",0,21,0],
        ["内部システム",0,26,0],["当日積算データ",5,1,0],
        ["累積積算データ",5,10,0],["KM-N1データ",0,35,0],["入力側",1,35,1],
        ["出力側",3,35,1],["入力側",6,35,1],["出力側",8,35,1],
        ["充電電力",8,21,1],["VA",9,22,4],["VA",9,23,4],["効率",8,24,3],["%",9,25,4]]
# KM-N1データパラメーター
# [0:address,1:byte,2:x,3:y,4:type,5:name,6:unit]
k_data=[[0x0001,1,0,36,1,"電圧1","V"],
        [0x0003,1,0,37,1,"電圧2","V"],
        [0x0005,1,0,38,1,"電圧3","V"],
        [0x0007,1,0,39,3,"電流1","A"],
        [0x0009,1,0,40,3,"電流2","A"],
        [0x000b,1,0,41,3,"電流3","A"],
        [0x000d,1,0,42,2,"力率",""],
        [0x000f,1,0,43,1,"周波数","Hz"],
        [0x0011,1,0,44,1,"有効電力","W"],
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
# ----------ハイブリッドインバーター設定パラメーター
# [0:address,1:byte,2:x,3:y,4:type,5:name,6:unit,7:len,8:min,9:max,10:set,11:data...]
s_data=[[0xe004,1, 0, 2,8,"蓄電池タイプ"    ,"",14,0,13,"設定08","ユーザー設定","密閉型鉛",
        "開放型鉛","ゲル型鉛","LFPx14","LEPx15","LFPx16","LFPx7","LFPx8","LFPx8",
        "NCAx7","NCAx8","NCAx13","NCAx14"],#0
        [0xe20f,1, 0, 3,8,"充電モード"      ,"",4,0,3,"設定06","PV優先","系統優先",
        "ハイブリッド","PV専用"],#1
        [0xe002,1, 0, 4,1,"蓄電池容量"      ,"Ah",900,0,900,""],#2
        [0xe005,1, 0, 5,4,"過電圧停止電圧"  ,"V",50,100,150,""],#3
        [0xe006,1, 0, 6,4,"充電上限電圧"    ,"V",50,100,150,""],#4
        [0xe009,1, 0, 7,4,"CV充電電圧"      ,"V",50,100,150,"設定11"],#5
        [0xe01c,1, 0, 8,1,"充電停止電流"    ,"A",100,0,100,"設定57"],#6
        [0xe008,1, 0, 9,4,"CC充電電圧"      ,"V",50,100,150,"設定09"],#7
        [0xe20a,1, 0,10,1,"最大充電電流"    ,"A",600,0,600,"設定07"],#8
        [0xe012,1, 0,11,0,"CC充電遅延時間"  ,"min",120,0,120,"設定10"],#9
        [0xe022,1, 0,12,4,"INV切替電圧"     ,"V",50,100,150,"設定05"],#10
        [0xe00a,1, 0,13,4,"充電再開電圧"    ,"V",50,100,150,"設定37"],#11
        [0xe00b,1, 0,14,4,"低電圧復帰電圧"  ,"V",50,100,150,"設定35"],#12
        [0xe00c,1, 0,15,4,"低電圧警告電圧"  ,"V",50,100,150,"設定14"],#13
        [0xe01b,1, 0,16,4,"バイパス切替電圧","V",50,100,150,"設定04"],#14
        [0xe00d,1, 0,17,4,"過放電遅延オフ電圧","V",50,100,150,"設定12"],#15
        [0xe010,1, 0,18,0,"過放電遅延オフ時間","s",60,0,60,"設定13"],#16
        [0xe00e,1, 0,19,4,"放電停止電圧"    ,"V",50,100,150,"設定15"],#17

        [0xe206,1, 0,20,6,"均等充電有無"    ,"",2,0,1,"設定16"],#18
        [0xe007,1, 0,21,4,"均等充電電圧"    ,"V",50,100,150,"設定17"],#19
        [0xe011,1, 0,22,0,"均等充電時間"    ,"min",120,0,120,"設定18"],#20
        [0xe023,1, 0,23,0,"均等充電遅延"    ,"min",120,0,120,"設定19"],#21
        [0xe013,1, 0,24,0,"均等充電間隔"    ,"day",7,0,7,"設定20"],#22
        [0xdf0d,1, 0,25,6,"均等充電有無"    ,"",2,0,1,"設定21"],#23
          
        [0xe215,1, 5, 2,8,"BMS通信"         ,"",3,0,2,"設定32","OFF","RS485-BMS","CAN-BMS"],#24
        [0xe21b,1, 5, 3,8,"BMSプロトコル"   ,"",18,0,17,"設定33","Pace","Rata","Allgrand",
        "Oliter","PCT","Sunwoda","Dyness","SRNE","Pylontech","","","","","","","",
        "WS Technicals","Uz Energy"],#25
        [0xe025,1, 5, 4,8,"充電制御"        ,"",3,0,2,"設定39","OFF","BMS制御","INV制御"],#26
        [0xe01e,1, 5, 5,0,"過放電警報SOC"   ,"%",100,0,100,"設定58"],#27
        [0xe00f,1, 5, 6,0,"放電停止SOC"     ,"%",100,0,100,"設定59"],#28
        [0xe01d,1, 5, 7,0,"充電停止SOC"     ,"%",100,0,100,"設定60"],#29
        [0xe01f,1, 5, 8,0,"バイパス切替SOC" ,"%",100,0,100,"設定61"],#30
        [0xe020,1, 5, 9,0,"INV切替SOC"      ,"%",100,1,100,"設定62"],#31
          
        [0xe11f,1, 5,10,0,"PV最大電圧"      ,"V",500,0,500,""],#32
        [0xe001,1, 5,11,1,"PV最大充電電流"  ,"A",800,0,800,"設定36"],#33
        [0xe120,1, 5,12,1,"最大充電電流"    ,"A",1000,0,1000,""],#34
          
        [0xe204,1, 5,13,8,"出力優先"        ,"",3,0,2,"設定01","PV優先","系統優先","蓄電池優先"],#35
        [0xe209,1, 5,14,2,"AC出力周波数"    ,"Hz",2,5000,6000,"設定02"],#36
        [0xe208,1, 5,15,1,"AC出力電圧"      ,"V",20,100,120,"設定38"],#37
        [0xe129,1, 5,16,1,"最大出力電流"    ,"A",420,0,420,""],#38
        [0xe118,1, 5,17,1,"定格出力"        ,"kW",100,0,100,""],#39
                 
        [0xe20b,1, 5,18,8,"AC入力電圧範囲"  ,"",2,0,1,"設定03","APL(90-280V)","UPS(90-140V)"],#40
        [0xe201,1, 5,19,8,"並列運転モード"  ,"",8,0,7,"設定31","単相","並列","2P0","2P1","2P2",
         "3P1","3P2","3P3"],#41
        [0xe205,1, 5,20,1,"AC充電最大電流"  ,"A",400,0,400,"設定28"],#42
        [0xe037,1, 5,21,8,"ハイブリッド出力","",3,0,2,"設定34","OFF","AC入力PV出力",
         "AC出力PV出力"],#43
          
        [0xe20c,1, 5,22,6,"省エネモード"    ,"",2,0,1,"設定22"],#44
        [0xe212,1, 5,23,6,"バイパス出力有無","",2,0,1,"設定27"],#45
        [0xe200,1, 5,24,0,"RS485アドレス"   ,"",240,1,240,"設定30"],#46
        [0xe214,1, 5,25,6,"分相変圧器"      ,"",2,0,1,"設定29"],#47
        [0xe207,1, 5,26,6,"N相出力"         ,"",2,0,1,"設定63"],#48
        [0xe038,1, 5,27,6,"漏電検知機能"    ,"",2,0,1,"設定56"],#49

        [0xe213,1, 5,28,6,"障害記録"        ,"",2,0,1,""],#50
        [0xe210,1, 5,29,6,"警報音有無"      ,"",2,0,1,"設定25"],#51
        [0xe211,1, 5,30,6,"電源切替警報有無","",2,0,1,"設定26"],#52
        [0xe20d,1, 5,31,6,"過負荷停止再起動","",2,0,1,"設定23"],#53
        [0xe20e,1, 5,32,6,"高温停止再起動"  ,"",2,0,1,"設定24"],#54

        #["設定54,55",""],
        [0xe026,1,10, 2,7,"充電開始時間１"  ,"",0,0,0,"設定40"],#55
        [0xe027,1,10, 3,7,"充電終了時間１"  ,"",0,0,0,"設定41"],#56
        [0xe028,1,10, 4,7,"充電開始時間２"  ,"",0,0,0,"設定42"],#57
        [0xe029,1,10, 5,7,"充電終了時間２"  ,"",0,0,0,"設定43"],#58
        [0xe02a,1,10, 6,7,"充電開始時間３"  ,"",0,0,0,"設定44"],#59
        [0xe02b,1,10, 7,7,"充電終了時間３"  ,"",0,0,0,"設定45"],#60
        [0xe02c,1,10, 8,6,"充電時間設定"    ,"",2,0,1,"設定46"],#61
        [0xe02d,1,10, 9,7,"放電開始時間１"  ,"",0,0,0,"設定47"],#62
        [0xe02e,1,10,10,7,"放電終了時間１"  ,"",0,0,0,"設定48"],#63
        [0xe02f,1,10,11,7,"放電開始時間２"  ,"",0,0,0,"設定49"],#64
        [0xe030,1,10,12,7,"放電終了時間２"  ,"",0,0,0,"設定50"],#65
        [0xe031,1,10,13,7,"放電開始時間３"  ,"",0,0,0,"設定51"],#66
        [0xe032,1,10,14,7,"放電終了時間３"  ,"",0,0,0,"設定52"],#67
        [0xe033,1,10,15,6,"放電時間設定"    ,"",0,0,0,"設定53"]]#68

# ----------Modbus接続設定
client1=ModbusClient(framer="rtu",port=USB1,                  # Hybrid1（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=5)
client2=ModbusClient(framer="rtu",port=USB2,                  # Hyblid2（Windows）
                       baudrate=9600,bytesize=8,stopbits=1,parity='N',timeout=5)
client3=ModbusClient(framer="rtu",port=USB3,                  # KM-N112（Windows）
                       baudrate=9600,bytesize=8,stopbits=2,parity='N',timeout=5)
# ----------Modbusデータ取得
def hybrid_modbus_read(slave_add,slave_count,err):                  # ファンクション03h
    try:
        err=[]
        if client1.connect():                                   # Hybrid ID1 ポート接続
            read_data1=client1.read_holding_registers(
                        address=slave_add,count=slave_count,slave=1)
            if read_data1.isError():
                read_data1=[0,0]
                err.append("Hybrid1 read error."+hex(slave_add))
                print(err)
            else:read_data1=read_data1.registers
        else:
            err.append("Hybrid1 No connect.")
            read_data1,read_data2=test_hybrid(slave_add)
        if client2.connect():                                   # Hybrid ID1 ポート接続
            read_data2=client2.read_holding_registers(
                        address=slave_add,count=slave_count,slave=2)
            if read_data2.isError():
                read_data2=[0,0]
                err.append("Hybrid2 read error."+hex(slave_add))
                print(err)
            else:read_data2=read_data2.registers
        else:
            err.append("Hybrid2 No connect.")
            read_data1,read_data2=test_hybrid(slave_add)
    except FileNotFoundError as e:
        err.append(str(e))
        read_data1,read_data2=test_hybrid(slave_add)
    except Exception as e:
        err.append(str(e))
        read_data1,read_data2=test_hybrid(slave_add)
    finally:                                                    # ポート切断
        #print(err)
        client1.close()  
        client2.close()
    return read_data1,read_data2,err
# ----------未接続時テスト用
def test_hybrid(slave_add):
    test_data=[ [0x010b,[1],[2]],[0x0100,[53],[54]],[0x0101,[524],[525]],[0x0102,[10],[10]],[0x010e,[241],[242]],
                [0x0107,[1727],[1728]],[0x0108,[14],[15]],[0x0109,[241],[242]],[0x0224,[4],[5]],
                [0x0210,[5],[3]],[0x0216,[999],[1000]],[0x0217,[23],[24]],[0x0218,[5999],[6000]],[0x0219,[25],[25]],
                [0x021b,[741],[742]],[0x021c,[761],[762]],[0x021f,[5],[6]],[0x0225,[20],[21]],
                [0x0212,[3930],[3931]],[0x0213,[0],[1000]],[0x0214,[0],[500]],[0x0215,[0],[6001]],[0x021e,[0],[30]],
                [0x0220,[410],[411]],[0x0221,[404],[405]],[0x0222,[537],[538]],[0x0223,[549],[560]],
                [0xf02d,[28],[29]],[0xf02e,[11],[12]],[0xf02f,[17],[18]],[0xf030,[16],[17]],
                [0xf03c,[30],[31]],[0xf03d,[16],[17]],[0xf03e,[6],[7]],[0xf03f,[0],[0]],
                [0xf034,[4079,0],[4080,0]],[0xf036,[2488,0],[2499,0]],[0xf038,[3061,0],[3090,0]],[0xf03a,[2022,0],[2080,0]],
                [0xf046,[991,0],[1000,0]],[0xf048,[825,0],[900,0]],[0xf04a,[752],[753]],[0xf04b,[83],[84]],

                [0xe004,[6],[0]],[0xe20f,[0],[0]],[0xe002,[300],[0]],[0xe005,[150],[0]],[0xe006,[149],[0]],[0xe009,[148],[0]],[0xe01c,[20],[0]],
                [0xe008,[147],[0]],[0xe20a,[1000],[0]],[0xe012,[120],[0]],[0xe022,[146],[0]],[0xe00a,[145],[0]],[0xe00b,[128],[0]],[0xe00c,[127],[0]],
                [0xe01b,[120],[0]],[0xe00d,[110],[0]],[0xe010,[60],[0]],[0xe00e,[100],[0]],[0xe206,[0],[0]],[0xe007,[144],[0]],[0xe011,[120],[0]],
                [0xe023,[120],[0]],[0xe013,[7],[0]],[0xdf0d,[0],[0]],[0xe215,[1],[0]],[0xe21b,[7],[0]],[0xe025,[1],[0]],[0xe01e,[30],[0]],
                [0xe00f,[5],[0]],[0xe01d,[100],[0]],[0xe01f,[10],[0]],[0xe020,[20],[0]],[0xe11f,[500],[0]],[0xe001,[1000],[0]],[0xe120,[1000],[0]],
                [0xe204,[2],[0]],[0xe209,[6000],[0]],[0xe208,[100],[0]],[0xe129,[420],[0]],[0xe118,[500],[0]],[0xe20b,[1],[0]],[0xe201,[2],[0]],
                [0xe205,[300],[0]],[0xe037,[2],[0]],[0xe20c,[0],[0]],[0xe212,[1],[0]],[0xe200,[1],[0]],[0xe214,[0],[0]],[0xe207,[0],[0]],[0xe038,[0],[0]],
                [0xe213,[1],[0]],[0xe210,[0],[0]],[0xe211,[0],[0]],[0xe20d,[1],[0]],[0xe20e,[0],[0]],
                [0xe026,[23,0],[0,0]],[0xe027,[7,0],[0,0]],[0xe028,[0,0],[0,0]],[0xe029,[0,0],[0,0]],[0xe02a,[0,0],[0,0]],[0xe02b,[0,0],[0,0]],[0xe02c,[0],[0]],
                [0xe02d,[7,0],[0,0]],[0xe02e,[23,0],[0,0]],[0xe02f,[0,0],[0,0]],[0xe030,[0,0],[0,0]],[0xe031,[0,0],[0,0]],[0xe032,[0,0],[0,0]],[0xe033,[0],[0]]]
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
def kmn1_modbus_read(slave_add,slave_count,err):                    # ファンクション03h
    try:
        if client3.connect():                                   # Hybrid ID1 ポート接続
            read_data1=client3.read_holding_registers(
                        address=slave_add,count=slave_count,slave=1)
            read_data2=client3.read_holding_registers(
                        address=slave_add,count=slave_count,slave=2)
            if read_data1.isError():
                read_data1=[0,0]
                err.append("KM-N1 read1 error."+hex(slave_add))
            else:read_data1=read_data1.registers
            if read_data2.isError():
                read_data2=[0,0]
                err.append("KM-N1 read2 error."+hex(slave_add))
            else:read_data2=read_data2.registers
        else:
            err.append("KM-N1 No connect.")
            read_data1,read_data2=test_km(slave_add)
    except FileNotFoundError as e:
        err.append(str(e))
        read_data1,read_data2=test_km(slave_add)
    except Exception as e:
        err.append(str(e))
        read_data1,read_data2=test_km(slave_add)
    finally:                                                    # ポート切断
        #print(err)
        client3.close()  
    
    return read_data1,read_data2,err

# ----------未接続時テスト用
def test_km(slave_add):
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
    if d_type==6:d=["On","Off"][rdd]                            # ON/OFF判定  
    if d_type==7:                                               # 時刻判定
        ym=hex(rdd)[4:].zfill(4)
        d=str(int(ym[0:2],16)).zfill(2)+":"+str(int(ym[2:4],16)).zfill(2)
    if d_type==8:d=n_data[d_add][rdd+11]                        # 拡張判定
    if d_type==9:print(rdd)
        
    return d

def data_read(n_data):                                          # データ処理
    dt_now=datetime.datetime.now()                              # 日時を取得
    date_time=dt_now.strftime('%y/%m/%d %H:%M:%S')
    hiwd1,hiwd2,hird1,hird2,o_data=[],[],[],[],[]
    kmwd1,kmwd2,kmrd1,kmrd2,err=[],[],[],[],[]
    n_list=len(n_data)
    for a in range(n_list):
        byte=n_data[a][1]
        d_type=n_data[a][4]
        r_data1,r_data2=[],[]
        r_data1,r_data2,err=hybrid_modbus_read(n_data[a][0],byte,err)
        #print(n_list,a,hex(n_data[a][0]),r_data1)
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
        #print(data1,data2)
        hiwd1.append(data1)
        hiwd2.append(data2)
        hird1.append(r_data1)
        hird2.append(r_data2)
    for a in range(len(k_data)):
        byte=k_data[a][1]
        d_type=k_data[a][4]
        r_data1,r_data2,err=kmn1_modbus_read(k_data[a][0],2,err)
        if byte==1:                                             # 16bitデータ変換
            data_list1=[r_data1[0],r_data1[0]/10,r_data1[0]/100,r_data1[0]/1000]
            data1=data_list1[d_type]
            data_list2=[r_data2[0],r_data2[0]/10,r_data2[0]/100,r_data2[0]/1000]
            data2=data_list2[d_type]            
        else:
            data1=change_type(d_type,
                    int(str(int(hex(r_data1[1])[2:].zfill(2)[0:4],16))
                    +str(int(hex(r_data1[0])[2:].zfill(2)[0:4],16)))
                    ,sys_volt1,byte,a,k_data)
            data2=change_type(d_type,
                    int(str(int(hex(r_data2[1])[2:].zfill(2)[0:4],16))
                    +str(int(hex(r_data2[0])[2:].zfill(2)[0:4],16)))
                    ,sys_volt1,byte,a,k_data)
        kmwd1.append(data1)
        kmwd2.append(data2)
        kmrd1.append(r_data1)
        kmrd2.append(r_data2)
    csv_data1=[date_time]
    csv_data2=[date_time]
    csv_data3=[date_time]
    csv_data1=csv_data1+hiwd1
    csv_data2=csv_data2+hiwd2
    csv_data3=csv_data3+kmwd1+kmwd2
    if n_data==p_data:
        q_rdd=[change_minus(hird1[14][0]),change_minus(hird2[14][0]),
               change_minus(hird1[8][0]),change_minus(hird2[8][0])]
        i_calc=((hird1[0][0]*hird1[1][0])+(hird2[0][0]*hird2[1][0]))/100        # AC入力電力
        o_calc=(q_rdd[2]+q_rdd[3])                                              # AC出力電力(有効)
        p_calc=(hird1[19][0]+hird2[19][0])                                      # PV入力電力
        b_calc=round(((hird1[13][0]*q_rdd[0]+hird2[13][0]*q_rdd[1])/100),1)     # バッテリー電力
        b_curr=((q_rdd[0]+q_rdd[1])/10)                                         # バッテリー電流
        s_calc=round((i_calc-o_calc+b_calc+p_calc),1)                           # 消費電力
        p_powr=round(((hird1[13][0]*hird1[20][0]+hird2[13][0]*hird1[20][0])/100),1)# PV充電電力
        g_powr=round(((hird1[13][0]*hird1[3][0]+hird2[13][0]*hird1[3][0])/100),1)# AC充電電力
        if b_calc<0 :o_powr=round((b_calc+o_calc)/(i_calc+p_calc)*100,1)
        else:o_powr=round(o_calc/(s_calc+o_calc)*100,1)                         # 効率
        calc_list=[b_curr,b_calc,p_calc,i_calc,o_calc,s_calc,
                hird1[27][0]+hird2[27][0],hird1[28][0]+hird2[28][0],
                (hird1[29][0]+hird2[29][0])/10,(hird1[30][0]+hird2[30][0])/10,
                hird1[31][0]+hird2[31][0],(hird1[32][0]+hird2[32][0])/10,
                p_powr,g_powr,o_powr]
        for s in range(len(calc_list)):o_data.append(calc_list[s])
    writer.writerow(hird1+hird2+kmrd1+kmrd2)
    
    return date_time,hiwd1,hiwd2,kmwd1,kmwd2,o_data,csv_data1,csv_data2,csv_data3,err

# ----------モニター画面
def create_gui():                                               # GUI作成
    root=tk.Tk()
    root.geometry('830x950+20+20')                              # ウインドウサイズ
    root.title("HYP4850U100-H 並列モニター")                     # ウインドウタイトル
    frame=tk.Frame(root)
    frame.grid(row=0,column=0,sticky=tk.NSEW,padx=5,pady=10)
    wid=[19,13,5,13,5]
    col1='#0000ff'                                              # データ文字色
    col2='#cccccc'                                              # データ背景色
    p_list,t_list,r_list,k_list=len(p_data),len(t_data),len(r_data),len(k_data)
    err=[]
    labels=[tk.Label(frame,width=wid[t_data[i][3]],text=t_data[i][0]#,font=("MS Gothic",9,)
                        ,anchor=tk.W)for i in range(t_list)]    
    [labels[h].grid(column=t_data[h][1],row=t_data[h][2])for h in range(t_list)]
    label0=tk.Label(frame,width=wid[0],text=date_time,#font="bold",# 種別表示"bold"太字
                    anchor=tk.W,borderwidth=1)
    label0.grid(column=5,row=0)                                 # ログ日時表示
    label1=tk.Label(frame,width=wid[0],text=err,#font="bold",# 種別表示"bold"太字
                    anchor=tk.W,borderwidth=1)
    label1.grid(column=0,row=0) 
    n_data,n_list=[p_data,k_data],[p_list,k_list]
    for j in range(2):
        for k in range(2):
            labels=[tk.Label(frame,width=wid[0],text=n_data[j][i][5]
                        ,anchor=tk.W)for i in range(n_list[j])] # 項目表示
            [labels[h].grid(column=n_data[j][h][2],row=n_data[j][h][3])for h in range(n_list[j])]
            labels=[tk.Label(frame,width=wid[4],text=n_data[j][i][6]
                            ,anchor=tk.W)for i in range(n_list[j])]# 単位表示
            [labels[h].grid(column=n_data[j][h][2]+2+2*k,row=n_data[j][h][3])for h in range(n_list[j])]
    labels1=[tk.Label(frame,width=wid[1],text=hiwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1# ID1データ表示
                      ,foreground=col1,background=col2)for x in range(p_list)]
    [labels1[h].grid(column=p_data[h][2]+1,row=p_data[h][3])for h in range(p_list)]
    labels2=[tk.Label(frame,width=wid[3],text=hiwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1# ID2データ表示
                      ,foreground=col1,background=col2)for x in range(p_list)]
    [labels2[h].grid(column=p_data[h][2]+3,row=p_data[h][3])for h in range(p_list)]
    labels4=[tk.Label(frame,width=wid[1],text=kmwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1# km1データ表示
                      ,foreground=col1,background=col2)for x in range(k_list)]
    [labels4[h].grid(column=k_data[h][2]+1,row=k_data[h][3])for h in range(k_list)]
    labels5=[tk.Label(frame,width=wid[3],text=kmwd1[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1# km2データ表示
                      ,foreground=col1,background=col2)for x in range(k_list)]
    [labels5[h].grid(column=k_data[h][2]+3,row=k_data[h][3])for h in range(k_list)]
    labels=[tk.Label(frame,width=wid[0],text=r_data[i][0]
                        ,anchor=tk.W)for i in range(r_list)]    # 追加項目表示
    [labels[h].grid(column=r_data[h][1],row=r_data[h][2])for h in range(r_list)]
    labels3=[tk.Label(frame,width=wid[1],text=o_data[x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1# 追加データ表示1
                      ,foreground=col1,background=col2)for x in range(r_list)]
    [labels3[h].grid(column=r_data[h][1]+1,row=r_data[h][2])for h in range(r_list)]

    labels6=[tk.Label(frame,width=wid[1],text=o_data[r_list+x]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1# 追加データ表示2
                      ,foreground=col1,background=col2)for x in range(2)]
    [labels6[h].grid(column=r_data[h][1]+3,row=r_data[h][2]+2)for h in range(2)]
    labels7=tk.Label(frame,width=wid[1],text=o_data[len(o_data)-1]
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1# 追加データ表示3
                      ,foreground=col1,background=col2)
    labels7.grid(column=8,row=25)
    
    labels=[tk.Label(frame,width=wid[4],text=r_data[i][5]
                        ,anchor=tk.W)for i in range(r_list)]    # 追加単位表示
    [labels[h].grid(column=r_data[h][1]+2,row=r_data[h][2])for h in range(r_list)]

    button1=tk.Button(frame,text="計測終了",command=root.destroy)# ループ終了
    button1.grid(column=8,row=0)
    button2=tk.Button(frame,text="機器設定",command=root.destroy)    # ループ終了
    button2.grid(column=6,row=0)
    button3=tk.Button(frame,text="チャート表示",command=root.destroy)  # チャート移動
    button3.grid(column=3,row=0)

    thread1=threading.Thread(target=update_data,args=([label0,
                                    [labels1[x]for x in range(p_list)],
                                    [labels2[x]for x in range(p_list)],
                                    [labels3[x]for x in range(r_list)],
                                    [labels4[x]for x in range(k_list)],
                                    [labels5[x]for x in range(k_list)],
                                    [labels6[x]for x in range(2)],
                                    labels7,label1
                                    ]))

    thread1.daemon=True                                         # スレッド終了
    thread1.start()                                             # スレッド処理開始
    root.mainloop()                                             # メインループ開始

# ----------データ更新処理
def update_data(label0,labels1,labels2,labels3,labels4,labels5,labels6,labels7,label1):# モニター画面更新＆ログ更新
        while True:
            p_list,k_list=len(p_data),len(k_data)
            date_time,hiwd1,hiwd2,kmwd1,kmwd2,o_data,csv_data1,csv_data2,csv_data3,err=data_read(p_data)
            writer1.writerow(csv_data1)
            writer2.writerow(csv_data2)
            writer3.writerow(csv_data3)
            label0.config(text=f"{date_time}")                  # GUIデータ更新
            [labels1[x].config(text=f"{hiwd1[x]}")for x in range(p_list)]
            [labels2[x].config(text=f"{hiwd2[x]}")for x in range(p_list)]
            [labels3[x].config(text=f"{o_data[x]}")for x in range(len(r_data))]
            [labels4[x].config(text=f"{kmwd1[x]}")for x in range(k_list)]
            [labels5[x].config(text=f"{kmwd2[x]}")for x in range(k_list)]
            [labels6[x].config(text=f"{o_data[len(o_data)-3+x]}")for x in range(2)]
            labels7.config(text=f"{o_data[len(o_data)-1]}")
            label1.config(text=f"{err}") 
            sleep(interval1+4)                                    # インターバルタイマー

# ----------設定モニター画面
def setting():                                                  # 設定画面作成
    root1=tk.Tk()
    root1.geometry('1000x720+861+20')                           # ウインドウサイズ
    root1.title("HYP4850U100-H Setting Monitor")                # ウインドウタイトル
    frame1=tk.Frame(root1)
    frame1.grid(row=0,column=0,sticky=tk.NSEW,padx=5,pady=10)
    n_list,n_data=len(s_data),s_data
    date_time,hiwd1,hiwd2,kmwd1,kmwd2,o_data,csv_data1,csv_data2,csv_data3=data_read(n_data)
    wid=[15,5,15,5]    
    col1='#0000ff'                                              # データ文字色
    col2='#cccccc'                                              # データ背景色
    labels=[tk.Label(frame1,width=wid[0],text=n_data[i][5]      # 設定名
                     ,anchor=tk.W)for i in range(n_list)]
    [labels[h].grid(column=n_data[h][2],row=n_data[h][3])for h in range(n_list)]
    labels=[tk.Label(frame1,width=wid[1],text=n_data[i][10]     # 設定番号
                     ,anchor=tk.W)for i in range(n_list)]
    [labels[h].grid(column=n_data[h][2]+1,row=n_data[h][3])for h in range(n_list)]
    labels1=[tk.Label(frame1,width=wid[2],text=hiwd1[x]         # 設定値
                      ,anchor=tk.E,relief=tk.SOLID,borderwidth=1
                      ,foreground=col1,background=col2)for x in range(n_list)]
    [labels1[h].grid(column=n_data[h][2]+2,row=n_data[h][3])for h in range(n_list)]
    labels=[tk.Label(frame1,width=wid[3],text=n_data[i][6]      # 単位
                        ,anchor=tk.W)for i in range(n_list)]
    [labels[h].grid(column=n_data[h][2]+3,row=n_data[h][3])for h in range(n_list)]

    button1=tk.Button(frame1,text="終了",command=root1.destroy) # ループ終了
    button1.grid(column=14,row=0)

    root1.mainloop()


def select_file():                                              # ファイル選択
    root3 = tk.Tk()
    root3.withdraw()                                            # メインウィンドウを表示しない
    file_path = filedialog.askopenfilename(
        filetypes=[('csv files','*.csv'),('elsx files','*.xlsx')])      # ファイル選択ダイアログを表示
    file_name = file_path.split('/')[-1]                        # ファイル名を取得
    print(f"選択されたファイル: {file_path}")                     # 選択されたファイルのパスとファイル名を表示
    print(f"ファイル名: {file_name}")
    return file_name



def chart():                                                    # チャート画面作成 
    file_path = select_file()    

    plt.figure(figsize=(15,8))                                  # フレームサイズ
    plt.xticks(rotation=80)                                     # x軸表示角度変更
    plt.yticks(rotation=-10)                                    # y軸表示角度変更
    font_path='c:/windows/Fonts/meiryo.ttc'
    jp_font=fm.FontProperties(fname=font_path)
    sns.set(style='darkgrid')                                   # デフォルト
    sns.set_palette('winter_r') 

    # ウィンドウの位置を指定
    manager = plt.get_current_fig_manager()
    manager.window.state("zoomed")  # フルスクリーン設定
    if file_path=='*.elsx':
        data0=pd.read_excel(file_path,usecols=[0])#time
        data1=pd.read_excel(file_path,usecols=[14])#pv_power
        data2=pd.read_excel(file_path,usecols=[22])#output_power
        data3=pd.read_excel(file_path,usecols=[3,4])
        data3.columns = ['Column3', 'Column4']
        data3['Column3'] = pd.to_numeric(data3['Column3'], errors='coerce')
        data3['Column4'] = pd.to_numeric(data3['Column4'], errors='coerce')
        data3['Product']=data3['Column3']*data3['Column4']
        data4=pd.read_excel(file_path,usecols=[8,9])
        data4.columns = ['Column8', 'Column9']
        data4['Column8'] = pd.to_numeric(data4['Column8'], errors='coerce')
        data4['Column9'] = pd.to_numeric(data4['Column9'], errors='coerce')
        data4['Product']=data4['Column8']*data4['Column9']
    else:
        data0=pd.read_csv(file_path,usecols=[0])#time
        data1=pd.read_csv(file_path,usecols=[20])#pv_power
        data2=pd.read_csv(file_path,usecols=[9])#output_power
        data3=pd.read_csv(file_path,usecols=[14,15])
        data3.columns = ['Column3', 'Column4']
        data3['Column3'] = pd.to_numeric(data3['Column3'], errors='coerce')
        data3['Column4'] = pd.to_numeric(data3['Column4'], errors='coerce')
        data3['Product']=data3['Column3']*data3['Column4']
        data4=pd.read_csv(file_path,usecols=[1,2])
        data4.columns = ['Column8', 'Column9']
        data4['Column8'] = pd.to_numeric(data4['Column8'], errors='coerce')
        data4['Column9'] = pd.to_numeric(data4['Column9'], errors='coerce')
        data4['Product']=data4['Column8']*data4['Column9']
    
    plt.plot(data0.iloc[:,0], data1.iloc[:,0],linewidth=2
             , label='PV_power',linestyle='-',color='red')      # プロット、凡用例ラベル
    plt.plot(data0.iloc[:,0], data2.iloc[:,0],linewidth=1
             , label='Inv_power',linestyle='-',color='blue')    # プロット、凡用例ラベル
    plt.plot(data0.iloc[:,0], data3.iloc[:,2],linewidth=2
             , label='Batt_power',linestyle='-',color='green')  # プロット、凡用例ラベル
    plt.plot(data0.iloc[:,0], data4.iloc[:,2],linewidth=1
             , label='grid_power',linestyle='-',color='brown')  # プロット、凡用例ラベル
    
    fig=plt.gcf()
    fig.canvas.manager.set_window_title('PV')
    plt.title('電力量', fontsize=20,fontproperties=jp_font)      # グラフタイトル
    plt.ylabel('電力', fontsize=14,fontproperties=jp_font)       # x軸ラベル
    plt.xlabel('日時', fontsize=14,fontproperties=jp_font)       # y軸ラベル
    plt.grid(True)
    plt.gca().xaxis.set_major_locator(ticker.MultipleLocator(10))# x軸の間隔設定
    plt.gca().yaxis.set_major_locator(ticker.MultipleLocator(500))# x軸の間隔設定
    plt.legend(bbox_to_anchor=(1, 1), loc='upper left')         # 凡例の位置設定

    plt.show()
    return

# ----------CSVファイル設定
err=[]
sys_volt1,sys_volt2,err=hybrid_modbus_read(0xe003,1,err)                # システム電圧読込
sys_id1,sys_id2,err=hybrid_modbus_read(0x0035,20,err)                   # プロダクトID読込
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
                [writer2.writerow([n_data2,u_data2][i])for i in range(2)]
                with open(file_name3,'w',newline='')as file3:
                    writer3=csv.writer(file3)
                    [writer3.writerow([n_data3,u_data3][i])for i in range(2)]
                    date_time,hiwd1,hiwd2,kmwd1,kmwd2,o_data,csv_data1,csv_data2,csv_data3,err=data_read(p_data)
                    writer1.writerow(csv_data1)
                    writer2.writerow(csv_data2)
                    writer3.writerow(csv_data3)
                    if __name__ == "__main__":
                        create_gui()
            
# 終了

