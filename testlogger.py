import serial
import csv
import time
import struct

# COMポートの設定
COM5_PORT = 'COM5'
COM6_PORT = 'COM6'
BAUD_RATE = 9600  # ボーレートは必要に応じて変更

# CSVファイルの設定
LOG_FILE = 'serial_log.csv'

# シリアルポートの初期化
ser_com5 = serial.Serial(COM5_PORT, BAUD_RATE, timeout=1)
try:
    ser_com6 = serial.Serial(COM6_PORT, BAUD_RATE, timeout=1)
    com6_connected = True
except serial.SerialException:
    print("COM6 is not connected. Using dummy data.")
    com6_connected = False

# アドレスとデータフレームのリスト
modbus_data = {
    0x000b: b'\x00\x04',
    0x0100: b'\x00\x35\x02\x0c\00\x0a',
    0x0107: b'\x06\xbf\x00\x0e\x00\xf1\x00\x00\x00\x01\x00\x00\x00\x00\x00\xf1\x00\x00',
    0x0204: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x17\x05\x01\x0e\x29\x37\x00\x00\x00\x05\xff\xff',
    0x0212: b'\x0f\x5a\x00\x00\x00\x00\x00\x00\x03\xe7\x00\x17\x17\x6f\x00\x19\x00\x00\x02\xe5\x02\xf9\x00\x00\x00\x00\x00\x05\x01\x9a\x01\x94\x02\x19\x02\x25\x00\x04\x00\x14',
    0xe004: b'\x00\x06', 
    0xe116: b'\x00\x22',
    0xf000: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf007: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf00e: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf015: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf01c: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf023: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf02d: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf034: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
    0xf040: b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00',
}

def calculate_crc(data):
    """
    Modbus CRC16を計算する関数。
    """
    crc = 0xFFFF
    for pos in data:
        crc ^= pos
        for _ in range(8):
            if crc & 0x0001:
                crc >>= 1
                crc ^= 0xA001
            else:
                crc >>= 1
    return struct.pack('<H', crc)  # リトルエンディアンでCRCを返す

def generate_modbus_response(request_frame):
    """
    Modbusリクエストフレームを解析し、対応するレスポンスフレームを生成する関数。
    """
    try:
        # Modbusリクエストの解析
        if len(request_frame) < 8:
            return b''  # 不正なフレームは無視

        function_code = request_frame[1]  # ファンクションコード
        register_address = int.from_bytes(request_frame[2:4], byteorder='big')  # レジスタアドレス

        # デバイスアドレスを固定で1に設定
        device_address = 0x01

        # アドレスに対応するデータを取得
        if function_code == 0x03:  # レジスタ読み取り
            response_data = modbus_data.get(register_address, b'')  # レジスタアドレスに対応するデータを取得
            if not response_data:
                print(f"Register Address 0x{register_address:04X} not found in modbus_data.")
                return b''  # データが存在しない場合は空レスポンス

            print(f"Register Address 0x{register_address:04X} found. Preparing response.")

            # データ長を計算（データのバイト数）
            data_length = len(response_data)

            # レスポンスフレームを生成
            response_frame = bytes([device_address, function_code, data_length]) + response_data

            # CRCを計算してフレームに追加
            crc = calculate_crc(response_frame)
            full_response = response_frame + crc

            print(f"Generated Response: {full_response.hex()}")
            return full_response
        else:
            # サポートされていないファンクションコードの場合
            print(f"Unsupported function code: {function_code}")
            return b''

    except Exception as e:
        print(f"Error generating Modbus response: {e}")
        return b''

try:
    print("Listening on COM5 and COM6...")
    while True:
        # COM5からデータを受信
        if ser_com5.in_waiting > 0:
            data_from_com5 = ser_com5.read(ser_com5.in_waiting)
            print(f"Received from COM5: {data_from_com5.hex()}")

            # レジスタアドレスを抽出して16進数で表示
            if len(data_from_com5) >= 4:  # レジスタアドレスを抽出するには最低4バイト必要
                register_address = int.from_bytes(data_from_com5[2:4], byteorder='big')
                print(f"Extracted Register Address: 0x{register_address:04X}")

            # データをCSVにログ
            with open(LOG_FILE, mode='a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), 'COM5->COM6', data_from_com5.hex()])

            # COM6へ送信またはModbusレスポンスを返す
            if com6_connected:
                ser_com6.write(data_from_com5)
            else:
                # Modbusレスポンスを生成してCOM5に返す
                modbus_response = generate_modbus_response(data_from_com5)
                if modbus_response:
                    print(f"Sending Modbus response to COM5: {modbus_response.hex()}")
                    ser_com5.write(modbus_response)

        # COM6からデータを受信
        if com6_connected and ser_com6.in_waiting > 0:
            data_from_com6 = ser_com6.read(ser_com6.in_waiting)
            print(f"Received from COM6: {data_from_com6.hex()}")

            # データをCSVにログ
            with open(LOG_FILE, mode='a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), 'COM6->COM5', data_from_com6.hex()])

            # COM5へ送信
            ser_com5.write(data_from_com6)

except KeyboardInterrupt:
    print("Terminating...")
finally:
    # シリアルポートを閉じる
    ser_com5.close()
    if com6_connected:
        ser_com6.close()