import serial
import csv
import time
import struct
import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports

def select_com_ports():
    """
    使用可能なCOMポートを選択するウィンドウを表示し、Logger_portとInverter_portを選択する。
    """
    ports = [port.device for port in serial.tools.list_ports.comports()]
    if not ports:
        print("No COM ports available.")
        return None, None

    selected_ports = {"Logger_port": None, "Inverter_port": None}

    def update_com_options(*args):
        """
        Logger_portとInverter_portの選択肢を動的に更新する。
        """
        selected_logger = logger_var.get()
        selected_inverter = inverter_var.get()

        # Logger_portの選択肢を更新（Inverter_portで選択されているポートを除外）
        logger_dropdown['values'] = [port for port in ports if port != selected_inverter]

        # Inverter_portの選択肢を更新（Logger_portで選択されているポートを除外）
        inverter_dropdown['values'] = ["なし"] + [port for port in ports if port != selected_logger]

    def on_submit():
        """
        ポート選択を確定してウィンドウを閉じる。
        """
        selected_ports["Logger_port"] = logger_var.get()
        selected_ports["Inverter_port"] = inverter_var.get()
        root.destroy()

    root = tk.Tk()
    root.title("USB Port Selection")

    # Logger_portの選択
    tk.Label(root, text="Logger Port:", anchor="w").grid(row=0, column=0, padx=10, pady=5)
    logger_var = tk.StringVar(value=ports[0] if ports else "")
    logger_var.trace("w", update_com_options)  # Logger_portの選択変更時にリストを更新
    logger_dropdown = ttk.Combobox(root, textvariable=logger_var, values=ports, state="readonly")
    logger_dropdown.grid(row=0, column=1, padx=10, pady=5)

    # Inverter_portの選択
    tk.Label(root, text="Inverter Port:", anchor="w").grid(row=1, column=0, padx=10, pady=5)
    inverter_var = tk.StringVar(value="なし")
    inverter_var.trace("w", update_com_options)  # Inverter_portの選択変更時にリストを更新
    inverter_dropdown = ttk.Combobox(root, textvariable=inverter_var, values=["なし"] + ports, state="readonly")
    inverter_dropdown.grid(row=1, column=1, padx=10, pady=5)

    # 選択されたポートを表示
    tk.Label(root, text="Selected Logger Port:").grid(row=2, column=0, padx=10, pady=5)
    selected_logger_label = tk.Label(root, textvariable=logger_var, anchor="w")
    selected_logger_label.grid(row=2, column=1, padx=10, pady=5)

    tk.Label(root, text="Selected Inverter Port:").grid(row=3, column=0, padx=10, pady=5)
    selected_inverter_label = tk.Label(root, textvariable=inverter_var, anchor="w")
    selected_inverter_label.grid(row=3, column=1, padx=10, pady=5)

    # 設定ボタン
    submit_button = tk.Button(root, text="設定", command=on_submit)
    submit_button.grid(row=4, column=0, columnspan=2, pady=10)

    root.mainloop()

    # Inverter_portが"なし"の場合はNoneを返す
    return selected_ports["Logger_port"], None if selected_ports["Inverter_port"] == "なし" else selected_ports["Inverter_port"]

# COMポートの設定
LOGGER_PORT, INVERTER_PORT = select_com_ports()
if not LOGGER_PORT:
    print("Logger port selection canceled.")
    exit()

BAUD_RATE = 9600  # ボーレートは必要に応じて変更

# CSVファイルの設定
LOG_FILE = 'serial_log.csv'

# シリアルポートの初期化
ser_logger = serial.Serial(LOGGER_PORT, BAUD_RATE, timeout=1)
try:
    if INVERTER_PORT:
        ser_inverter = serial.Serial(INVERTER_PORT, BAUD_RATE, timeout=1)
        inverter_connected = True
    else:
        print("Inverter is not connected. Using dummy data.")
        inverter_connected = False
except serial.SerialException:
    print("Inverter is not connected. Using dummy data.")
    inverter_connected = False

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
    print("Listening on Logger and Inverter...")
    while True:
        # Loggerからデータを受信
        if ser_logger.in_waiting > 0:
            data_from_logger = ser_logger.read(ser_logger.in_waiting)
            print(f"Received from Logger: {data_from_logger.hex()}")

            # レジスタアドレスを抽出して16進数で表示
            if len(data_from_logger) >= 4:  # レジスタアドレスを抽出するには最低4バイト必要
                register_address = int.from_bytes(data_from_logger[2:4], byteorder='big')
                print(f"Extracted Register Address: 0x{register_address:04X}")

            # データをCSVにログ
            with open(LOG_FILE, mode='a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), 'Logger->Inverter', data_from_logger.hex()])

            # Inverterへ送信またはModbusレスポンスを返す
            if inverter_connected:
                ser_inverter.write(data_from_logger)
            else:
                # Modbusレスポンスを生成してLoggerに返す
                modbus_response = generate_modbus_response(data_from_logger)
                if modbus_response:
                    print(f"Sending Modbus response to Logger: {modbus_response.hex()}")
                    ser_logger.write(modbus_response)

        # Inverterからデータを受信
        if inverter_connected and ser_inverter.in_waiting > 0:
            data_from_inverter = ser_inverter.read(ser_inverter.in_waiting)
            print(f"Received from Inverter: {data_from_inverter.hex()}")

            # データをCSVにログ
            with open(LOG_FILE, mode='a', newline='') as file:
                writer = csv.writer(file)
                writer.writerow([time.strftime('%Y-%m-%d %H:%M:%S'), 'Inverter->Logger', data_from_inverter.hex()])

            # Loggerへ送信
            ser_logger.write(data_from_inverter)

except KeyboardInterrupt:
    print("Terminating...")
finally:
    # シリアルポートを閉じる
    ser_logger.close()
    if inverter_connected:
        ser_inverter.close()
