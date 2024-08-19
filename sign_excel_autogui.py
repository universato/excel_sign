import psutil
import subprocess
from time import sleep
from pyautogui import hotkey, press

EXCEL_APP_PATH = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

def main():
    excel_file_path = r'C:\Users\unive\dev\vba_labo\a.xlsm'
    sign_excel(excel_file_path)

def sign_excel(excel_file_path):
    app = subprocess.Popen([EXCEL_APP_PATH, excel_file_path])
    sleep(4.75)

    # エディタウィンドウ ｢Visual Basic｣を開く.
    press('Alt')
    press('L')
    press('V')
    sleep(0.50)

    # ツール(T) -> ダイアログ･ボックス｢デジタル署名｣を開く.
    hotkey('Alt', 'T', 'D')
    sleep(0.50)

    # [選択(C)]で、｢Windowsセキュリティ｣を開く.
    hotkey('Alt', 'C')
    sleep(0.50)

    # ｢Windowsセキュリティ｣で証明書を付与して閉じる.
    press('Tab')   # [証明書のプロパティを表示します] -> [OK]
    press('Enter') # [OK]
    sleep(0.25)

    # ｢デジタル署名｣を[OK]で閉じる.
    press('Tab')   # [選択(C)] -> [OK]
    press('Enter') # [OK]
    sleep(0.25)

    hotkey('Alt', 'q') # ウィンドウ｢Visual Basic｣を閉じる.
    sleep(0.50)

    # 上書き保存
    hotkey('Ctrl', 'S')

    # Excelプロセスを終了する
    for proc in psutil.process_iter():
        if proc.name().lower() == "excel.exe":
            # proc.terminate()  # 強制終了
            # proc.wait()  # 終了するまで待つ
            print("終了")

main()
