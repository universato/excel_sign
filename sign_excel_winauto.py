import os
import re
import sys
from pywinauto import Application
from pywinauto.keyboard import send_keys
from pyautogui import screenshot
from time import sleep

# pip install pywinauto
# pip install pyautogui pyscreeze Pillow

EXCEL_APP_PATH = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"

def main():
    # find_executabl_excel_app_paths()
    excel_file_path = r'C:\Users\unive\dev\vba_labo\a.xlsm'
    # excel_file_path = r'C:\Users\unive\dev\vba_labo\not_found.xlsm'
    # excel_file_path = r'C:\Users\unive\dev\vba_labo\a.png'
    # sign_excel(excel_file_path)
    screenshot_excel_signature(excel_file_path, region=(1371,13,544,782))



# 機能: 引数のExcelファイルを開いて、署名して、上書き保存して閉じる.
def sign_excel(excel_file_path):
    # Excelファイルを立ち上げ、｢デジタル署名｣ダイアログを開く.
    excel_app = open_excel_and_open_signature_dialog(excel_file_path)
    if excel_app is None:
        return

    # ｢Windowsセキュリティ｣を開く.
    send_keys('%C') # [選択(C)] Alt + C
    sleep(0.25)

    # ｢Windowsセキュリティ｣で[OK]で証明書を付与.
    send_keys('{TAB} {ENTER}') # [証明書のプロパティを表示します] -> [OK]
    sleep(0.25)

    # ｢デジタル署名｣を[OK]で閉じる
    send_keys('{TAB} {ENTER}') # [選択(C)] -> [OK]
    sleep(0.25)

    # ｢Visual Basic｣を閉じる
    send_keys('%q')  # Alt + q
    sleep(0.25)

    # 上書き保存
    send_keys('% F S') # Alt -> F -> S
    sleep(4.75)

    excel_app.kill()

# 機能: 引数のExcelファイルを立ち上げ、｢デジタル署名｣ダイアログまで開く.
# 返り値: 立ち上げたExcelファイル
def open_excel_and_open_signature_dialog(excel_file_path):
    # 引数のExcelファイルが存在しなければ、Noneを返し、呼び出し元に処理を中止してもらう.
    if not os.path.exists(excel_file_path):
        print(f"[ALERT] Excel file Not Found: {excel_file_path}")
        return

    # 引数がExcelファイルでなければ、Noneを返し、呼び出し元に処理を中止してもらう.
    extension_name = excel_file_path.split('.')[-1].lower()
    if extension_name not in ['xls', 'xlsx', 'xlsm']:
        print(f"[ALERT] Not Excel file: {excel_file_path}")
        return

    if is_excel_running():
        print(f"[Alert] Aleady Excel Opend: Excelが既に動いてしまっている状態で次のファイルの処理が開始されました.")
        print(f"\t{excel_file_path}")
        # sys.exit(1)

    # 変数EXCEL_APP_PATHに代入されたパスにEXCEL.EXEが存在するか確かめる。
    # 存在しなければ、探してEXCEL_APP_PATHを更新する。
    # 探しても、見つからなければエラー終了。
    find_excel_exe()

    # Excelを開く
    excel_app = Application().start(f'"{EXCEL_APP_PATH}" "{excel_file_path}"')
    sleep(4.75)

    # ｢Visual Basic｣を開く.
    send_keys('% L V') # 順番に、Alt -> L -> V
    sleep(0.75)

    # ｢デジタル署名｣を開く.
    send_keys('%TD') # Altを押したまま、T -> D. Excel16以降は、Altを1度離した(Alt + T) -> (Alt + D) も可.
    sleep(0.75)

    return excel_app


# 機能: Exccelに署名して、保存する.
# 前提: 署名対象のExcelフィルの｢デジタル署名｣ダイアログが開いてフォーカスがあたっている状態で使う.
def screenshot_excel_signature(excel_file_path, region=None, output_folder=None):
    # 引数output_folderに指定がなければ、スクリーンショットの保存先を対象ファイルと同じフォルダにする.
    if output_folder is None:
        output_folder = os.path.split(excel_file_path)[0]

    # Excelファイルを立ち上げ、｢デジタル署名｣ダイアログを開く.
    excel_app = open_excel_and_open_signature_dialog(excel_file_path)
    if excel_app is None:
        return

    # [詳細]から、｢証明書｣画面を開く.
    send_keys('%D')  # [詳細(D)] Alt + D
    sleep(0.50)

    excel_filename = os.path.basename(excel_file_path) # 拡張子あり
    # excel_filename = os.path.splitext(os.path.basename(excel_file_path))[0] # 拡張子なし

    send_keys('{TAB}') # [全般タブ]
    screenshot(f"{output_folder}\\{excel_filename}_1_全般.png", region=region)
    sleep(0.25)

    send_keys('{RIGHT}') # [全般]タブ -> [詳細]タブ
    screenshot(f"{output_folder}\\{excel_filename}_2_詳細.png", region=region)
    sleep(0.25)

    send_keys('{RIGHT}') # [詳細]タブ -> [証明書のパス]タブ
    screenshot(f"{output_folder}\\{excel_filename}_3_パス.png", region=region)
    sleep(0.25)

    # ｢証明書｣画面を閉じ、｢デジタル署名｣ダイアログを閉じる.
    send_keys('{ESC} {ESC}')
    sleep(0.25)

    excel_app.kill()

# 機能: EXCEL.EXEが立ち上がっているかどうか真偽値を返す.
def is_excel_running() -> bool:
    try:
        excel_app = Application().connect(path="EXCEL.EXE", timeout=1)
        return True
    except:
        return False

def find_executabl_excel_app_paths(message=True):
    program_paths = [
        r"C:\Program Files",
        r"C:\Program Files (x86)",
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)")
    ]
    program_paths = set(filter(None, program_paths))

    possible_paths = []
    for program_path in program_paths:
        for root in ["", "root"]:
            for version in ["Office14", "Office15", "Office16"]:
                possible_paths.append(os.path.join(program_path, "Microsoft Office", root, version, "EXCEL.EXE"))
    # print("\n".join(possible_paths))

    exist_paths = []
    for path in possible_paths:
        if os.path.exists(path):
            exist_paths.append(path)

    if message:
        if len(exist_paths) == 0:
            print("[Alert] EXCEL.EXE Not found: 可能性のある場所を探しましたが、Excelアプリが見つかりませんでした.")
        else:
            print("[INFO] EXCEL.EXE found at: ")
            print("\t" + "\n\t".join(exist_paths))

    return exist_paths

# 機能: EXCEL.EXEを探す
# 詳細: EXCEL.EXEが探しても見つからなかった場合、エラーを返して、sys.exit(1)でプログラム全体を終了する。
#       そのため、この関数が2回目の場合は、既にEXCEL.EXEを見つけているはずなので、何度も
#       EXCEL.EXEが見つかった場合、excel_app_existsをTrueにし、何度も探さないようにする。
#       変数EXCEL_APP_PATHにあるPATHに、EXCEL.EXEがあれば、それを使う。
#       ここになければ、可能性のあるPATHを列挙して、
def find_excel_exe() -> None:
    global EXCEL_APP_PATH

    if os.path.exists(EXCEL_APP_PATH):
        return

    excel_app_paths = find_executabl_excel_app_paths(message=False)
    if excel_app_paths:
        if 1 < len(excel_app_paths):
            print(f"[INFO] 複数のEXCEL.EXEが見つかりました。最後のを用います")
            print("\t" + "\n\t".join(excel_app_paths))
        EXCEL_APP_PATH = excel_app_paths[-1]
    else:
        print(f"[ERROR] EXCEL.EXE Not Found: {EXCEL_APP_PATH}")
        print(f"\tPCにExcelは入っていますか? ")
        print(f"\t入っているなら、EXCEL_APP_PATHにEXCEL.EXEのパスを代入して下さい。")
        sys.exit(1)

main()
# pip install pywinauto
# pip install pyautogui pyscreeze Pillow
