import os
import re
import sys
from pywinauto import Application
from pywinauto.keyboard import send_keys
from PIL import ImageGrab
# from pyautogui import screenshot
from time import sleep

# pip install pywinauto
# pip install pyautogui pyscreeze Pillow

EXCEL_APP_PATH = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"


def main() -> None:
    # find_executabl_excel_app_paths()
    excel_file_path = r'C:\Users\unive\dev\vba_labo\a.xlsm'
    # excel_file_path = r'C:\Users\unive\dev\vba_labo\not_found.xlsm'
    # excel_file_path = r'C:\Users\unive\dev\vba_labo\a.png'
    # sign_excel(excel_file_path)
    excel_app = open_excel_and_open_signature_dialog(excel_file_path)
    # print(excel_app[u"デジタル署名"].exists())
    # screenshot_excel_signature(excel_file_path)
    # screenshot_excel_signature(excel_file_path, region=(1371,283,542,785))
    # screenshot_excel_signature(excel_file_path, region=(1371, 13, 544, 782))


def sign_excel(excel_file_path: str):
    """
    # 機能: 引数のExcelファイルを開いて、署名して、上書き保存して閉じる.
    # 引数: 署名したいExcelファイルのパス文字列.
    """

    # Excelファイルを立ち上げ、｢デジタル署名｣ダイアログを開く.
    excel_app = open_excel_and_open_signature_dialog(excel_file_path)
    if excel_app is None:
        return

    # ｢Windowsセキュリティ｣を開く.
    send_keys('C')  # [選択(C)]
    sleep(0.25)

    # ｢Windowsセキュリティ｣で[OK]で証明書を付与.
    send_keys('{TAB} {ENTER}')  # [証明書のプロパティを表示します] -> [OK]
    sleep(0.25)

    # ｢デジタル署名｣を[OK]で閉じる
    send_keys('{TAB} {ENTER}')  # [選択(C)] -> [OK]
    sleep(0.25)

    # ｢Visual Basic｣を閉じる
    send_keys('%q')  # Alt + q
    sleep(0.25)

    # 上書き保存
    send_keys('% F S')  # Alt -> F -> S
    sleep(4.75)

    excel_app.kill()

    return True


def open_excel_and_open_signature_dialog(excel_file_path: str):
    """
    # 機能: 引数のExcelファイルを立ち上げ、｢デジタル署名｣ダイアログまで開く.
    # 引数: ｢デジタル署名｣ダイアログを開きたい、Excelファイルのパス文字列。
    # 返り値: 立ち上げたExcelファイル. 開くことが出来なければ、Noneを返す。
    """

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

    # Excelファイルを開く
    excel_app = Application().start(f'"{EXCEL_APP_PATH}" "{excel_file_path}"')
    sleep(4.75)

    # ｢Visual Basic｣を開く.
    send_keys('% L V')  # 順番に、Alt -> L -> V
    sleep(0.75)

    # ｢デジタル署名｣を開く.
    # Altを押したまま、T -> D. Excel16以降は、Altを1度離した(Alt + T) -> (Alt + D) も可.
    send_keys('%TD') # Alt -> T -> D
    sleep(0.75)

    return excel_app


def screenshot_excel_signature(excel_file_path: str, region=None, output_folder=None):
    """
    # 機能: Exccelファイルに署名して、保存する.
    # 引数: excel_file_path: 署名して保存したいExcelファイルのパス文字列.
    # 引数: region: スクリーンショットする座標･サイズ. (left, top, width, height)で表現する. Noneなら全画面.
    # 引数: output_folder: スクリーンショットを保存するフォルダ.
    """

    # 引数output_folderに指定がなければ、スクリーンショットの保存先を対象ファイルと同じフォルダにする.
    if output_folder is None:
        output_folder = os.path.split(excel_file_path)[0]

    # Excelファイルを立ち上げ、｢デジタル署名｣ダイアログを開く.
    excel_app = open_excel_and_open_signature_dialog(excel_file_path)
    if excel_app is None:
        return

    # [詳細]から、｢証明書｣画面を開く.
    send_keys('D')  # [詳細(D)]
    sleep(0.50)

    excel_filename = os.path.basename(excel_file_path)  # 拡張子あり
    # excel_filename = os.path.splitext(os.path.basename(excel_file_path))[0]  # 拡張子なし

    send_keys('{TAB}')  # [全般タブ]
    screenshot(f"{output_folder}\\{excel_filename}_1_全般.png", region=region)
    sleep(0.25)

    send_keys('{RIGHT}')  # [全般]タブ -> [詳細]タブ
    screenshot(f"{output_folder}\\{excel_filename}_2_詳細.png", region=region)
    sleep(0.25)

    send_keys('{RIGHT}')  # [詳細]タブ -> [証明書のパス]タブ
    screenshot(f"{output_folder}\\{excel_filename}_3_パス.png", region=region)
    sleep(0.25)

    excel_app.kill()


def is_excel_running() -> bool:
    """
    # 機能: Excelが立ち上がっているかどうか真偽値を返す.
    # 返り値: Excelファイルが開いていることがわかれば、True. 開いてなさそうなら、False.
    # その他: from pywinauto import Application
    """
    try:
        excel_app = Application().connect(path="EXCEL.EXE", timeout=1)
        return True
    except:
        return False


def find_executabl_excel_app_paths(message: bool) -> list[str]:
    """
    # 機能: 実行可能なEXCEL.EXEのあるパスを返す.
    # 詳細: 12通り(最大24通り)の可能性の高いパスにEXCEL.EXEがあるか探して、存在するパスのリストを返す.
    # 引数: message: Trueなら見つかったパスを出力する.
    # 返り値: 見つかったEXCEL.EXEのパス文字列のリスト
    # その他: import os
    # その他: C:/Program Files (x86)/Microsoft Office/root/Office16/EXCEL.EXEのようなパスを探す.
    """
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
                possible_paths.append(os.path.join(
                    program_path, "Microsoft Office", root, version, "EXCEL.EXE"))
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


def find_excel_exe() -> None:
    """
    # 機能: EXCEL.EXEを探す
    # 詳細: 変数EXCEL_APP_PATHにあるPATHに、EXCEL.EXEがあれば、それで良い.
    #       存在しなければ存在する可能性の高いパスを探して、EXCEL.EXE見つけ出してEXCEL_APP_PATHを更新.
    #       EXCEL.EXEが見つからなかった場合、エラーを返して、sys.exit(1)でプログラム全体を終了.
    """
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

def screenshot(output_file_path, region=None):
    bbox = None
    if region:
        left,top,width,height = region
        right  = left + width
        bottom = top  + height
        bbox = (left,top,right,bottom)
    screenshot = ImageGrab.grab(bbox=bbox)
    screenshot.save(output_file_path)

if __name__=="__main__":
    main()
# pip install pywinauto
# pip install pyautogui pyscreeze Pillow
