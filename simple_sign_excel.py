def simple_sign_excel(excel_file_path):
    # Excelファイルを開く
    excel_app = Application().start(f'"{EXCEL_APP_PATH}" "{excel_file_path}"')
    sleep(4.75)

    # ｢Visual Basic｣を開く.
    send_keys('% L V')  # 順番に、Alt -> L -> V
    sleep(0.75)

    # ｢デジタル署名｣を開く.
    send_keys('% T D')  # Alt -> T -> D
    sleep(0.75)

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
