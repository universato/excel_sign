
# コンパイルエラー
オブジェクトが必要です。

VBAにおいて、文字列は値型。あたいがた??
`Set`は不要。
`Set`があったら、オブジェク型のデータ変数が必要なはずで、
値型の文字列の変数がくると、おかしい。

```vb
Sub hello()
    Dim str As String
    Set str = "Hello, world!!"
    Debug.Print (str)
End Sub
```
↓
```vb
Sub hello()
    Dim str As String
    str = "Hello, world!!"
    Debug.Print (str)
End Sub
```


# 実行時エラー'13'

次のコードは、通常のシート以外に、｢グラフシート｣があったりすると、エラー13となる。
ループで回す変数shはWorksheetという型だが、グラフシートはWorksheetという型ではない。
```vb
Sub f13()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets()
        Debug.Print (sh.Name)
    Next sh
End Sub
```

通常のシートだけで回したい場合は、`Sheets()`ではなく、`Worksheets()`を用いる。
ちなみに、丸括弧なしで`Sheets`でも`Worksheets`でも同じで、丸括弧の有無は関係ない。

```vb
Sub f13()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        Debug.Print (sh.Name)
    Next sh
End Sub
```

グラフシートも回したければ、shの型を変更する。

```vb
Sub f13()
    Dim sh As Variant
    For Each sh In ThisWorkbook.Sheets
        Debug.Print (sh.Name)
    Next sh
End Sub
```

# 実行時エラー'91':
オブジェクト変数またはwithブロック変数が設定されていません。

```vb
Sub f91()
    Dim wb As Workbook
    wb = ThisWorkbook
    'Debug.Print (wb.Name())
End Sub
```
↓ 単に代入するのではなく、`Set`を用いる。
```vb
Sub f91()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    'Debug.Print (wb.Name())
End Sub
```

# 丸括弧があってもなくても良い

よくわかってないけれど、Rubyみたいだな?
```vb
Sub filename()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Debug.Print (wb.Name())
    Debug.Print (wb.Name)
End Sub
```


#

Subプロシージャなのに引数がある!!??
実行されなくなってしまう。
```vb
Sub hello(str)
    Dim str As String
    str = "Hello, world!!"
    Debug.Print (str)
End Sub
```
