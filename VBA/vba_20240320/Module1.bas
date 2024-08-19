Attribute VB_Name = "Module1"
Option Explicit

'// アクティブブックの全モジュールをエクスポート
'// （アクティブブックがない場合はこの Module が書かれているブック（ADINなど））

Sub SaveModule()

    Dim parentFolder    As String   '// 親フォルダのパス
    Dim bookName        As String   '// ブック名
    Dim bookBaseName    As String   '// 拡張子を除いたブック名
    Dim saveFolder      As String   '// 保存フォルダ
    Dim xlMod           As Object   '// モジュールオブジェクト

    parentFolder = "VBA"

    '// アクティブブックがない場合
    If ActiveWorkbook Is Nothing Then
        bookName = ThisWorkbook.name
    Else
        bookName = ActiveWorkbook.name
    End If

    With CreateObject("Scripting.FileSystemObject")
        bookBaseName = .GetBaseName(bookName)
        '// ブック名＋日付を保存フォルダ名にする。
        saveFolder = MakeFolder(.BuildPath(parentFolder, bookBaseName & "_" & Format(Date, "yyyymmdd")))
        For Each xlMod In Workbooks(bookName).VBProject.VBComponents
            xlMod.Export .BuildPath(saveFolder, xlMod.name & GetModuleExt(xlMod.Type))
        Next xlMod
    End With

    '// 保存フォルダを開く
    Shell "explorer.exe " & saveFolder, vbNormalFocus

End Sub

'// モジュールタイプに対応する拡張子を返す

Private Function GetModuleExt(ByVal module_type As Integer) As String

    Select Case module_type
        Case 1
            GetModuleExt = ".bas"
        Case 2, 100
            GetModuleExt = ".cls"
        Case 3
            GetModuleExt = ".frm"
    End Select

End Function

'// 作成したフォルダのパスを返す

Private Function MakeFolder(ByVal folder_path As String) As String

    CreateObject("WScript.shell").Run "cmd /c md " & Chr(34) & folder_path & Chr(34), 0, True
    MakeFolder = folder_path

End Function
