Attribute VB_Name = "Module1"
Option Explicit

Sub ExportAllVBA_UTF8()

    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    Dim FileName As String
    Dim FolderPath As String
    Dim tempFileName As String
    Dim fileContent As String
    Dim fso As Object

    ' エクスポート先のフォルダパスを指定します
    FolderPath = "C:\Users\unive\dev\vba_labo\VBA\vba_20240811" '適宜変更してください

    ' フォルダが存在しない場合、作成します
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If

    Set VBProj = ThisWorkbook.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 各コンポーネントをエクスポートします
    For Each VBComp In VBProj.VBComponents
        
        ' 一時ファイル名を作成します
        tempFileName = FolderPath & VBComp.name & "_temp.bas"

        ' フォームやクラスモジュール、標準モジュールなどで保存するファイルの拡張子を決めます
        Select Case VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                FileName = FolderPath & VBComp.name & ".bas"
            Case vbext_ct_MSForm
                FileName = FolderPath & VBComp.name & ".frm"
            Case vbext_ct_Document
                FileName = FolderPath & VBComp.name & ".cls"
            Case Else
                FileName = FolderPath & VBComp.name & ".txt"
        End Select
        
        ' コードを一時ファイルにエクスポートします
        VBComp.Export tempFileName

        ' 一時ファイルの内容をUTF-8として再保存します
        fileContent = ReadFileAsUTF8(tempFileName)
        WriteFileAsUTF8 FileName, fileContent

        ' 一時ファイルを削除します
        fso.DeleteFile tempFileName
        
    Next VBComp

    MsgBox "VBAコードをUTF-8でエクスポートしました！", vbInformation

End Sub

Function ReadFileAsUTF8(filePath As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim fileContent As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 1, False, -2) ' -2: default encoding
    fileContent = ts.ReadAll
    ts.Close
    
    ReadFileAsUTF8 = fileContent
End Function

Sub WriteFileAsUTF8(filePath As String, content As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    stream.Type = 2 ' Text data
    stream.Mode = 3 ' Read/Write
    stream.Charset = "UTF-8"
    
    stream.Open
    stream.WriteText content
    stream.SaveToFile filePath, 2 ' 2 = overwrite
    stream.Close
End Sub

