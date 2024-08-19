Attribute VB_Name = "Module1"
Option Explicit

'// �A�N�e�B�u�u�b�N�̑S���W���[�����G�N�X�|�[�g
'// �i�A�N�e�B�u�u�b�N���Ȃ��ꍇ�͂��� Module ��������Ă���u�b�N�iADIN�Ȃǁj�j

Sub SaveModule()

    Dim parentFolder    As String   '// �e�t�H���_�̃p�X
    Dim bookName        As String   '// �u�b�N��
    Dim bookBaseName    As String   '// �g���q���������u�b�N��
    Dim saveFolder      As String   '// �ۑ��t�H���_
    Dim xlMod           As Object   '// ���W���[���I�u�W�F�N�g

    parentFolder = "VBA"

    '// �A�N�e�B�u�u�b�N���Ȃ��ꍇ
    If ActiveWorkbook Is Nothing Then
        bookName = ThisWorkbook.name
    Else
        bookName = ActiveWorkbook.name
    End If

    With CreateObject("Scripting.FileSystemObject")
        bookBaseName = .GetBaseName(bookName)
        '// �u�b�N���{���t��ۑ��t�H���_���ɂ���B
        saveFolder = MakeFolder(.BuildPath(parentFolder, bookBaseName & "_" & Format(Date, "yyyymmdd")))
        For Each xlMod In Workbooks(bookName).VBProject.VBComponents
            xlMod.Export .BuildPath(saveFolder, xlMod.name & GetModuleExt(xlMod.Type))
        Next xlMod
    End With

    '// �ۑ��t�H���_���J��
    Shell "explorer.exe " & saveFolder, vbNormalFocus

End Sub

'// ���W���[���^�C�v�ɑΉ�����g���q��Ԃ�

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

'// �쐬�����t�H���_�̃p�X��Ԃ�

Private Function MakeFolder(ByVal folder_path As String) As String

    CreateObject("WScript.shell").Run "cmd /c md " & Chr(34) & folder_path & Chr(34), 0, True
    MakeFolder = folder_path

End Function
