VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub aff()
  ThisWorkbook.Unprotect
  Worksheets(1).Visible = xlSheetVeryHidden
  ThisWorkbook.Protect
  ThisWorkbook.Unprotect
End Sub
