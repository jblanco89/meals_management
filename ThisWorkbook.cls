VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

Application.ScreenUpdating = False

ThisWorkbook.Sheets("ListaArchivos").Activate
ThisWorkbook.Sheets(1).Activate

Application.ScreenUpdating = True
End Sub

