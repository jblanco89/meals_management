VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub BoxSheetList_Click()
Application.ScreenUpdating = False

Dim source As Range
Dim Source_workbook As Object
Dim lr, lc As Long


With Hoja5
    .Range("D10", .Cells(.Rows.Count, .Columns.Count)).Clear
End With

Set source = Hoja5.Range("B3")
Set Source_workbook = Workbooks.Open(source)

With Hoja5.BoxSheetList
'        MsgBox .Value
        Source_workbook.Sheets(.Value).UsedRange.Copy
        Hoja5.Range("D10").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        Source_workbook.Close SaveChanges:=False
End With




Application.ScreenUpdating = True

End Sub



Private Sub ComboBox1_Change()

Application.ScreenUpdating = False
    Dim valor As String

    valor = Hoja5.ComboBox1.Value
    Hoja5.Range("C2").Value = valor
    Call getWorbookInfo
    
    On Error Resume Next
    With Hoja5
       .Range("D10", .Cells(.Rows.Count, .Columns.Count)).Clear
    End With
    
    
Application.ScreenUpdating = True
End Sub



'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'If ActiveSheet.Name <> "PlatosPrincipales" Then
''        MsgBox "change sheet's name is not allowed, contact developer", vbExclamation
'        ActiveSheet.Name = "PlatosPrincipales"
'End If
'End Sub
