VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()

Call ListLinks
Call heapColumns

End Sub

'this function is executed every time C1 value is changed
Private Sub Worksheet_Change(ByVal Target As Range)

Application.ScreenUpdating = False
On Error Resume Next
If Target.Address = "$C$1" Then
        Call ListLinks
        Call heapColumns
    End If

    Hoja4.Range("B1").Select

    Application.ScreenUpdating = True

End Sub

'This function calls Explorer macro located in 'Utilities' module
'Explorer macro opens file system in order to select a folder directory
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If selection.Count = 1 Then
        If Not Intersect(Target, Range("C1")) Is Nothing Then
            Call Explorer
        End If
    End If
    
    If ActiveSheet.Name <> "ListaArchivos" Then
        MsgBox "change sheet's name is not allowed, contact developer", vbExclamation
        ActiveSheet.Name = "ListaArchivos"
    End If
    
End Sub
