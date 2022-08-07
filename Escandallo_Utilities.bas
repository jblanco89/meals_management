Attribute VB_Name = "Utilities"
Option Explicit

'Subroutine which gets links from parent directory, including files in first child subfolder
Sub ListLinks()
Application.ScreenUpdating = False

Dim wb As Workbook
Dim sh As Worksheet
Dim selection As Range
Dim selection2 As Range
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim subFolder As Object
Dim objSubFolder, objSubFiles, file As Object
Dim i, j As Integer
Dim lastRow As Long
Dim pathFolder As Range
Dim answer As Integer




'We set sheet and workbook
Set wb = ThisWorkbook
Set sh = wb.Sheets("ListaArchivos")
'Set sh2 = wb.Sheets("PlatosPrincipales")

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Set parent path
Set pathFolder = sh.Range("C1")

If pathFolder.Value = "" Then
    answer = MsgBox("Path is empty, please check 'ListaArchivos' sheet and set one path", vbCritical)
    Else
    
    'Get the folder object
    Set objFolder = objFSO.GetFolder(pathFolder)
    
    'Calling error function
    On Error GoTo Manage_Error
    
    
    'we clear contents before searching new files
    If sh.Range("B4") <> " " Or sh.Range("D4") <> " " Or sh.Range("G4") <> " " Then
        Call clearLinks(sh, 4)
    End If
    
    
    'loops through each file in the directory
    j = 3
    i = 3
    
    ' looking for files from parent folder
    For Each file In objFolder.Files
        If file.Name Like "*.xls*" Then
            Set selection = sh.Range(Cells(j + 1, 2), Cells(j + 1, 2))
            sh.Hyperlinks.Add Anchor:=selection, Address:= _
                        file.Path
                        j = j + 1
            End If
    Next file
    
    'looking for files from subfolders
    Set subFolder = objFolder.SubFolders
    
    For Each objSubFolder In subFolder
        Set objSubFiles = objSubFolder
        For Each objFile In objSubFiles.Files
            If objFile.Name Like "*.xls*" Then
                'select cell
                Set selection = sh.Range(Cells(i + 1, 4), Cells(i + 1, 4))
                'create hyperlink in selected cell
                sh.Hyperlinks.Add Anchor:=selection, Address:= _
                    objFile.Path ' , _
                    TextToDisplay:=objFile.Name
                    i = i + 1
            End If
        Next objFile
    Next objSubFolder
End If
        
    
    
    'Error handling
Manage_Error:
If Err.Number = 70 Then ' error 70 means user does not have permissions to read some documents
    Resume Next
End If

Application.ScreenUpdating = True

End Sub

'macro which cleans filled cells in order not to acumulate data
Sub clearLinks(sh As Worksheet, index As Integer)
    Application.ScreenUpdating = False
    
    Dim lrB, lrD As Long
    
    lrB = sh.Range("B" & Rows.Count).End(xlUp).Row
    lrD = sh.Range("D" & Rows.Count).End(xlUp).Row
    
    sh.Range("B" & index, "B50").ClearContents
    sh.Range("D" & index, "D50").ClearContents
    
    Application.ScreenUpdating = True


End Sub

'macro which stacks files from folder and subfolders in same column
Sub heapColumns()
Attribute heapColumns.VB_ProcData.VB_Invoke_Func = " \n14"

Application.ScreenUpdating = False

Dim wb As Workbook
Dim sh As Worksheet
Dim lrB, lrD, lrG, lrH As Long



Set wb = ThisWorkbook
Set sh = wb.Sheets("ListaArchivos")

With sh
    .Range("G4", .Cells(.Rows.Count, .Columns.Count)).Clear
End With

lrB = sh.Range("B" & Rows.Count).End(xlUp).Row
lrD = sh.Range("D" & Rows.Count).End(xlUp).Row



On Error Resume Next
    With sh
        .Range("A4:A" & lrB).Select
        selection.Copy
        .Range("G4:G" & lrB).Select
        selection.PasteSpecial Paste:=xlPasteValues
        lrG = .Range("G" & Rows.Count).End(xlUp).Row

    
        .Range("B4:B" & lrB).Select
        selection.Copy
    
        .Range("H4:H" & lrB).Select
        selection.PasteSpecial Paste:=xlPasteValues
    
    
        lrH = .Range("H" & Rows.Count).End(xlUp).Row
    
        .Range("C4:C" & lrD).Select
        selection.Copy
        .Range("G" & lrG + 1).Select
        selection.PasteSpecial Paste:=xlPasteValues
    
        .Range("D4:D" & lrD).Select
        selection.Copy
        .Range("H" & lrH + 1).Select
        selection.PasteSpecial Paste:=xlPasteValues
    
    Application.CutCopyMode = False
End With
Application.ScreenUpdating = True

End Sub

'this macro extract all info required from selected file
Sub getWorbookInfo()
Application.ScreenUpdating = False
Dim wb As Workbook
Dim sh As Worksheet
Dim source As Range
Dim Source_workbook As Object
Dim i As Long
Dim ListOfSheets As Collection

Set wb = ThisWorkbook
Set sh = wb.Sheets("PlatosPrincipales")

Set source = sh.Range("B3")
Set ListOfSheets = New Collection

On Error Resume Next
Set Source_workbook = Workbooks.Open(source)

With Hoja5.BoxSheetList
    .Clear
    For i = 1 To Source_workbook.Sheets.Count
    'Either we can put all names in an array , here we are printing all the names in Sheet 2
    'sh.Range("N" & i + 5) = Source_workbook.Sheets(i).Name
        ListOfSheets.Add Source_workbook.Sheets(i).Name
        .AddItem ListOfSheets(i)
    Next i
End With

Source_workbook.Close SaveChanges:=False
Application.ScreenUpdating = True

End Sub

'Explorer macro opens file system in order to select a folder directory
Sub Explorer()
Dim strPath As String
Dim wb As Workbook
Dim sh As Worksheet

Set wb = ThisWorkbook
Set sh = wb.Sheets("ListaArchivos")


With Application.FileDialog(msoFileDialogFolderPicker)
    If .Show <> 0 Then
    strPath = .SelectedItems(1)
    End If
    
    sh.Range("C1").Value = strPath
End With

Debug.Print strPath

End Sub
