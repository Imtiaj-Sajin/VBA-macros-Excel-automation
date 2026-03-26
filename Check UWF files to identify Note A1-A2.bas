' macro to check Note A1,A2 (by capital Structure Summeary)



Option Explicit

Sub Check_UWF_Files()

    Dim rootPath As String
    Dim wsResult As Worksheet
    Dim nextRow As Long
    
    rootPath = "E:\SAJIN\RPQC - 2026.03.24 Batch 5\"
    
    ' Create / clear result sheet
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Results")
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "Results"
    Else
        wsResult.Cells.Clear
    End If
    On Error GoTo 0
    
    wsResult.Range("A1").Value = "Files WITH 'Capital Structure Summary'"
    nextRow = 2
    
    ' Start recursive scan
    ScanFolder rootPath, wsResult, nextRow
    
    MsgBox "Done checking files!", vbInformation

End Sub


Sub ScanFolder(folderPath As String, wsResult As Worksheet, ByRef nextRow As Long)

    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    ' Loop files in current folder
    For Each file In folder.Files
        
        If LCase(fso.GetExtensionName(file.Name)) = "xlsm" _
        Or LCase(fso.GetExtensionName(file.Name)) = "xlsx" _
        Or LCase(fso.GetExtensionName(file.Name)) = "xls" Then
            
            CheckFile file.Path, wsResult, nextRow
            
        End If
        
    Next file
    
    ' Loop subfolders (recursive)
    For Each subFolder In folder.SubFolders
        ScanFolder subFolder.Path, wsResult, nextRow
    Next subFolder

End Sub


Sub CheckFile(filePath As String, wsResult As Worksheet, ByRef nextRow As Long)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim foundCell As Range
    
    On Error Resume Next
    
    ' Open workbook (read-only, no update links)
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    
    If wb Is Nothing Then Exit Sub
    
    ' Try to access "Loan Analysis"
    Set ws = Nothing
    Set ws = wb.Sheets("Loan Analysis")
    
    If Not ws Is Nothing Then
        
        ' Search for the text
        Set foundCell = ws.Cells.Find(What:="Capital Structure Summary", _
                                     LookIn:=xlValues, _
                                     LookAt:=xlPart, _
                                     MatchCase:=False)
        
        If Not foundCell Is Nothing Then
            ' Found ? log file path
            wsResult.Cells(nextRow, 1).Value = filePath
            nextRow = nextRow + 1
        End If
        
    End If
    
    ' Close file without saving
    wb.Close SaveChanges:=False
    
    On Error GoTo 0

End Sub

