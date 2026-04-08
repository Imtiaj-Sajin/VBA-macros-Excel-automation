' extract two datapoints dscr nsf and debt yeld data from UWF loan analysis sheets. Just select your root folder thats enough
Option Explicit

Sub Extract_UWF_LoanData()
    Dim rootPath As String
    Dim wsResult As Worksheet
    Dim nextRow As Long
    Dim fso As Object
    Dim folderDialog As Object
    
    ' --- Folder Picker ---
    Set folderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    folderDialog.Title = "Select Root Folder to Process"
    If folderDialog.Show <> -1 Then
        MsgBox "No folder selected. Macro cancelled.", vbExclamation
        Exit Sub
    End If
    rootPath = folderDialog.SelectedItems(1)
    If Right(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
    
    ' --- Setup Results Sheet ---
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("UWF_Results")
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "UWF_Results"
    Else
        wsResult.Cells.Clear
    End If
    On Error GoTo 0
    
    ' --- Headers ---
    With wsResult
        .Cells(1, 1).Value = "File Path"
        .Cells(1, 2).Value = "File Name"
        .Cells(1, 3).Value = "W12 Value (Debt Yield NCF)"
        .Cells(1, 4).Value = "X12 Value (DSCR NCF @ 10k)"
        .Cells(1, 5).Value = "Notes"
        .Range("A1:E1").Font.Bold = True
    End With
    
    nextRow = 2
    
    ' --- Scan Folders ---
    ScanFolders_UWF rootPath, wsResult, nextRow
    
    ' --- Autofit columns ---
    wsResult.Columns("A:E").AutoFit
    
    MsgBox "Done! Results written to 'UWF_Results' sheet.", vbInformation
End Sub

Sub ScanFolders_UWF(folderPath As String, wsResult As Worksheet, ByRef nextRow As Long)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim ext As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set folder = fso.GetFolder(folderPath)
    On Error GoTo 0
    If folder Is Nothing Then Exit Sub
    
    ' Check files in current folder
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "xlsm" Or ext = "xlsx" Or ext = "xls" Then
            ' Only process UWF files (filename starts with "UWF" or "uwf")
            If Left(LCase(fso.GetBaseName(file.Name)), 3) = "uwf" Then
                ProcessUWF_File file.Path, wsResult, nextRow
            End If
        End If
    Next file
    
    ' Recurse into subfolders
    For Each subFolder In folder.SubFolders
        ScanFolders_UWF subFolder.Path, wsResult, nextRow
    Next subFolder
End Sub

Sub ProcessUWF_File(filePath As String, wsResult As Worksheet, ByRef nextRow As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim f3Val As String
    Dim w12Val As Variant
    Dim x12Val As Variant
    Dim notes As String
    Dim w12Header As String
    Dim x12Header As String
    Dim fileName As String
    
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    notes = ""
    w12Val = ""
    x12Val = ""
    
    ' Skip the macro workbook itself
    If filePath = ThisWorkbook.FullName Then Exit Sub
    
    ' Open workbook read-only
    On Error Resume Next
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)
    On Error GoTo 0
    
    If wb Is Nothing Then
        ' Could not open file
        wsResult.Cells(nextRow, 1).Value = filePath
        wsResult.Cells(nextRow, 2).Value = fileName
        wsResult.Cells(nextRow, 3).Value = ""
        wsResult.Cells(nextRow, 4).Value = ""
        wsResult.Cells(nextRow, 5).Value = "ERROR: Could not open file"
        nextRow = nextRow + 1
        Exit Sub
    End If
    
    ' Try to get Loan Analysis sheet
    On Error Resume Next
    Set ws = wb.Sheets("Loan Analysis")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' No "Loan Analysis" sheet found
        wsResult.Cells(nextRow, 1).Value = filePath
        wsResult.Cells(nextRow, 2).Value = fileName
        wsResult.Cells(nextRow, 3).Value = ""
        wsResult.Cells(nextRow, 4).Value = ""
        wsResult.Cells(nextRow, 5).Value = "NOTE: 'Loan Analysis' sheet not found in this file"
        nextRow = nextRow + 1
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Check F3 for "Loan Sizing"
    On Error Resume Next
    f3Val = Trim(CStr(ws.Range("F3").Value))
    On Error GoTo 0
    
    If InStr(1, LCase(f3Val), "loan sizing") = 0 Then
        ' F3 does not say "Loan Sizing"
        wsResult.Cells(nextRow, 1).Value = filePath
        wsResult.Cells(nextRow, 2).Value = fileName
        wsResult.Cells(nextRow, 3).Value = ""
        wsResult.Cells(nextRow, 4).Value = ""
        wsResult.Cells(nextRow, 5).Value = "NOTE: F3 does not contain 'Loan Sizing'. F3 value = '" & f3Val & "'"
        nextRow = nextRow + 1
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Get W12 and X12 values
    On Error Resume Next
    w12Val = ws.Range("W12").Value
    x12Val = ws.Range("X12").Value
    On Error GoTo 0
    
    ' Check W12 header (row 3 of column W) for "Debt Yield"
    On Error Resume Next
    w12Header = Trim(CStr(ws.Range("W3").Value))
    x12Header = Trim(CStr(ws.Range("X3").Value))
    On Error GoTo 0
    
    ' Build notes based on header checks
    If InStr(1, LCase(w12Header), "debt yield") = 0 Then
        notes = notes & "WARN: W column header '" & w12Header & "' does not match expected 'Debt Yield'. "
    End If
    If InStr(1, LCase(x12Header), "dscr") = 0 Then
        notes = notes & "WARN: X column header '" & x12Header & "' does not match expected 'DSCR'. "
    End If
    If notes = "" Then
        notes = "OK"
    End If
    
    ' Write to results
    wsResult.Cells(nextRow, 1).Value = filePath
    wsResult.Cells(nextRow, 2).Value = fileName
    wsResult.Cells(nextRow, 3).Value = w12Val
    wsResult.Cells(nextRow, 4).Value = x12Val
    wsResult.Cells(nextRow, 5).Value = notes
    
    nextRow = nextRow + 1
    
    wb.Close SaveChanges:=False
End Sub
