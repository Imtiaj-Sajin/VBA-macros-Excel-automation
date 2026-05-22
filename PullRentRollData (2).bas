Attribute VB_Name = "PullRentRollData"
Option Explicit

'==========================================================================
' Macro: PullRentRollData
' Purpose: Loops through every .xlsx file in a user-selected folder and
'          appends rent-roll data to the "RR as of 5.21.2026" tab of the
'          ACTIVE workbook, starting at the first empty row (default 2081).
'
' How to use:
'   1. Open your destination workbook (the one with the "RR as of 5.21.2026" tab).
'   2. Press Alt+F11 -> Insert -> Module -> paste this code.
'   3. Run "PullRentRollData".
'   4. Pick the folder containing the source rent-roll files.
'==========================================================================

Public Sub PullRentRollData()

    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim folderPath As String
    Dim fileName As String
    Dim destRow As Long
    Dim startRow As Long
    Dim filesProcessed As Long
    Dim rowsAdded As Long
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean

    Set wbDest = ActiveWorkbook

    ' ---- Find the destination sheet ----
    On Error Resume Next
    Set wsDest = wbDest.Sheets("RR as of 5.21.2026")
    On Error GoTo 0
    If wsDest Is Nothing Then
        MsgBox "Could not find a tab named ""RR as of 5.21.2026"" in the active workbook.", vbCritical
        Exit Sub
    End If

    ' ---- Ask user for source folder ----
    folderPath = PickFolder()
    If folderPath = "" Then Exit Sub
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' ---- Determine the first empty destination row ----
    ' Default expectation: start at row 2081. If user already pulled more,
    ' we'll continue at the next empty row in column D.
    startRow = 2081
    destRow = startRow
    Do While Len(CStr(wsDest.Cells(destRow, "D").Value)) > 0
        destRow = destRow + 1
    Loop

    ' ---- Speed up Excel ----
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    filesProcessed = 0
    rowsAdded = 0

    ' ---- Loop through every .xlsx file in the folder ----
    fileName = Dir(folderPath & "*.xlsx")
    Do While Len(fileName) > 0
        ' Skip Excel temp/lock files (~$...)
        If Left$(fileName, 2) <> "~$" Then
            Set wbSrc = Nothing
            On Error Resume Next
            Set wbSrc = Workbooks.Open(fileName:=folderPath & fileName, ReadOnly:=True, UpdateLinks:=0)
            On Error GoTo 0

            If Not wbSrc Is Nothing Then
                Set wsSrc = GetSourceSheet(wbSrc)
                If Not wsSrc Is Nothing Then
                    Dim addedThisFile As Long
                    addedThisFile = AppendFromSource(wsSrc, wsDest, destRow)
                    destRow = destRow + addedThisFile
                    rowsAdded = rowsAdded + addedThisFile
                    filesProcessed = filesProcessed + 1
                End If
                wbSrc.Close SaveChanges:=False
            End If
        End If
        fileName = Dir
    Loop

    ' ---- Restore Excel ----
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreen
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = True
    Application.CalculateFull

    MsgBox "Done." & vbCrLf & _
           "Files processed: " & filesProcessed & vbCrLf & _
           "Rows added: " & rowsAdded & vbCrLf & _
           "First row written: " & startRow & vbCrLf & _
           "Last row written: " & (destRow - 1), vbInformation

End Sub

'--------------------------------------------------------------------------
' Identify the data sheet inside a source rent-roll workbook.
' Rule: the data sheet has "Property Look-Up Code" in cell A7.
' (Each file has 2-3 tabs; the first one is the data tab, but we detect by
'  content instead of relying on order, in case the file is structured a bit
'  differently.)
'--------------------------------------------------------------------------
Private Function GetSourceSheet(wb As Workbook) As Worksheet

    Dim sh As Worksheet
    Dim a7 As String

    ' Prefer a sheet that has "Property Look-Up Code" in A7
    For Each sh In wb.Worksheets
        a7 = ""
        On Error Resume Next
        a7 = CStr(sh.Range("A7").Value)
        On Error GoTo 0
        If LCase$(Trim$(a7)) = "property look-up code" Then
            Set GetSourceSheet = sh
            Exit Function
        End If
    Next sh

    ' Fallback: first sheet whose name isn't "Report Parameters"
    For Each sh In wb.Worksheets
        If LCase$(sh.Name) <> "report parameters" Then
            Set GetSourceSheet = sh
            Exit Function
        End If
    Next sh

End Function

'--------------------------------------------------------------------------
' Copies the data rows from a source sheet into the destination, mapping
' each destination column by header name.
' Returns the number of rows added.
'--------------------------------------------------------------------------
Private Function AppendFromSource(wsSrc As Worksheet, wsDest As Worksheet, destStartRow As Long) As Long

    Dim srcPropName As Variant   ' from source A3
    Dim srcAsOfDate As Variant   ' from source A4
    Dim headerRow As Long
    Dim lastSrcRow As Long
    Dim lastSrcCol As Long
    Dim r As Long, c As Long
    Dim headerName As String
    Dim destCol As Long
    Dim rowsWritten As Long

    srcPropName = wsSrc.Range("A3").Value
    srcAsOfDate = wsSrc.Range("A4").Value

    headerRow = 7
    lastSrcCol = wsSrc.Cells(headerRow, wsSrc.Columns.Count).End(xlToLeft).Column
    ' Last data row: use column A (Property Look-Up Code) under the header.
    lastSrcRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    ' Build a dictionary of source-column header -> source column index
    Dim srcHeaders As Object
    Set srcHeaders = CreateObject("Scripting.Dictionary")
    srcHeaders.CompareMode = 1 ' text compare, case-insensitive
    For c = 1 To lastSrcCol
        headerName = Trim$(CStr(wsSrc.Cells(headerRow, c).Value))
        If Len(headerName) > 0 Then
            If Not srcHeaders.Exists(headerName) Then
                srcHeaders.Add headerName, c
            End If
        End If
    Next c

    rowsWritten = 0

    For r = headerRow + 1 To lastSrcRow
        Dim plc As String
        plc = Trim$(CStr(wsSrc.Cells(r, 1).Value))

        ' Skip blank rows and the "Total" footer row
        If Len(plc) = 0 Then GoTo NextRow
        If LCase$(plc) = "total" Then GoTo NextRow
        ' Skip any row that doesn't have a Bldg-Unit / Unit / Floor Plan combo
        If Len(Trim$(CStr(wsSrc.Cells(r, 2).Value))) = 0 And _
           Len(Trim$(CStr(wsSrc.Cells(r, 3).Value))) = 0 Then GoTo NextRow

        Dim destR As Long
        destR = destStartRow + rowsWritten

        ' Source-derived fixed fields
        wsDest.Cells(destR, "D").Value = srcPropName    ' Source Property Name
        wsDest.Cells(destR, "E").Value = srcAsOfDate    ' RR As Of Date

        ' Map every yellow column by header name
        WriteMapped wsSrc, wsDest, r, destR, "Property Look-Up Code", "G", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Bldg-Unit", "H", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Unit", "I", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Floor Plan", "J", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Unit Type", "K", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "SQFT", "L", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Unit Status", "N", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Resident", "O", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Lease Start", "Q", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Move-In", "R", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Lease End", "S", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Expected Move-Out", "T", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Lease Term", "U", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Deposit Held", "V", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Market Rent", "W", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Balance", "X", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Actual Charges", "AD", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Total Rent (Rent + Amenity)", "AE", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "RENT", "AF", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Amenity Rent", "AG", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "ACLR", "AH", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Month-to-Month Rent [SYSTEM]", "AI", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "CONC/SPECL", "AJ", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "CABLE", "AK", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "CAM", "AL", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "COMELEC", "AM", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "COMGAS", "AN", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "GAS", "AO", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "SEWER", "AP", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "TRASH", "AQ", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "VALETTRASH", "AR", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "WATR", "AS", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "STRMWTR", "AT", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "UTILITY SETUP FEE", "AU", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "DEPOSIT ALTERNATIVE", "AV", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "LIFESTYLE FEE", "AW", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "EMPLCRED", "AX", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "HOA", "AY", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "INSURE", "AZ", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "RLL", "BA", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "INTERNET", "BB", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "PACKLOCKRENT", "BC", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "GARAGE (Add on)", "BD", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "GARAGE (Base)", "BE", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "PARKING (Add on)", "BF", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "MODELCRED", "BG", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "PARKING (Base)", "BH", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "PEST", "BI", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "PETFEE", "BJ", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "PETRENT", "BK", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "PETRENT2", "BL", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "STORAGE (Add on)", "BM", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "STORAGE (Base)", "BN", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "TAX", "BO", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "TECHFEE", "BP", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "CLEANING", "BQ", srcHeaders
        WriteMapped wsSrc, wsDest, r, destR, "Total", "BR", srcHeaders

        rowsWritten = rowsWritten + 1
NextRow:
    Next r

    AppendFromSource = rowsWritten

End Function

'--------------------------------------------------------------------------
' Helper: look up header in source, copy that cell to destination column.
' If the header doesn't exist in the source, the destination cell is left
' blank (as the user requested).
'--------------------------------------------------------------------------
Private Sub WriteMapped(wsSrc As Worksheet, wsDest As Worksheet, srcRow As Long, destRow As Long, _
                        srcHeader As String, destColLetter As String, srcHeaders As Object)

    If Not srcHeaders.Exists(srcHeader) Then Exit Sub
    Dim srcCol As Long
    srcCol = srcHeaders(srcHeader)
    wsDest.Range(destColLetter & destRow).Value = wsSrc.Cells(srcRow, srcCol).Value

End Sub

'--------------------------------------------------------------------------
' Folder picker dialog
'--------------------------------------------------------------------------
Private Function PickFolder() As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select the folder containing source rent-roll files"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = ""
    End If

End Function
