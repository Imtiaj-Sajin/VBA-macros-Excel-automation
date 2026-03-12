Sub ExportFast()

    Dim srcWb As Workbook
    Dim newWb As Workbook
    Dim srcWs As Worksheet
    Dim newWs As Worksheet
    Dim ddCell As Range
    Dim allValues() As String
    Dim i As Integer, idx As Integer
    Dim sheetName As String
    Dim valRange As Range

    Set srcWb = ThisWorkbook
    Set srcWs = srcWb.Sheets("Cash Flows")
    Set ddCell = srcWs.Range("PoolPropChoice")

    ' --- Speed settings ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' --- Get dropdown values ---
    Dim valSource As String
    valSource = ddCell.Validation.Formula1
    If Left(valSource, 1) = "=" Then valSource = Mid(valSource, 2)

    On Error Resume Next
    Set valRange = srcWb.Names(valSource).RefersToRange
    If valRange Is Nothing Then Set valRange = Application.Range(valSource)
    On Error GoTo 0

    If valRange Is Nothing Then
        MsgBox "Could not read dropdown list. Exiting.", vbCritical
        GoTo Cleanup
    End If

    ' --- Build values array ---
    idx = 0
    ReDim allValues(1 To valRange.Cells.Count)
    Dim c As Range
    For Each c In valRange.Cells
        If Trim(CStr(c.Value)) <> "" Then
            idx = idx + 1
            allValues(idx) = Trim(CStr(c.Value))
        End If
    Next c
    ReDim Preserve allValues(1 To idx)

    ' --- Create output workbook ---
    Set newWb = Workbooks.Add
    Do While newWb.Sheets.Count > 1
        newWb.Sheets(newWb.Sheets.Count).Delete
    Loop

    Dim savePath As String
    savePath = srcWb.Path & "\CashFlows_Export_" & Format(Now, "YYYYMMDD_HHMMSS") & ".xlsx"

    ' --- Main loop ---
    For i = 1 To idx

        Application.StatusBar = "Exporting " & i & " of " & idx & ": " & allValues(i)

        ' 1. Set dropdown value
        ddCell.Value = allValues(i)

        ' 2. FULL recalculation - calculate entire workbook twice
        '    (second pass catches any formulas dependent on first-pass results)
        Application.Calculate
        Application.Calculate

        ' 3. Verify the dropdown actually changed (safety check)
        If ddCell.Value <> allValues(i) Then
            ddCell.Value = allValues(i)
            Application.Calculate
            Application.Calculate
        End If

        ' 4. Build valid sheet name
        sheetName = Left(allValues(i), 31)
        Dim badChars As Variant
        badChars = Array("/", "\", "*", "?", ":", "[", "]")
        Dim ch As Variant
        For Each ch In badChars
            sheetName = Join(Split(sheetName, CStr(ch)), "_")
        Next ch
        sheetName = Trim(sheetName)

        ' 5. Copy ENTIRE sheet to new workbook (preserves 100% of layout,
        '    merged cells, formatting, row heights, col widths, charts)
        srcWs.Copy After:=newWb.Sheets(newWb.Sheets.Count)
        Set newWs = newWb.Sheets(newWb.Sheets.Count)

        ' 6. On the COPIED sheet: replace every formula with its current value
        '    Do this cell-by-cell only on formula cells (fastest + safest)
        Dim rng As Range
        Dim cell As Range

        On Error Resume Next
        Set rng = newWs.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0

        If Not rng Is Nothing Then
            ' Copy entire used range and paste values in one shot
            newWs.UsedRange.Copy
            newWs.UsedRange.PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
        End If

        ' 7. Rename sheet
        newWs.Name = sheetName

        ' 8. Save after every sheet (prevents data loss if crash mid-run)
        If i = 1 Then
            newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
        Else
            newWb.Save
        End If

    Next i

    ' --- Remove default blank Sheet1 ---
    On Error Resume Next
    newWb.Sheets("Sheet1").Delete
    On Error GoTo 0

    ' --- Final save ---
    newWb.Save

    MsgBox "Done! " & idx & " sheets exported successfully." & vbNewLine & _
           "File saved to: " & vbNewLine & savePath, vbInformation, "Export Complete"

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.CutCopyMode = False

End Sub
