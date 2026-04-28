Sub ReplaceColorAcrossWorkbook()

    Dim colorCell1 As Range
    Dim colorCell2 As Range
    Dim color1 As Long
    Dim color2 As Long
    Dim ws As Worksheet
    Dim cell As Range
    Dim replaced As Long
    Dim msg As String

    '------------------------------------------------------------
    ' STEP 1: Pick the SOURCE color cell
    '------------------------------------------------------------
    On Error Resume Next
    Set colorCell1 = Application.InputBox( _
        Prompt:="Step 1 of 2:" & vbCrLf & vbCrLf & _
                "Click on ANY cell whose background color you want to REPLACE.", _
        Title:="Select Source Color Cell", _
        Type:=8)
    On Error GoTo 0

    If colorCell1 Is Nothing Then
        MsgBox "Cancelled. No changes were made.", vbInformation, "Cancelled"
        Exit Sub
    End If

    color1 = colorCell1.Interior.Color

    ' Warn if the cell has no fill (xlNone)
    If colorCell1.Interior.ColorIndex = xlNone Then
        MsgBox "The cell you selected has NO background color (transparent)." & vbCrLf & _
               "Please select a cell that has a visible background color.", _
               vbExclamation, "No Color Detected"
        Exit Sub
    End If

    '------------------------------------------------------------
    ' STEP 2: Pick the REPLACEMENT color cell
    '------------------------------------------------------------
    On Error Resume Next
    Set colorCell2 = Application.InputBox( _
        Prompt:="Step 2 of 2:" & vbCrLf & vbCrLf & _
                "Click on ANY cell whose background color you want to USE AS REPLACEMENT.", _
        Title:="Select Replacement Color Cell", _
        Type:=8)
    On Error GoTo 0

    If colorCell2 Is Nothing Then
        MsgBox "Cancelled. No changes were made.", vbInformation, "Cancelled"
        Exit Sub
    End If

    If colorCell2.Interior.ColorIndex = xlNone Then
        MsgBox "The replacement cell has NO background color (transparent)." & vbCrLf & _
               "Please select a cell that has a visible background color.", _
               vbExclamation, "No Color Detected"
        Exit Sub
    End If

    color2 = colorCell2.Interior.Color

    If color1 = color2 Then
        MsgBox "Both cells have the same background color. Nothing to replace.", _
               vbInformation, "Same Color"
        Exit Sub
    End If

    '------------------------------------------------------------
    ' STEP 3: Confirm before applying
    '------------------------------------------------------------
    msg = "Are you sure you want to replace ALL cells across ALL sheets?" & vbCrLf & vbCrLf & _
          "  Source color (to be replaced):   RGB(" & _
              ((color1 Mod 256)) & ", " & _
              (Int(color1 / 256) Mod 256) & ", " & _
              (Int(color1 / 65536) Mod 256) & ")" & vbCrLf & _
          "  New color (replacement):          RGB(" & _
              ((color2 Mod 256)) & ", " & _
              (Int(color2 / 256) Mod 256) & ", " & _
              (Int(color2 / 65536) Mod 256) & ")" & vbCrLf & vbCrLf & _
          "This will affect ALL worksheets in this workbook."

    If MsgBox(msg, vbQuestion + vbYesNo, "Confirm Color Replacement") = vbNo Then
        MsgBox "Cancelled. No changes were made.", vbInformation, "Cancelled"
        Exit Sub
    End If

    '------------------------------------------------------------
    ' STEP 4: Replace color across all sheets
    '------------------------------------------------------------
    Application.ScreenUpdating = False
    replaced = 0

    For Each ws In ThisWorkbook.Worksheets
        For Each cell In ws.UsedRange
            If cell.Interior.ColorIndex <> xlNone Then
                If cell.Interior.Color = color1 Then
                    cell.Interior.Color = color2
                    replaced = replaced + 1
                End If
            End If
        Next cell
    Next ws

    Application.ScreenUpdating = True

    '------------------------------------------------------------
    ' STEP 5: Done — show summary
    '------------------------------------------------------------
    MsgBox "Done! " & replaced & " cell(s) were updated across " & _
           ThisWorkbook.Worksheets.Count & " sheet(s).", _
           vbInformation, "Replacement Complete"

End Sub
