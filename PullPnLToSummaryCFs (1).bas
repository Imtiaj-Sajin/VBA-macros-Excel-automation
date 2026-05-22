Attribute VB_Name = "PullPnLToSummaryCFs"
Option Explicit

'==========================================================================
' Macro: PullPnLToSummaryCFs
'
' Purpose: For each P&L workbook in a user-selected folder, pull the
'   2025 / TTM Apr 2026 / April 2026 Reforecasts values from the
'   "Adjustments" tab and drop them into the matching property block of
'   the "New Summary CFs" tab in the ACTIVE workbook.
'
' How to use:
'   1. Open Summary Cash Flows workbook (the one with "New Summary CFs").
'   2. Alt+F11 -> Insert -> Module -> paste this code.
'   3. Run "PullPnLToSummaryCFs".
'   4. Pick the folder with the P&L source files.
'
' Source-file expectations (verified against 3.5.4.x ... P&L files):
'   - Sheet "Adjustments" exists and has the standard layout:
'       Row 4-8   : stats (Rooms Avail / Occupied / Occ% / ADR / RevPAR)
'       Row 11-15 : revenue (Rooms, F&B, Other Op., Misc, Total Rev)
'       Row 18-22 : profits  (Rooms, F&B, Other Op., Misc, Total Profit)
'       Row 25-30 : undistributed (A&G, Info, S&M, R&M, Util, Total Undist.)
'       Row 32    : Gross Operating Profit
'       Row 34    : Management Fees
'       Row 36    : Income Before Fixed Charges
'       Row 39-43 : non-op (Other Non Op Rev, Rent, Taxes, Insurance, Total)
'       Row 47    : EBITDA
'       Row 49    : Reserve
'       Row 51    : Net Operating Income
'   - Source P&L block columns: B=2021, C=2022, D=2023, E=Dec 2024,
'     F=Dec 2025, G=Apr 2026, H=Dec 2026 (fcst). We pull F, G, H.
'   - Adjustment block columns: T=2021, U=2022, V=2023, W=Dec 2024,
'     X=Dec 2025, Y=Apr 2026, Z=Dec 2026. We pull X, Y, Z.
'   - Property name is the workbook filename with the leading
'     "<numbers> <numbers> - " stripped and the trailing " P&L.xlsx"
'     stripped.  E.g. "3.5.4.10 260520 - Houston Marriott P&L.xlsx"
'     -> "Houston Marriott".
'
' Yellow highlighting:
'   Every cell the macro writes is filled with yellow so you can see
'   what changed.  To turn this OFF, search this file for the keyword
'   "YELLOW_HIGHLIGHT" and set the constant HIGHLIGHT_YELLOW to False.
'
' Adjustment lines (rows 65-74):
'   The macro will also fill in the adjustment lines from the source
'   Adjustments tab (consistent with how past years 2021-2024 are filled).
'   To disable this and leave rows 65-74 blank for the new columns,
'   search this file for "FILL_ADJUSTMENTS" and set the constant to False.
'==========================================================================

' === YELLOW_HIGHLIGHT toggle ===
' Set to False to skip yellow shading of modified cells.
Private Const HIGHLIGHT_YELLOW As Boolean = True
Private Const YELLOW_COLOR As Long = &HFFFF&   ' standard Excel yellow (RGB 255,255,0)

' === FILL_ADJUSTMENTS toggle ===
' Set to True to also populate adjustment lines (rows 65-74) for the new
' columns.  Past-year columns (2021-2024) already have these filled, so
' setting this True keeps the new columns consistent.  Set to False if you
' want to match the existing Houston Marriott pattern exactly (which left
' rows 65-74 blank for 2025/TTM/Reforecasts).
Private Const FILL_ADJUSTMENTS As Boolean = True

Public Sub PullPnLToSummaryCFs()

    Dim wbDest As Workbook
    Dim wsDest As Worksheet
    Dim wbSrc As Workbook
    Dim folderPath As String
    Dim fileName As String
    Dim filesProcessed As Long
    Dim filesSkipped As Long
    Dim skippedNames As String
    Dim prevCalc As XlCalculation
    Dim prevScreen As Boolean
    Dim prevEvents As Boolean

    Set wbDest = ActiveWorkbook

    ' ---- Find the destination sheet ----
    On Error Resume Next
    Set wsDest = wbDest.Sheets("New Summary CFs")
    On Error GoTo 0
    If wsDest Is Nothing Then
        MsgBox "Could not find a tab named ""New Summary CFs"" in the active workbook.", vbCritical
        Exit Sub
    End If

    ' ---- Build the property-name -> column map from row 3 of destination ----
    Dim propMap As Object
    Set propMap = BuildPropertyColumnMap(wsDest)
    If propMap.Count = 0 Then
        MsgBox "Could not parse property header row (row 3) on 'New Summary CFs'.", vbCritical
        Exit Sub
    End If

    ' ---- Ask user for source folder ----
    folderPath = PickFolder()
    If folderPath = "" Then Exit Sub
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' ---- Speed up Excel ----
    prevCalc = Application.Calculation
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    filesProcessed = 0
    filesSkipped = 0
    skippedNames = ""

    fileName = Dir(folderPath & "*.xlsx")
    Do While Len(fileName) > 0
        ' Skip Excel temp/lock files (~$...)
        If Left$(fileName, 2) <> "~$" Then
            Dim propName As String
            propName = PropertyNameFromFileName(fileName)
            Dim startCol As Long
            startCol = 0
            If propMap.Exists(propName) Then startCol = propMap(propName)

            If startCol = 0 Then
                filesSkipped = filesSkipped + 1
                skippedNames = skippedNames & " - " & fileName & " (no match for """ & propName & """)" & vbCrLf
            Else
                Set wbSrc = Nothing
                On Error Resume Next
                Set wbSrc = Workbooks.Open(fileName:=folderPath & fileName, ReadOnly:=True, UpdateLinks:=0)
                On Error GoTo 0

                If Not wbSrc Is Nothing Then
                    Dim wsAdj As Worksheet, wsOM As Worksheet
                    On Error Resume Next
                    Set wsAdj = wbSrc.Sheets("Adjustments")
                    Set wsOM = wbSrc.Sheets("OM P&L")
                    On Error GoTo 0

                    If wsAdj Is Nothing Then
                        filesSkipped = filesSkipped + 1
                        skippedNames = skippedNames & " - " & fileName & " (no 'Adjustments' tab)" & vbCrLf
                    Else
                        FillBlockFromSource wsDest, startCol, wsAdj, wsOM
                        filesProcessed = filesProcessed + 1
                    End If
                    wbSrc.Close SaveChanges:=False
                Else
                    filesSkipped = filesSkipped + 1
                    skippedNames = skippedNames & " - " & fileName & " (could not open)" & vbCrLf
                End If
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

    Dim msg As String
    msg = "Done." & vbCrLf & _
          "Files processed: " & filesProcessed & vbCrLf & _
          "Files skipped:   " & filesSkipped
    If Len(skippedNames) > 0 Then msg = msg & vbCrLf & vbCrLf & "Skipped:" & vbCrLf & skippedNames
    MsgBox msg, vbInformation

End Sub

'--------------------------------------------------------------------------
' For one property block, write the three columns: 2025, TTM Apr 2026,
' April 2026 Reforecasts.  startCol points at the property's "2021" column;
' "2025" is at startCol+4, "TTM Apr 2026" is at startCol+5,
' "April 2026 Reforecasts" is at startCol+6.
'
' Source Adjustments tab column map for P&L block:
'   Dec 2025 = F (6), Apr 2026 = G (7), Dec 2026 = H (8)
' Source Adjustments tab column map for ADJUSTMENT block (rows 54+):
'   Dec 2025 = X (24), Apr 2026 = Y (25), Dec 2026 = Z (26)
'--------------------------------------------------------------------------
Private Sub FillBlockFromSource(wsDest As Worksheet, startCol As Long, _
                                wsAdj As Worksheet, wsOM As Worksheet)

    Dim offsets As Variant
    Dim plCols As Variant
    Dim adjCols As Variant
    offsets = Array(4, 5, 6)         ' dest col offsets: 2025, TTM, Reforecasts
    plCols = Array(6, 7, 8)          ' source Adj P&L cols
    adjCols = Array(24, 25, 26)      ' source Adj adjustment-block cols

    Dim k As Long
    For k = 0 To 2
        Dim dCol As Long, pCol As Long, aCol As Long
        dCol = startCol + offsets(k)
        pCol = plCols(k)
        aCol = adjCols(k)

        ' Stats (rows 4-8 in Adj P&L block)
        WriteCell wsDest, 8, dCol, wsAdj.Cells(4, pCol).Value   ' Rooms Available
        WriteCell wsDest, 9, dCol, wsAdj.Cells(5, pCol).Value   ' Occupied Rooms
        WriteCell wsDest, 10, dCol, wsAdj.Cells(6, pCol).Value  ' Occupancy %
        WriteCell wsDest, 11, dCol, wsAdj.Cells(7, pCol).Value  ' ADR
        WriteCell wsDest, 12, dCol, wsAdj.Cells(8, pCol).Value  ' RevPAR

        ' Revenue (rows 11-14 in Adj P&L block)
        WriteCell wsDest, 15, dCol, wsAdj.Cells(11, pCol).Value ' Rooms Rev
        WriteCell wsDest, 16, dCol, wsAdj.Cells(12, pCol).Value ' F&B Rev
        WriteCell wsDest, 17, dCol, wsAdj.Cells(13, pCol).Value ' Other Op. Departments Rev
        WriteCell wsDest, 18, dCol, wsAdj.Cells(14, pCol).Value ' Miscellaneous Rev
        ' Row 19 Total Revenue is a formula already in destination -> skip

        ' Rows 22-26 Department Expense: already formulas -> skip

        ' Department Profits (rows 18-21 in Adj P&L block)
        WriteCell wsDest, 29, dCol, wsAdj.Cells(18, pCol).Value ' Rooms profit
        WriteCell wsDest, 30, dCol, wsAdj.Cells(19, pCol).Value ' F&B profit
        WriteCell wsDest, 31, dCol, wsAdj.Cells(20, pCol).Value ' Other Op. Departments profit
        WriteCell wsDest, 32, dCol, wsAdj.Cells(21, pCol).Value ' Miscellaneous profit
        ' Row 33 Total Profit is a formula -> skip

        ' Undistributed Expenses (rows 25-29 in Adj P&L block)
        WriteCell wsDest, 36, dCol, wsAdj.Cells(25, pCol).Value ' Admin & General
        WriteCell wsDest, 37, dCol, wsAdj.Cells(26, pCol).Value ' Info & Telecom
        WriteCell wsDest, 38, dCol, wsAdj.Cells(27, pCol).Value ' Sales & Marketing
        WriteCell wsDest, 39, dCol, wsAdj.Cells(28, pCol).Value ' Repairs & Maintenance
        WriteCell wsDest, 40, dCol, wsAdj.Cells(29, pCol).Value ' Utilities
        ' Row 41 Total Undistributed is a formula -> skip
        ' Row 43 Gross Op Profit is a formula -> skip

        WriteCell wsDest, 44, dCol, wsAdj.Cells(34, pCol).Value ' Management Fees
        ' Row 45 Income Before Fixed is a formula -> skip

        ' Non Op Expenses (rows 39-42 in Adj P&L block)
        WriteCell wsDest, 48, dCol, wsAdj.Cells(39, pCol).Value ' Other Non Op. Revenues
        WriteCell wsDest, 49, dCol, wsAdj.Cells(40, pCol).Value ' Rent
        WriteCell wsDest, 50, dCol, wsAdj.Cells(41, pCol).Value ' Taxes
        WriteCell wsDest, 51, dCol, wsAdj.Cells(42, pCol).Value ' Insurance
        ' Row 52 Total Non Op is a formula -> skip
        ' Row 54 Pre-Fee EBITDA = R45 - R52 formula -> skip
        ' Row 56 EBITDA = SUM(row 54) formula -> skip
        ' Row 57 Reserve: leave BLANK (matches Houston Marriott pattern)
        ' Row 58 NOI = R56 - R57 formula -> skip
        ' Row 60 Source: leave BLANK
        ' Row 61 Variance = R58 - R60 formula -> skip

        ' --- Adjustment lines (rows 65-74) ---
        ' Pull "Total X Adjustments" sums from the adjustment block (cols T-Z)
        ' Rooms and F&B are sign-flipped (dest convention: negative = expense add)
        If FILL_ADJUSTMENTS Then
            WriteCell wsDest, 65, dCol, NegNumOrBlank(wsAdj.Cells(59, aCol).Value)  ' -Total Rooms Adj
            WriteCell wsDest, 66, dCol, NegNumOrBlank(wsAdj.Cells(73, aCol).Value)  ' -Total F&B Adj
            WriteCell wsDest, 67, dCol, wsAdj.Cells(84, aCol).Value                  ' Total A&G Adj
            WriteCell wsDest, 68, dCol, wsAdj.Cells(90, aCol).Value                  ' Total S&M Adj
            WriteCell wsDest, 69, dCol, wsAdj.Cells(97, aCol).Value                  ' Total R&M Adj

            ' Row 70 Management Fees adjustment = OM mgmt - Adj mgmt
            Dim omV As Variant, adjV As Variant
            If Not wsOM Is Nothing Then
                omV = wsOM.Cells(38, pCol + 1).Value   ' OM P&L col is one to the right (G for 2025)
            Else
                omV = Empty
            End If
            adjV = wsAdj.Cells(34, pCol).Value
            If IsNumeric(omV) And IsNumeric(adjV) Then
                WriteCell wsDest, 70, dCol, CDbl(omV) - CDbl(adjV)
            End If

            WriteCell wsDest, 71, dCol, wsAdj.Cells(103, aCol).Value ' Total Other Non-Op Exp Adj
            WriteCell wsDest, 72, dCol, wsAdj.Cells(109, aCol).Value ' Total Rent Adj
            WriteCell wsDest, 73, dCol, wsAdj.Cells(116, aCol).Value ' Total Tax Adj
            WriteCell wsDest, 74, dCol, wsAdj.Cells(121, aCol).Value ' Total Insurance Adj
        End If
        ' Row 75 Total Adjustments = SUM(R65:R74) formula -> skip
        ' Row 78 ADJUSTED NOI = R58 - R75 formula -> skip
    Next k

End Sub

'--------------------------------------------------------------------------
' Write a value to a cell and (if HIGHLIGHT_YELLOW is on) shade it yellow.
' Blank/empty input clears the value but still highlights.
'--------------------------------------------------------------------------
Private Sub WriteCell(ws As Worksheet, r As Long, c As Long, v As Variant)
    ws.Cells(r, c).Value = v
    ' YELLOW_HIGHLIGHT: shade modified cells so it's obvious what changed.
    If HIGHLIGHT_YELLOW Then
        ws.Cells(r, c).Interior.Color = YELLOW_COLOR
    End If
End Sub

'--------------------------------------------------------------------------
' Negate a numeric value; return Empty if input is non-numeric.
'--------------------------------------------------------------------------
Private Function NegNumOrBlank(v As Variant) As Variant
    If IsNumeric(v) And Not IsEmpty(v) Then
        NegNumOrBlank = -CDbl(v)
    Else
        NegNumOrBlank = Empty
    End If
End Function

'--------------------------------------------------------------------------
' Read row 3 of destination ("Charlotte Airport Hilton", ...) and return
' a dictionary: property name -> the column where its "2021" lives
' (i.e., the first column of that property's 7-column block).
'--------------------------------------------------------------------------
Private Function BuildPropertyColumnMap(wsDest As Worksheet) As Object

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1   ' case-insensitive text compare

    Dim lastCol As Long
    lastCol = wsDest.Cells(3, wsDest.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then lastCol = 200   ' fallback if End() misbehaves

    Dim c As Long
    For c = 1 To lastCol
        Dim name As String
        name = Trim$(CStr(wsDest.Cells(3, c).Value))
        Dim yr As Variant
        yr = wsDest.Cells(4, c).Value
        ' Only record the FIRST column of each property block (the "2021" col)
        If Len(name) > 0 And IsNumeric(yr) Then
            If CLng(yr) = 2021 Then
                If Not d.Exists(name) Then d.Add name, c
            End If
        End If
    Next c

    Set BuildPropertyColumnMap = d

End Function

'--------------------------------------------------------------------------
' Extract a clean property name from a filename like:
'   "3.5.4.10 260520 - Houston Marriott P&L.xlsx"  ->  "Houston Marriott"
'   "3.5.4.2 260520 - Charlotte Airport Hilton P&L.xlsx"
'                                          ->  "Charlotte Airport Hilton"
' Falls back to filename-without-extension if the pattern doesn't match.
'--------------------------------------------------------------------------
Private Function PropertyNameFromFileName(fn As String) As String

    Dim s As String
    s = fn
    ' Drop extension
    Dim dotPos As Long
    dotPos = InStrRev(s, ".")
    If dotPos > 0 Then s = Left$(s, dotPos - 1)

    ' Drop everything up to and including " - "
    Dim dashPos As Long
    dashPos = InStr(1, s, " - ")
    If dashPos > 0 Then s = Mid$(s, dashPos + 3)

    ' Drop trailing " P&L" or " P_L" (filenames that use _ instead of &)
    s = TrimSuffix(s, " P&L")
    s = TrimSuffix(s, " P_L")
    s = TrimSuffix(s, " P L")
    s = Trim$(s)

    PropertyNameFromFileName = s

End Function

Private Function TrimSuffix(s As String, suffix As String) As String
    If Len(s) >= Len(suffix) Then
        If StrComp(Right$(s, Len(suffix)), suffix, vbTextCompare) = 0 Then
            TrimSuffix = Left$(s, Len(s) - Len(suffix))
            Exit Function
        End If
    End If
    TrimSuffix = s
End Function

'--------------------------------------------------------------------------
' Folder picker dialog
'--------------------------------------------------------------------------
Private Function PickFolder() As String

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Select the folder containing P&L source files"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = ""
    End If

End Function
