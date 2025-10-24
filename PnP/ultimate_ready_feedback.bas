Option Explicit

Sub ApplyPctToLower_Simple()
    Const PRICES_SHEET As String = "Prices"
    Const IMPORT_SHEET As String = "Import"
    Const HEADER_TXT   As String = "% to Lower"
    Const STEP_PCT     As Double = 5
    
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsP As Worksheet, wsI As Worksheet
    Set wsP = wb.Sheets(PRICES_SHEET)
    Set wsI = wb.Sheets(IMPORT_SHEET)
    
    ' --- dynamic controls to the right of "end" in Import!row 1 ---
    Dim endHdr As Range: Set endHdr = wsI.Rows(1).Find("end", LookIn:=xlValues, LookAt:=xlWhole)
    If endHdr Is Nothing Then Err.Raise vbObjectError + 1, , "Import row 1 has no 'end'."
    Dim colLetterC As Long: colLetterC = endHdr.Column + 1   ' row 2: column letter
    Dim startRowC  As Long: startRowC  = endHdr.Column + 2   ' row 2: start row
    Dim endRowC    As Long: endRowC    = endHdr.Column + 3   ' row 2: end row
    
    Dim baseLetter As String: baseLetter = LettersOnly(CStr(wsI.Cells(2, colLetterC).Value))
    If Len(baseLetter) = 0 Then baseLetter = "R"
    Dim pasteCol As Long: pasteCol = ColLetterToIndex(baseLetter) + 1
    Dim pasteStart As Long: pasteStart = CLng0(wsI.Cells(2, startRowC).Value)
    Dim pasteEnd   As Long: pasteEnd   = CLng0(wsI.Cells(2, endRowC).Value)
    If pasteEnd < pasteStart Then Err.Raise vbObjectError + 2, , "End row < Start row."
    
    ' --- suppliers (Import!B1.. until "end") ---
    Dim sup As Collection: Set sup = New Collection
    Dim c As Long: c = 2
    Do While Len(Trim$(wsI.Cells(1, c).Value)) > 0
        If LCase$(Trim$(wsI.Cells(1, c).Value)) = "end" Then Exit Do
        sup.Add CStr(wsI.Cells(1, c).Value)
        c = c + 1
    Loop
    If sup.Count = 0 Then Err.Raise vbObjectError + 3, , "No suppliers in Import row 1."
    
    ' --- prices table bounds + minima per row ---
    Dim firstSupCol As Long: firstSupCol = 4 ' D
    Dim lastSupCol As Long
    Dim endColCell As Range: Set endColCell = wsP.Rows(1).Find("end", LookIn:=xlValues, LookAt:=xlWhole)
    lastSupCol = IIf(endColCell Is Nothing, wsP.Cells(1, wsP.Columns.Count).End(xlToLeft).Column, endColCell.Column - 1)
    
    Dim endRowCell As Range: Set endRowCell = wsP.Columns(1).Find("end", LookIn:=xlValues, LookAt:=xlWhole)
    Dim finalRow As Long: finalRow = IIf(endRowCell Is Nothing, wsP.Cells(wsP.Rows.Count, 1).End(xlUp).Row, endRowCell.Row - 1)
    Dim initRow As Long: initRow = 2
    
    Dim pricesRows As Long: pricesRows = Application.Max(0, finalRow - initRow + 1)
    If pricesRows = 0 Then Err.Raise vbObjectError + 4, , "No data rows in Prices."
    
    Dim minRow() As Double: ReDim minRow(1 To pricesRows)
    Dim r As Long, v As Variant, mn As Double, found As Boolean
    For r = 1 To pricesRows
        found = False
        For c = firstSupCol To lastSupCol
            v = wsP.Cells(initRow + r - 1, c).Value
            If IsNumeric(v) And v > 0 Then
                If Not found Then mn = v: found = True Else If v < mn Then mn = v
            End If
        Next c
        minRow(r) = IIf(found, mn, 0)
    Next r
    
    ' --- map supplier name -> Prices column (case-insensitive) ---
    Dim mapCol As Object: Set mapCol = CreateObject("Scripting.Dictionary")
    mapCol.CompareMode = 1 ' vbTextCompare
    For c = firstSupCol To lastSupCol
        v = Trim$(CStr(wsP.Cells(1, c).Value))
        If Len(v) > 0 Then mapCol(v) = c
    Next c
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim i As Long, wsS As Worksheet, nameS As String
    Dim nRows As Long: nRows = Application.Min(pricesRows, pasteEnd - pasteStart + 1)
    Dim arr() As Variant: ReDim arr(1 To nRows, 1 To 1)
    
    For i = 1 To sup.Count
        nameS = CStr(sup(i))
        If Not mapCol.Exists(nameS) Then GoTo NextSheet
        
        On Error Resume Next
        Set wsS = wb.Sheets(nameS)
        On Error GoTo 0
        If wsS Is Nothing Or wsS.ProtectContents Then GoTo NextSheet
        
        ' 1) Insert fresh column at pasteCol and wipe any CF on that column
        wsS.Columns(pasteCol).Insert Shift:=xlToRight
        wsS.Columns(pasteCol).FormatConditions.Delete
        
        ' 2) Build labels for nRows
        Dim pct As Double, base As Double, label As String
        For r = 1 To nRows
            v = wsP.Cells(initRow + r - 1, mapCol(nameS)).Value
            If VarType(v) = vbString And Trim$(CStr(v)) = "Blank" Then
                arr(r, 1) = vbNullString
            ElseIf IsNumeric(v) And v > 0 And minRow(r) > 0 Then
                If v <= minRow(r) Then
                    arr(r, 1) = "Good"
                Else
                    pct = (v - minRow(r)) / v * 100    ' if you prefer vs. minimum: /(minRow(r))*100
                    If pct > 90 Then
                        arr(r, 1) = "90 - 95%"
                    Else
                        base = CeilToStep(pct, STEP_PCT)
                        arr(r, 1) = CStr(base) & " - " & CStr(base + STEP_PCT) & "%"
                    End If
                End If
            Else
                arr(r, 1) = "NA"
            End If
        Next r
        
        ' 3) Paste values
        Dim rng As Range
        Set rng = wsS.Range(wsS.Cells(pasteStart, pasteCol), wsS.Cells(pasteStart + nRows - 1, pasteCol))
        rng.Value = arr
        
        ' 4) Truly blank out rows whose source was "Blank"
        For r = 1 To nRows
            v = wsP.Cells(initRow + r - 1, mapCol(nameS)).Value
            If VarType(v) = vbString And Trim$(CStr(v)) = "Blank" Then
                With wsS.Cells(pasteStart + r - 1, pasteCol)
                    .FormatConditions.Delete
                    .ClearFormats
                    .ClearContents
                End With
            End If
        Next r
        On Error Resume Next
        rng.SpecialCells(xlCellTypeBlanks).FormatConditions.Delete
        rng.SpecialCells(xlCellTypeBlanks).ClearFormats
        On Error GoTo 0
        
        ' 5) Place header BEFORE each non-empty block
        Dim wasEmpty As Boolean: wasEmpty = True
        Dim rCell As Range, hdrRow As Long
        For r = 1 To nRows
            Set rCell = wsS.Cells(pasteStart + r - 1, pasteCol)
            If Len(rCell.Value) > 0 Then
                If wasEmpty Then
                    hdrRow = pasteStart + r - 2
                    If hdrRow >= 1 Then wsS.Cells(hdrRow, pasteCol).Value = HEADER_TXT
                    wasEmpty = False
                End If
            Else
                wasEmpty = True
            End If
        Next r
        
        ' 6) Borders on non-empty constants only
        On Error Resume Next
        wsS.Range(rng.Address).SpecialCells(xlCellTypeConstants).Borders.LineStyle = xlContinuous
        wsS.Range(rng.Address).SpecialCells(xlCellTypeConstants).Borders.Weight = xlThin
        wsS.Range(rng.Address).SpecialCells(xlCellTypeConstants).HorizontalAlignment = xlCenter
        wsS.Range(rng.Address).SpecialCells(xlCellTypeConstants).VerticalAlignment = xlCenter
        On Error GoTo 0
        
        ' 7) Conditional formatting for data cells only (ignore blanks)
        rng.FormatConditions.Delete
        Dim firstA1 As String: firstA1 = rng.Cells(1, 1).Address(False, False)
        Dim sep As String: sep = Application.International(xlListSeparator)
        Dim fNA As String, fGood As String, fOther As String
        fNA = "=AND(LEN(" & firstA1 & ")>0" & sep & firstA1 & "=""" & "NA" & """)"
        fGood = "=AND(LEN(" & firstA1 & ")>0" & sep & firstA1 & "=""" & "Good" & """)"
        fOther = "=AND(LEN(" & firstA1 & ")>0" & sep & "NOT(OR(" & firstA1 & "=""" & "Good" & """" & sep & firstA1 & "=""" & "NA" & """)))"
        
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fNA)
            .Interior.Color = RGB(217, 217, 217): .Font.Color = RGB(0, 0, 0)
        End With
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fGood)
            .Interior.Color = RGB(198, 239, 206): .Font.Color = RGB(0, 97, 0)
        End With
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:=fOther)
            .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
        End With
        
        ' 8) FINAL SWEEP: make every "% to Lower" cell static & yellowish (no CF, no stray formats)
        ForceHeaderLook wsS, pasteCol, pasteStart, nRows, HEADER_TXT
        
        wsS.Columns(pasteCol).AutoFit
        
NextSheet:
        Set wsS = Nothing
    Next i

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' --- FINAL HEADER LOOK: erase formats & CF wherever the header text exists, then paint uniform yellow ---
Private Sub ForceHeaderLook(ws As Worksheet, ByVal colIdx As Long, ByVal startRow As Long, ByVal nRows As Long, ByVal headerTxt As String)
    Dim topRow As Long: topRow = Application.Max(1, startRow - 50) ' generous scan up for all header rows
    Dim bottomRow As Long: bottomRow = startRow + nRows - 1
    Dim rngScan As Range
    Set rngScan = ws.Range(ws.Cells(topRow, colIdx), ws.Cells(bottomRow, colIdx))
    
    Dim cell As Range
    For Each cell In rngScan
        If CStr(cell.Value) = headerTxt Then
            cell.FormatConditions.Delete
            cell.ClearFormats
            cell.Value = headerTxt
            ' static yellowish look (no CF)
            cell.Font.Bold = True
            cell.Font.Color = RGB(0, 0, 0)
            cell.Interior.Color = RGB(255, 235, 156)
            cell.HorizontalAlignment = xlCenter
            cell.VerticalAlignment = xlCenter
            cell.WrapText = False
            cell.Borders.LineStyle = xlLineStyleNone
        End If
    Next cell
End Sub

' --- utilities ---
Private Function LettersOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then out = out & ch
    Next i
    LettersOnly = out
End Function

Private Function ColLetterToIndex(ByVal s As String) As Long
    Dim i As Long: s = UCase$(Trim$(s))
    For i = 1 To Len(s)
        ColLetterToIndex = ColLetterToIndex * 26 + Asc(Mid$(s, i, 1)) - 64
    Next i
End Function

Private Function CLng0(ByVal v As Variant) As Long
    If IsNumeric(v) Then CLng0 = CLng(v) Else CLng0 = 0
End Function

Private Function CeilToStep(ByVal x As Double, ByVal stepSz As Double) As Double
    CeilToStep = stepSz * WorksheetFunction.RoundUp(x / stepSz, 0)
End Function
