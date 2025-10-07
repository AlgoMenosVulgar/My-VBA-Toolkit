Option Explicit

Sub ApplyPctToLower_FromImport_FinalPolished()
    Const PRICES_SHEET   As String = "Prices"
    Const IMPORT_SHEET   As String = "Import"
    Const HEADER_TXT     As String = "% to Lower"
    Const STEP_PCT       As Double = 5

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsPrices As Worksheet: Set wsPrices = wb.Sheets(PRICES_SHEET)
    Dim wsImport As Worksheet: Set wsImport = wb.Sheets(IMPORT_SHEET)

    ' === Detect supplier range ===
    Dim firstSupCol As Long: firstSupCol = 4
    Dim endColCell As Range, lastSupCol As Long
    Set endColCell = wsPrices.Rows(1).Find("end", , xlValues, xlWhole)
    If Not endColCell Is Nothing Then
        lastSupCol = endColCell.Column - 1
    Else
        lastSupCol = wsPrices.Cells(1, wsPrices.Columns.Count).End(xlToLeft).Column
    End If

    ' === Detect rows ===
    Dim endRowCell As Range, finalRowPrices As Long
    Set endRowCell = wsPrices.Columns(1).Find("end", , xlValues, xlWhole)
    If Not endRowCell Is Nothing Then
        finalRowPrices = endRowCell.Row - 1
    Else
        finalRowPrices = wsPrices.Cells(wsPrices.Rows.Count, 1).End(xlUp).Row
    End If
    Dim initRowPrices As Long: initRowPrices = 2
    Dim numRowsPrices As Long: numRowsPrices = finalRowPrices - initRowPrices + 1
    If numRowsPrices <= 0 Then Exit Sub

    ' === Compute minima per row ===
    Dim minRow() As Double: ReDim minRow(1 To numRowsPrices)
    Dim r As Long, c As Long, v As Variant, mn As Double, found As Boolean
    For r = 1 To numRowsPrices
        found = False
        For c = firstSupCol To lastSupCol
            v = wsPrices.Cells(initRowPrices + r - 1, c).Value
            If IsNumeric(v) And v > 0 Then
                If Not found Then mn = v: found = True Else If v < mn Then mn = v
            End If
        Next c
        minRow(r) = IIf(found, mn, 0)
    Next r

    ' === Supplier names from Import row 1 ===
    Dim supNames As Collection: Set supNames = New Collection
    c = 2
    Do While Len(Trim$(wsImport.Cells(1, c).Value)) > 0
        If LCase$(Trim$(wsImport.Cells(1, c).Value)) = "end" Then Exit Do
        supNames.Add CStr(wsImport.Cells(1, c).Value)
        c = c + 1
    Loop

    ' === Destination placement ===
    Dim baseColLetter As String, pasteCol As Long, pasteStart As Long, pasteEnd As Long
    baseColLetter = Trim$(CStr(wsImport.Range("I2").Value))
    If Len(baseColLetter) = 0 Then baseColLetter = "R"
    pasteCol = ColLetterToIndex(baseColLetter) + 1
    pasteStart = CLng(wsImport.Range("J2").Value)
    pasteEnd = CLng(wsImport.Range("K2").Value)
    Dim pasteNumRows As Long: pasteNumRows = pasteEnd - pasteStart + 1
    If pasteNumRows <= 0 Then Exit Sub

    ' === Map supplier columns ===
    Dim priceCol As Object: Set priceCol = CreateObject("Scripting.Dictionary")
    For c = firstSupCol To lastSupCol
        Dim name As String: name = Trim$(CStr(wsPrices.Cells(1, c).Value))
        If Len(name) > 0 Then priceCol(name) = c
    Next c

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim i As Long, wsSup As Worksheet, label As String, pct As Double, base As Double
    Dim arr() As Variant: ReDim arr(1 To numRowsPrices, 1 To 1)

    For i = 1 To supNames.Count
        name = supNames(i)
        If Not priceCol.exists(name) Then GoTo NextSupplier
        On Error Resume Next: Set wsSup = wb.Sheets(name): On Error GoTo 0
        If wsSup Is Nothing Then GoTo NextSupplier

        ' === Header one row above Import pasteStart ===
        Dim headerCell As Range
        Set headerCell = wsSup.Cells(pasteStart - 1, pasteCol)
        headerCell.Value = HEADER_TXT
        headerCell.Interior.Color = RGB(255, 235, 156)
        headerCell.Font.Color = RGB(0, 0, 0)
        headerCell.Font.Bold = True
        headerCell.HorizontalAlignment = xlCenter

        ' === Compute labels ===
        For r = 1 To numRowsPrices
            v = wsPrices.Cells(initRowPrices + r - 1, priceCol(name)).Value
            label = "NA"
            If IsNumeric(v) And v > 0 And minRow(r) > 0 Then
                If v <= minRow(r) Then
                    label = "Good"
                Else
                    pct = (v - minRow(r)) / v * 100
                    If pct > 90 Then
                        label = "90 - 95%"
                    Else
                        base = CeilToStep(pct, STEP_PCT)
                        label = base & " - " & base + STEP_PCT & "%"
                    End If
                End If
            End If
            arr(r, 1) = label
        Next r

        ' === Write directly into Import-defined range ===
        Dim rng As Range
        Set rng = wsSup.Range(wsSup.Cells(pasteStart, pasteCol), wsSup.Cells(pasteEnd, pasteCol))
        rng.Value = arr

        ' === Borders and centering ===
        With wsSup.Range(headerCell, rng)
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .Borders.Color = RGB(0, 0, 0)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' === Conditional formatting (Feedback colors) ===
        rng.FormatConditions.Delete
        Dim cf1 As FormatCondition, cf2 As FormatCondition, cf3 As FormatCondition
        Set cf1 = rng.FormatConditions.Add(xlCellValue, xlEqual, "=""NA""")
        cf1.Interior.Color = RGB(217, 217, 217)
        cf1.Font.Color = RGB(0, 0, 0)
        Set cf2 = rng.FormatConditions.Add(xlCellValue, xlEqual, "=""Good""")
        cf2.Interior.Color = RGB(198, 239, 206)
        cf2.Font.Color = RGB(0, 97, 0)
        Set cf3 = rng.FormatConditions.Add(xlCellValue, xlNotEqual, "=""Good""")
        cf3.Interior.Color = RGB(255, 199, 206)
        cf3.Font.Color = RGB(156, 0, 6)

        ' === AutoFit column width ===
        wsSup.Columns(pasteCol).AutoFit

NextSupplier:
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' === Helpers ===
Private Function CeilToStep(x As Double, stepSz As Double) As Double
    CeilToStep = stepSz * WorksheetFunction.RoundUp(x / stepSz, 0)
End Function

Private Function ColLetterToIndex(s As String) As Long
    Dim i As Long: s = UCase(Trim(s))
    For i = 1 To Len(s)
        ColLetterToIndex = ColLetterToIndex * 26 + Asc(Mid$(s, i, 1)) - 64
    Next i
End Function
