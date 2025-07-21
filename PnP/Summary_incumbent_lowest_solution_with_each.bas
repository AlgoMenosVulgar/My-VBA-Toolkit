Option Explicit

'–– CONSTANTS
Public Const HDR_ROW            As Long = 1       ' Header row on Summary
Public Const COL_VOL            As Long = 2       ' Volume      (B)
Public Const COL_BASE           As Long = 3       ' Baseline    (C)
Public Const COL_LABELS         As Long = 4       ' Label/blank (D)
Public Const FIRST_SUP_COL      As Long = 4       ' First supplier (E)
Public Const COLUMNS_PER_BLOCK  As Long = 6       ' Supplier, Unit Price, Total Price, Baseline, Savings $, Savings %
Public Const FIRST_DATA_ROW     As Long = 3       ' First data row on Summary

Sub Analytics_With_Baseline_IncumbentLowestLSI_Final()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim wsData As Worksheet: Set wsData = wb.Sheets(1)
    Dim wsPrices As Worksheet: Set wsPrices = wb.Sheets("Prices")
    Dim wsSummary As Worksheet
    Dim lastRow As Long, supplierEnd As Long, r As Long, currentRow As Long
    Dim hdrRange As String, dataRange As String, hdrLow As String, dataLow As String, vals As String
    Dim dropdown As Range
    Dim containsBlank As Boolean
    Dim incumbentStart As Long, incumbentEnd As Long, lowestStart As Long, lowestEnd As Long, lsiStart As Long, lsiEnd As Long
    Dim finalSummaryStartCol As Long, finalSummaryEndCol As Long

    '— EXCEL SETTINGS —
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .ErrorCheckingOptions.NumberAsText = False
    End With
    Application.Calculation = xlCalculationAutomatic

    '— SUMMARY SHEET —
    On Error Resume Next: Set wsSummary = wb.Sheets("Summary"): On Error GoTo 0
    If wsSummary Is Nothing Then
        Set wsSummary = wb.Sheets.Add(After:=wsData)
        wsSummary.Name = "Summary"
    Else
        wsSummary.Cells.Clear
    End If
    wsSummary.Tab.Color = RGB(128, 0, 128) ' purple

    lastRow = wsPrices.Cells(wsPrices.Rows.Count, 1).End(xlUp).Row
    supplierEnd = wsData.Rows(HDR_ROW).Find(What:="End", LookIn:=xlValues, LookAt:=xlWhole).Column - 1

    '================================================================
    ' 1) VOLUME & BASELINE
    '================================================================
    Call HeaderCell(wsSummary.Cells(HDR_ROW + 1, COL_VOL), "Volume")
    Call HeaderCell(wsSummary.Cells(HDR_ROW + 1, COL_BASE), "Baseline")
    For r = 2 To lastRow - 1
        If wsPrices.Cells(r, "B").Value <> "Blank" Then
            With wsSummary.Cells(r + 1, COL_VOL)
                .Formula = "=Prices!B" & r
                .NumberFormat = "#,##0"
                Call Bordered(.Cells)
            End With
        End If
        If wsPrices.Cells(r, "C").Value <> "Blank" Then
            With wsSummary.Cells(r + 1, COL_BASE)
                .Formula = "=IF(Prices!C" & r & "=" & """NA""" & ",""NA"",Prices!C" & r & "*Prices!B" & r & ")"
                .NumberFormat = "$#,##0.00"
                Call Bordered(.Cells)
            End With
        End If
    Next r

    '================================================================
    ' 2) INCUMBENT SOLUTION
    '================================================================
    incumbentStart = COL_BASE + 2
    incumbentEnd = incumbentStart + COLUMNS_PER_BLOCK - 1
    hdrRange = wsPrices.Range(wsPrices.Cells(HDR_ROW, FIRST_SUP_COL), wsPrices.Cells(HDR_ROW, supplierEnd)).Address(False, False, xlA1, True)

    ' Title & subheaders
    With wsSummary.Range(wsSummary.Cells(HDR_ROW, incumbentStart), wsSummary.Cells(HDR_ROW, incumbentEnd))
        .Merge: .Value = "Incumbent Solution"
        .Interior.Color = RGB(255, 192, 0): .Borders.Weight = xlThin
    End With
    With wsSummary.Range(wsSummary.Cells(HDR_ROW + 1, incumbentStart), wsSummary.Cells(HDR_ROW + 1, incumbentEnd))
        .Value = Array("Supplier", "Unit Price", "Total Price", "Baseline", "Savings $", "Savings %")
        .Interior.Color = RGB(202, 237, 251): .Borders.Weight = xlThin
    End With

    currentRow = HDR_ROW + 2
    For r = 2 To lastRow
        If wsPrices.Cells(r, 1).Value = "end" Then Exit For
        If wsPrices.Cells(r, FIRST_SUP_COL).Value <> "Blank" Then
            wsSummary.Cells(currentRow, incumbentStart).Formula = "=Prices!A" & r
            Call Bordered(wsSummary.Cells(currentRow, incumbentStart))

            dataRange = wsPrices.Range(wsPrices.Cells(r, FIRST_SUP_COL), wsPrices.Cells(r, supplierEnd)).Address(False, False, xlA1, True)
            ' Unit Price
            With wsSummary.Cells(currentRow, incumbentStart + 1)
                .Formula = "=IFERROR(INDEX(" & dataRange & ",MATCH(" & _
                           wsSummary.Cells(currentRow, incumbentStart).Address(False, False) & "," & _
                           hdrRange & ",0)),""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            ' Total Price
            With wsSummary.Cells(currentRow, incumbentStart + 2)
                .Formula = "=IFERROR(" & wsSummary.Cells(currentRow, incumbentStart + 1).Address(False, False) & "*" & _
                           wsSummary.Cells(currentRow, COL_VOL).Address(False, False) & ",""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            ' Baseline
            With wsSummary.Cells(currentRow, incumbentStart + 3)
                .Formula = "=IFERROR(Prices!" & wsPrices.Cells(r, COL_BASE).Address(False, True) & "*" & _
                           wsSummary.Cells(currentRow, COL_VOL).Address(False, False) & ",""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            ' Savings $ & %
            Call AddSavings(wsSummary.Cells(currentRow, incumbentStart + 4), _
                            wsSummary.Cells(currentRow, incumbentStart + 3), _
                            wsSummary.Cells(currentRow, incumbentStart + 2))
            Call AddPercent(wsSummary.Cells(currentRow, incumbentStart + 5), _
                            wsSummary.Cells(currentRow, incumbentStart + 3), _
                            wsSummary.Cells(currentRow, incumbentStart + 2))
        End If
        currentRow = currentRow + 1
    Next r

    '================================================================
    ' 3) LOWEST SOLUTION
    '================================================================
    lowestStart = incumbentEnd + 2: lowestEnd = lowestStart + COLUMNS_PER_BLOCK - 1
    Set dropdown = wsSummary.Cells(HDR_ROW, lowestStart)
    With dropdown.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="1,2,3,4,5"
        .InCellDropdown = True
    End With
    Call Bordered(dropdown)

    With wsSummary.Range(wsSummary.Cells(HDR_ROW, lowestStart + 1), wsSummary.Cells(HDR_ROW, lowestEnd))
        .Merge: .Value = "Lowest Solution"
        .Interior.Color = RGB(255, 192, 0): .Borders.Weight = xlThin
    End With
    With wsSummary.Range(wsSummary.Cells(HDR_ROW + 1, lowestStart), wsSummary.Cells(HDR_ROW + 1, lowestEnd))
        .Value = Array("Supplier", "Unit Price", "Total Price", "Baseline", "Savings $", "Savings %")
        .Interior.Color = RGB(202, 237, 251): .Borders.Weight = xlThin
    End With

    currentRow = HDR_ROW + 2
    For r = 2 To lastRow
        If wsPrices.Cells(r, 1).Value = "end" Then Exit For

        hdrLow = wsPrices.Range(wsPrices.Cells(HDR_ROW, FIRST_SUP_COL), wsPrices.Cells(HDR_ROW, supplierEnd)).Address(False, False)
        dataLow = wsPrices.Range(wsPrices.Cells(r, FIRST_SUP_COL), wsPrices.Cells(r, supplierEnd)).Address(False, False)
        containsBlank = Application.CountIf( _
                        wsPrices.Range(wsPrices.Cells(r, FIRST_SUP_COL), wsPrices.Cells(r, supplierEnd)), "Blank") > 0

        If Not containsBlank Then
            With wsSummary.Cells(currentRow, lowestStart + 1)
                .Formula = "=IFERROR(SMALL(Prices!" & dataLow & "," & _
                           wsSummary.Cells(HDR_ROW, lowestStart).Address(True, True) & "),""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            With wsSummary.Cells(currentRow, lowestStart + 2)
                .Formula = "=IFERROR(" & wsSummary.Cells(currentRow, lowestStart + 1).Address(False, False) & "*" & _
                           wsSummary.Cells(currentRow, COL_VOL).Address(False, False) & ",""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            With wsSummary.Cells(currentRow, lowestStart)
                .Formula = "=IFERROR(INDEX(Prices!" & hdrLow & ",MATCH(" & _
                           wsSummary.Cells(currentRow, lowestStart + 1).Address(False, False) & _
                           ",Prices!" & dataLow & ",0)),""NA"")"
                Call Bordered(.Cells)
            End With
            With wsSummary.Cells(currentRow, lowestStart + 3)
                .Formula = "=IFERROR(Prices!" & wsPrices.Cells(r, COL_BASE).Address(False, True) & "*" & _
                           wsSummary.Cells(currentRow, COL_VOL).Address(False, False) & ",""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            Call AddSavings(wsSummary.Cells(currentRow, lowestStart + 4), _
                            wsSummary.Cells(currentRow, lowestStart + 3), _
                            wsSummary.Cells(currentRow, lowestStart + 2))
            Call AddPercent(wsSummary.Cells(currentRow, lowestStart + 5), _
                            wsSummary.Cells(currentRow, lowestStart + 3), _
                            wsSummary.Cells(currentRow, lowestStart + 2))
        End If
        currentRow = currentRow + 1
    Next r

    '================================================================
    ' 4) LSI SOLUTION
    '================================================================
    lsiStart = lowestEnd + 2: lsiEnd = lsiStart + COLUMNS_PER_BLOCK - 1

    With wsSummary.Range(wsSummary.Cells(HDR_ROW, lsiStart), wsSummary.Cells(HDR_ROW, lsiEnd))
        .Merge: .Value = "LSI Solution"
        .Interior.Color = RGB(255, 192, 0): .Borders.Weight = xlThin
    End With
    With wsSummary.Range(wsSummary.Cells(HDR_ROW + 1, lsiStart), wsSummary.Cells(HDR_ROW + 1, lsiEnd))
        .Value = Array("Supplier", "Unit Price", "Total Price", "Baseline", "Savings $", "Savings %")
        .Interior.Color = RGB(202, 237, 251): .Borders.Weight = xlThin
    End With

    currentRow = HDR_ROW + 2
    For r = 2 To lastRow
        If wsPrices.Cells(r, 1).Value = "end" Then Exit For
        If wsPrices.Cells(r, FIRST_SUP_COL).Value <> "Blank" Then
            vals = Join(Application.Index( _
                   wsPrices.Range(wsPrices.Cells(HDR_ROW, FIRST_SUP_COL), _
                                  wsPrices.Cells(HDR_ROW, supplierEnd)).Value, 1, 0), ",")
            With wsSummary.Cells(currentRow, lsiStart)
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, Formula1:=vals
                Call Bordered(.Cells)
            End With
            With wsSummary.Cells(currentRow, lsiStart + 1)
                .Formula = "=IFERROR(INDEX(Prices!" & _
                           wsPrices.Cells(r, FIRST_SUP_COL).Address(False, False) & ":" & _
                           wsPrices.Cells(r, supplierEnd).Address(False, False) & "," & _
                           "MATCH(" & wsSummary.Cells(currentRow, lsiStart).Address(False, False) & _
                           ",Prices!" & _
                           wsPrices.Cells(HDR_ROW, FIRST_SUP_COL).Address(False, False) & ":" & _
                           wsPrices.Cells(HDR_ROW, supplierEnd).Address(False, False) & ",0)),""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            With wsSummary.Cells(currentRow, lsiStart + 2)
                .Formula = "=IFERROR(" & wsSummary.Cells(currentRow, lsiStart + 1).Address(False, False) & "*" & _
                           wsSummary.Cells(currentRow, COL_VOL).Address(False, False) & ",""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            With wsSummary.Cells(currentRow, lsiStart + 3)
                .Formula = "=IFERROR(Prices!" & wsPrices.Cells(r, COL_BASE).Address(False, True) & "*" & _
                           wsSummary.Cells(currentRow, COL_VOL).Address(False, False) & ",""NA"")"
                .NumberFormat = "$#,##0.00": Call Bordered(.Cells)
            End With
            Call AddSavings(wsSummary.Cells(currentRow, lsiStart + 4), _
                            wsSummary.Cells(currentRow, lsiStart + 3), _
                            wsSummary.Cells(currentRow, lsiStart + 2))
            Call AddPercent(wsSummary.Cells(currentRow, lsiStart + 5), _
                            wsSummary.Cells(currentRow, lsiStart + 3), _
                            wsSummary.Cells(currentRow, lsiStart + 2))
        End If
        currentRow = currentRow + 1
    Next r

    '================================================================
    ' ONE GLOBAL SUMMARY INSERTION AFTER ALL DATA IS POPULATED
    '================================================================
    finalSummaryStartCol = incumbentStart
    finalSummaryEndCol = lsiEnd
    Call InsertBlockSummaries(wsSummary, COL_BASE, finalSummaryStartCol, finalSummaryEndCol)
    
    '================================================================
    ' ONE GLOBAL SUMMARY INSERTION – NORMALIZED TOTAL
    '================================================================
    Dim totalRow As Long, c As Long
    With wsSummary
        totalRow = .Cells(.Rows.Count, COL_LABELS).End(xlUp).Row + 2
        .Rows(totalRow).Insert xlDown: .Rows(totalRow).ClearFormats
        .Cells(totalRow, COL_LABELS).Value = "Normalized Total"

        For c = finalSummaryStartCol To finalSummaryEndCol Step COLUMNS_PER_BLOCK
            ' — inside your With ws … End With block, and within your c-loop —

        With .Cells(totalRow, c + 2)
            .Formula = "=SUMIF(" & _
                          .Parent.Range( _
                              .Parent.Cells(FIRST_DATA_ROW, "D"), _
                              .Parent.Cells(totalRow - 2, "D") _
                          ).Address(False, False) & _
                          ",""Normalized Sub Total""," & _
                          .Parent.Range( _
                              .Parent.Cells(FIRST_DATA_ROW, c + 2), _
                              .Parent.Cells(totalRow - 2, c + 2) _
                          ).Address(False, False) & _
                       ")"
            .NumberFormat = "$#,##0.00"
            Bordered .Cells
            CondFmt .Cells
        End With
        
        With .Cells(totalRow, c + 3)
            .Formula = "=SUMIF(" & _
                          .Parent.Range( _
                              .Parent.Cells(FIRST_DATA_ROW, "D"), _
                              .Parent.Cells(totalRow - 2, "D") _
                          ).Address(False, False) & _
                          ",""Normalized Sub Total""," & _
                          .Parent.Range( _
                              .Parent.Cells(FIRST_DATA_ROW, c + 3), _
                              .Parent.Cells(totalRow - 2, c + 3) _
                          ).Address(False, False) & _
                       ")"
            .NumberFormat = "$#,##0.00"
            Bordered .Cells
            CondFmt .Cells
        End With

            With .Cells(totalRow, c + 4)
                .Formula = "=" & _
                                .Parent.Cells(totalRow, c + 3).Address(False, False) & _
                                "-" & .Parent.Cells(totalRow, c + 2).Address(False, False)
                .NumberFormat = "$#,##0.00": Bordered .Cells: CondFmt .Cells
            End With
            With .Cells(totalRow, c + 5)
                .Formula = "=" & _
                                .Parent.Cells(totalRow, c + 4).Address(False, False) & _
                                "/" & .Parent.Cells(totalRow, c + 3).Address(False, False)
                .NumberFormat = "0%": Bordered .Cells: CondFmt .Cells
            End With
            .Cells(totalRow, c).ClearContents
            c = c + 1
        Next c
    End With

    '— FINAL FORMATTING —
    With wsSummary
        .Columns.AutoFit
        .Cells.HorizontalAlignment = xlCenter
        .Cells.VerticalAlignment = xlCenter
    End With

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .ErrorCheckingOptions.NumberAsText = True
    End With
End Sub

'====================================================================
'  BLOCK SUMMARY INSERTION –– ONE-ROW VERSION (now clears formats on blanks)
'====================================================================
Private Sub InsertBlockSummaries(ws As Worksheet, _
                                 baseCol As Long, _
                                 firstSupCol As Long, _
                                 lastSupCol As Long)

    Dim r As Long, lastR As Long, startR As Long, endR As Long, c As Long
    Dim totRg As String, baseRg As String, summaryRow As Long

    lastR = ws.Cells(ws.Rows.Count, baseCol).End(xlUp).Row
    r = FIRST_DATA_ROW

    Do While r <= lastR
        If ws.Cells(r, baseCol).Value <> "" Then
            startR = r
            Do While r <= lastR And ws.Cells(r, baseCol).Value <> ""
                r = r + 1
            Loop
            endR = r - 1

            ws.Rows(r).Insert xlDown
            ws.Rows(r).ClearFormats
            lastR = lastR + 1

            ws.Rows(r + 1).Insert xlDown
            ws.Rows(r + 1).ClearFormats
            lastR = lastR + 1

            summaryRow = r + 1
            ws.Cells(summaryRow, COL_LABELS).Value = "Normalized Sub Total"

            For c = firstSupCol To lastSupCol Step COLUMNS_PER_BLOCK
                totRg = ws.Range(ws.Cells(startR, c + 2), ws.Cells(endR, c + 2)).Address(0, 0)
                baseRg = ws.Range(ws.Cells(startR, c + 3), ws.Cells(endR, c + 3)).Address(0, 0)

                With ws.Cells(summaryRow, c + 2)
                    FormulaCell .Cells, "=SUMIF(" & baseRg & ",""<>NA""," & totRg & ")", "$#,##0.00"
                End With
                With ws.Cells(summaryRow, c + 3)
                    FormulaCell .Cells, "=SUMIF(" & totRg & ",""<>NA""," & baseRg & ")", "$#,##0.00"
                End With

                AddSavings ws.Cells(summaryRow, c + 4), ws.Cells(summaryRow, c + 3), ws.Cells(summaryRow, c + 2)
                AddPercent ws.Cells(summaryRow, c + 5), ws.Cells(summaryRow, c + 3), ws.Cells(summaryRow, c + 2)

                ws.Cells(summaryRow, c).ClearContents
                c = c + 1
            Next c

            r = summaryRow + 1
            Do While r <= lastR And ws.Cells(r, baseCol).Value = ""
                r = r + 1
            Loop
        Else
            r = r + 1
        End If
    Loop
End Sub

'====================================================================
'  HELPERS
'====================================================================
Private Sub HeaderCell(c As Range, v As String)
    With c
        .Value = v
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub Bordered(rng As Range)
    rng.Borders.LineStyle = xlContinuous
End Sub

Private Sub CondFmt(rng As Range)
    With rng.FormatConditions: .Delete
        With .Add(xlCellValue, xlEqual, "NA"): .Interior.Color = RGB(217, 217, 217): .Font.Color = RGB(0, 0, 0): End With
        With .Add(xlCellValue, xlGreater, "0"): .Interior.Color = RGB(198, 239, 206): .Font.Color = RGB(0, 97, 0): End With
        With .Add(xlCellValue, xlLess, "0"): .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6): End With
        With .Add(xlCellValue, xlEqual, "0"): .Interior.Color = RGB(255, 235, 156): .Font.Color = RGB(156, 87, 0): End With
    End With
End Sub

Private Sub FormulaCell(tgt As Range, F As String, Optional nf As String = "")
    With tgt
        .Formula = F
        If nf <> "" Then .NumberFormat = nf
        Bordered tgt
        CondFmt tgt
    End With
End Sub

Private Sub AddSavings(tgt As Range, bas As Range, prc As Range)
    With tgt
        .Formula = "=IFERROR(" & bas.Address(False, False) & "-" & prc.Address(False, False) & ",""NA"")"
        .NumberFormat = "$#,##0.00"
        Call Bordered(.Cells): Call CondFmt(tgt)
    End With
End Sub

Private Sub AddPercent(tgt As Range, bas As Range, prc As Range)
    With tgt
        .Formula = "=IFERROR((" & bas.Address(False, False) & "-" & prc.Address(False, False) & ")/" & bas.Address(False, False) & ",""NA"")"
        .NumberFormat = "0%"
        Call Bordered(.Cells): Call CondFmt(tgt)
    End With
End Sub
