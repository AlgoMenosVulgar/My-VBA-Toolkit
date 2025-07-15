Attribute VB_Name = "Module1"
'====================================================================
'  ANALYTICS WITH BASELINE – SUPPLIERS-ONLY
'  • Adds a GRAND-TOTAL section (“Total Normalized Bid / Baseline”)
'    after all block summaries, separated by one blank row.
'  • Saving $ / % for the grand total follow the same rules
'    (NA when either value is NA or both are zero).
'  • Low-% column is left empty in every summary area.
'  • Every variable declared is used exactly once.
'====================================================================
Option Explicit

'–– CONSTANTS (shared by all procedures)
Public Const HDR_ROW         As Long = 1
Public Const FIRST_DATA_ROW As Long = 3
Public Const COL_VOL         As Long = 2        ' Volume  (B)
Public Const COL_BASE        As Long = 3        ' Baseline(C)
Public Const COL_LABELS      As Long = 4        ' label / blank (D)
Public Const FIRST_SUP_COL   As Long = 5        ' first supplier (E)
Public Const SUP_BLOCK_W     As Long = 5        ' 4 data cols + 1 gap

'====================================================================
'  MAIN
'====================================================================
Sub Analytics_With_Baseline_SuppliersOnly_Final_Compact()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.ErrorCheckingOptions.NumberAsText = False
    
    '–– SHEETS
    Dim wsData As Worksheet, wsPrices As Worksheet, wsA As Worksheet
    Set wsData = ActiveWorkbook.Sheets(1)
    Set wsPrices = ActiveWorkbook.Sheets("Prices")
    On Error Resume Next: Set wsA = ActiveWorkbook.Sheets("Analysis"): On Error GoTo 0
    If wsA Is Nothing Then
        Set wsA = wsData.Parent.Sheets.Add(After:=wsData): wsA.Name = "Analysis"
    Else
        wsA.Cells.Clear
    End If
    ' Add this line to set the tab color to purple
    wsA.Tab.Color = RGB(128, 0, 128) ' This is a common RGB for purple
    
    
    '–– SUPPLIER BOUNDS
    Dim supplierStart As Long: supplierStart = 4          ' Prices D
    Dim supplierEnd   As Long
    supplierEnd = wsData.Cells(HDR_ROW, wsData.Columns.Count).End(xlToLeft).Column - 1
    
    Dim lastRowP As Long
    lastRowP = wsPrices.Cells(wsPrices.Rows.Count, 1).End(xlUp).Row
    
    '================================================================
    ' 1. VOLUME & BASELINE
    '================================================================
    HeaderCell wsA.Cells(HDR_ROW + 1, COL_VOL), "Volume"
    HeaderCell wsA.Cells(HDR_ROW + 1, COL_BASE), "Baseline"
    
    Dim r As Long, tgt As Range
    For r = 2 To lastRowP
        ' Volume
        Set tgt = wsA.Cells(r + 1, COL_VOL)
        If wsPrices.Cells(r, "B").Value <> "Blank" And wsPrices.Cells(r, "B").Value <> "" Then
            tgt.Formula = "=Prices!B" & r
            tgt.NumberFormat = "#,##0"
            Bordered tgt
        End If
    
        ' Baseline
        Set tgt = wsA.Cells(r + 1, COL_BASE)
        If wsPrices.Cells(r, "C").Value <> "Blank" And wsPrices.Cells(r, "C").Value <> "" Then
            tgt.Formula = "=IF(OR(" & _
                wsA.Cells(r + 1, COL_VOL).Address(0, 0) & "=""NA""," & _
                "'Prices'!C" & r & "=""NA"")," & _
                """NA""," & _
                wsA.Cells(r + 1, COL_VOL).Address(0, 0) & "*" & _
                "'Prices'!C" & r & _
            ")"
            tgt.NumberFormat = "$#,##0.00"
            Bordered tgt
        End If
    Next r


    
    '================================================================
    ' 2. SUPPLIER DATA
    '================================================================
    Dim supCol As Long: supCol = FIRST_SUP_COL
    Dim j As Long, supName As String, curRow As Long, lowRg As String
    
    For j = supplierStart To supplierEnd
        supName = wsData.Cells(HDR_ROW, j).Value
        
        ' main header
        With wsA.Range(wsA.Cells(HDR_ROW, supCol), wsA.Cells(HDR_ROW, supCol + 3))
            .Merge: .Value = supName: .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(255, 192, 0): .Borders.Weight = xlThin
        End With
        ' sub-headers
        With wsA.Range(wsA.Cells(HDR_ROW + 1, supCol), wsA.Cells(HDR_ROW + 1, supCol + 3))
            .Value = Array("Total Price", "Savings $", "Savings %", "Low %")
            .Interior.Color = RGB(202, 237, 251): .Borders.Weight = xlThin
            .HorizontalAlignment = xlCenter
        End With
        
        curRow = FIRST_DATA_ROW
        For r = 2 To lastRowP
            If wsPrices.Cells(r, j).Value <> "Blank" Then
                ' Total Price
                'FormulaCell wsA.Cells(curRow, supCol), _
                                '"='Prices'!" & wsPrices.Cells(r, j).Address(False, True), "$#,##0.00"
                FormulaCell wsA.Cells(curRow, supCol), _
                                "=IF(OR(" & wsA.Cells(curRow, COL_VOL).Address(0, 0) & "=""NA""," & _
                                "'Prices'!" & wsPrices.Cells(r, j).Address(False, True) & "=""NA"")," & _
                                """NA""," & _
                                wsA.Cells(curRow, COL_VOL).Address(0, 0) & "*" & _
                                "'Prices'!" & wsPrices.Cells(r, j).Address(False, True) & ")", "$#,##0.00"

                ' Savings $
                AddSavings wsA.Cells(curRow, supCol + 1), _
                               wsA.Cells(curRow, COL_BASE), wsA.Cells(curRow, supCol)
                ' Savings %
                AddPercent wsA.Cells(curRow, supCol + 2), _
                               wsA.Cells(curRow, COL_BASE), wsA.Cells(curRow, supCol)
                ' Low %
                lowRg = "'Prices'!" & wsPrices.Range(wsPrices.Cells(r, supplierStart), _
                                                     wsPrices.Cells(r, supplierEnd)).Address(False, False)
                FormulaCell wsA.Cells(curRow, supCol + 3), _
                    "=IF(" & wsA.Cells(curRow, supCol).Address(0, 0) & _
                    "=""NA"",""NA"",IFERROR((SMALL(" & lowRg & ",@FREQUENCY(" & lowRg & ",0)+1)-" & _
                    "'Prices'!" & wsPrices.Cells(r, j).Address(False, True) & ")/" & _
                    "'Prices'!" & wsPrices.Cells(r, j).Address(False, True) & ",""NA""))", "0%"
                CondFmt wsA.Cells(curRow, supCol + 3)
            End If
            curRow = curRow + 1
        Next r
        
        supCol = supCol + SUP_BLOCK_W
    Next j
    
    '================================================================
    ' 3. BLOCK-WISE NORMALIZATION
    '================================================================
    InsertBlockSummaries wsA, COL_BASE, FIRST_SUP_COL, supCol - SUP_BLOCK_W
    
    '================================================================
    ' 4. GRAND-TOTAL SECTION
    '================================================================
    AddGrandTotals wsA, FIRST_SUP_COL, supCol - SUP_BLOCK_W
    
    wsA.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

'====================================================================
'  HELPERS
'====================================================================
Private Sub HeaderCell(tgt As Range, txt As String)
    With tgt
        .Value = txt: .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0): .Borders.Weight = xlThin
    End With
End Sub

Private Sub Bordered(tgt As Range)
    tgt.HorizontalAlignment = xlCenter: tgt.VerticalAlignment = xlCenter
    tgt.Borders.LineStyle = xlContinuous
End Sub

Private Sub FormulaCell(tgt As Range, f As String, Optional nf As String = "")
    tgt.Formula = f
    If nf <> "" Then tgt.NumberFormat = nf
    Bordered tgt
End Sub

'–– Conditional formatting set
Private Sub CondFmt(tgt As Range)
    Dim fc As FormatCondition
    With tgt.FormatConditions
        .Delete
        Set fc = .Add(xlCellValue, xlEqual, "=""NA""")
        fc.Interior.Color = RGB(217, 217, 217): fc.Font.Color = vbBlack
        Set fc = .Add(xlCellValue, xlGreater, "0")
        fc.Interior.Color = RGB(198, 239, 206): fc.Font.Color = RGB(0, 97, 0)
        Set fc = .Add(xlCellValue, xlLess, "0")
        fc.Interior.Color = RGB(255, 199, 206): fc.Font.Color = RGB(156, 0, 6)
        Set fc = .Add(xlCellValue, xlEqual, "0")
        fc.Interior.Color = RGB(255, 235, 156): fc.Font.Color = RGB(156, 87, 0)
    End With
End Sub

Private Sub AddSavings(tgt As Range, baseCell As Range, totCell As Range)
    FormulaCell tgt, _
        "=IF(OR(" & totCell.Address(0, 0) & "=0," & totCell.Address(0, 0) & "=""NA""," & _
        baseCell.Address(0, 0) & "=""NA""),""NA""," & _
        baseCell.Address(0, 0) & "-" & totCell.Address(0, 0) & ")", "$#,##0.00"
    CondFmt tgt
End Sub

Private Sub AddPercent(tgt As Range, baseCell As Range, totCell As Range)
    FormulaCell tgt, _
        "=IF(OR(" & totCell.Address(0, 0) & "=""NA""," & baseCell.Address(0, 0) & "=""NA"",AND(" & _
        totCell.Address(0, 0) & "=0," & baseCell.Address(0, 0) & "=0)),""NA""," & _
        "(" & baseCell.Address(0, 0) & "-" & totCell.Address(0, 0) & ")/" & _
        baseCell.Address(0, 0) & ")", "0%"
    CondFmt tgt
End Sub

'====================================================================
'  BLOCK SUMMARY INSERTION
'====================================================================
Private Sub InsertBlockSummaries(ws As Worksheet, _
                                     baseCol As Long, _
                                     firstSupCol As Long, _
                                     lastSupCol As Long)
    Dim r As Long, lastR As Long, startR As Long, endR As Long, c As Long
    Dim totRg As String, baseRg As String
    
    lastR = ws.Cells(ws.Rows.Count, baseCol).End(xlUp).Row
    r = FIRST_DATA_ROW
    
    Do While r <= lastR
        If ws.Cells(r, baseCol).Value <> "" Then
            startR = r
            Do While ws.Cells(r, baseCol).Value <> "" And r <= lastR
                r = r + 1
            Loop
            endR = r - 1
            
            ' Insert one blank row before "Normalized Bid"
            ws.Rows(r).Insert Shift:=xlDown
            ws.Rows(r).ClearFormats ' Clear all formats from the newly inserted row
            lastR = lastR + 1 ' Adjust lastR as a row was inserted
            
            ' Move r down by 1 to account for the inserted row
            r = r + 1
            
            ws.Cells(r, COL_LABELS).Value = "Normalized Bid"
            ws.Cells(r + 1, COL_LABELS).Value = "Normalized Baseline"
            
            For c = firstSupCol To lastSupCol Step SUP_BLOCK_W
                totRg = ws.Cells(startR, c).Address(0, 0) & ":" & ws.Cells(endR, c).Address(0, 0)
                baseRg = ws.Cells(startR, baseCol).Address(0, 0) & ":" & ws.Cells(endR, baseCol).Address(0, 0)
                
                FormulaCell ws.Cells(r, c), "=SUMIF(" & baseRg & ",""<>NA""," & totRg & ")", "$#,##0.00"
                FormulaCell ws.Cells(r + 1, c), "=SUMIF(" & totRg & ",""<>NA""," & baseRg & ")", "$#,##0.00"
                
                AddSavings ws.Cells(r, c + 1), ws.Cells(r + 1, c), ws.Cells(r, c)
                AddPercent ws.Cells(r, c + 2), ws.Cells(r + 1, c), ws.Cells(r, c)
                
                ws.Cells(r, c + 3).Clear
                ws.Cells(r + 1, c + 3).Clear
            Next c
            
            ' Insert one blank row after "Normalized Baseline"
            ws.Rows(r + 2).Insert Shift:=xlDown
            ws.Rows(r + 2).ClearFormats ' Clear all formats from the newly inserted row
            lastR = lastR + 1 ' Adjust lastR again
            
            r = r + 3 ' Move past the inserted rows and the summary rows
        Else
            r = r + 1
        End If
    Loop
End Sub

'====================================================================
'  GRAND TOTAL (across all blocks)
'====================================================================
Private Sub AddGrandTotals(ws As Worksheet, _
                                     firstSupCol As Long, _
                                     lastSupCol As Long)
    ' find last used row in label column, then leave one blank row
    Dim lastR As Long, startR As Long, lblRg As String, c As Long
    
    lastR = ws.Cells(ws.Rows.Count, COL_LABELS).End(xlUp).Row
    startR = lastR + 2                                      ' blank row separator
    
    ws.Cells(startR, COL_LABELS).Value = "Total Normalized Bid"
    ws.Cells(startR + 1, COL_LABELS).Value = "Total Normalized Baseline"
    
    lblRg = ws.Cells(FIRST_DATA_ROW, COL_LABELS).Address(0, 0) & ":" & _
            ws.Cells(lastR, COL_LABELS).Address(0, 0)
    
    For c = firstSupCol To lastSupCol Step SUP_BLOCK_W
        Dim colRg As String
        colRg = ws.Cells(FIRST_DATA_ROW, c).Address(0, 0) & ":" & ws.Cells(lastR, c).Address(0, 0)
        
        FormulaCell ws.Cells(startR, c), _
            "=SUMIF(" & lblRg & ",""Normalized Bid""," & colRg & ")", "$#,##0.00"
        FormulaCell ws.Cells(startR + 1, c), _
            "=SUMIF(" & lblRg & ",""Normalized Baseline""," & colRg & ")", "$#,##0.00"
        
        AddSavings ws.Cells(startR, c + 1), ws.Cells(startR + 1, c), ws.Cells(startR, c)
        AddPercent ws.Cells(startR, c + 2), ws.Cells(startR + 1, c), ws.Cells(startR, c)
        
        ws.Cells(startR, c + 3).Clear
        ws.Cells(startR + 1, c + 3).Clear
    Next c
End Sub








