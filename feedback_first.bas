Sub CreateFeedbackTab()

    Dim wsFeedback As Worksheet, wsPrices As Worksheet
    Dim endColumn As Long, endRow As Long, lastSupplierCol As Long
    Dim firstSupplierCol As Long: firstSupplierCol = 4   '=== columna F ===
    Dim i As Long, minimumFormula As String
    Dim dropdownValues As String, cell As Range
    Dim subHeaderRange As Range, dropdownRange As Range
    Dim cfRange As Range

    '------------------------------------------------------------------
    ' 1. Comprobar que existe la hoja "Prices"
    '------------------------------------------------------------------
    On Error Resume Next
    Set wsPrices = Worksheets("Prices")
    On Error GoTo 0
    If wsPrices Is Nothing Then
        MsgBox """Prices"" sheet not found!", vbExclamation
        Exit Sub
    End If

    '------------------------------------------------------------------
    ' 2. Localizar la palabra "end"  (marca final)  y delimitar área
    '------------------------------------------------------------------
    With wsPrices
        Dim foundCell As Range
        Set foundCell = .Columns("A").Find(What:="end", LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            MsgBox "'end' not found in column A of 'Prices' tab.", vbExclamation
            Exit Sub
        End If
        endRow = foundCell.Row
        endColumn = .Rows(1).Find(What:="end", LookIn:=xlValues, LookAt:=xlWhole).Column
        lastSupplierCol = endColumn - 1    ' última columna de proveedor
    End With

    '------------------------------------------------------------------
    ' 3. Crear / limpiar hoja "Feedback"
    '------------------------------------------------------------------
    On Error Resume Next
    Set wsFeedback = Worksheets("Feedback")
    On Error GoTo 0
    If wsFeedback Is Nothing Then
        Set wsFeedback = Worksheets.Add
        wsFeedback.Name = "Feedback"
    Else
        wsFeedback.Cells.Clear
    End If

    '------------------------------------------------------------------
    ' 4. Encabezados y formato base
    '------------------------------------------------------------------
    wsFeedback.Range("A1:D1").Value = Array("Vendor", "Percentage Range", "Current Price", "% to Lower")
    Set subHeaderRange = wsFeedback.Range("A1:D1")
    With subHeaderRange
        .Font.Bold = True
        .Interior.Color = RGB(202, 237, 251)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    wsFeedback.Columns("A:D").AutoFit
    wsFeedback.Columns("A").ColumnWidth = WorksheetFunction.Max( _
        wsFeedback.Columns("A").ColumnWidth, Len("Landsberg Orora") * 1.1)

    '------------------------------------------------------------------
    ' 5. Lista desplegable de proveedores (fila 1, columnas F … lastSupplierCol)
    '------------------------------------------------------------------
    Set dropdownRange = wsPrices.Range(wsPrices.Cells(1, firstSupplierCol), wsPrices.Cells(1, lastSupplierCol))
    dropdownValues = ""
    For Each cell In dropdownRange
        dropdownValues = dropdownValues & cell.Value & ","
    Next cell
    If Len(dropdownValues) > 0 Then dropdownValues = Left(dropdownValues, Len(dropdownValues) - 1)
    With wsFeedback.Range("A2")
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                        Operator:=xlBetween, Formula1:=dropdownValues
        .HorizontalAlignment = xlCenter
    End With

    '------------------------------------------------------------------
    ' 6. Formato de la celda B2  (rangos en % con signo +)
    '------------------------------------------------------------------
    With wsFeedback.Range("B2")
        .Font.Color = RGB(255, 0, 0)
        .NumberFormat = "+0;-0"
        .HorizontalAlignment = xlCenter
    End With

    '------------------------------------------------------------------
    ' 7. Fórmulas dinámicas (desde la fila 2 hasta fila anterior a “end”)
    '------------------------------------------------------------------
    For i = 2 To endRow - 1
        ' ?————— comprobamos si TODA la fila es “Blank”
        If WorksheetFunction.CountIf( _
             wsPrices.Range(wsPrices.Cells(i, firstSupplierCol), _
                            wsPrices.Cells(i, lastSupplierCol)), _
             "Blank" _
           ) < (lastSupplierCol - firstSupplierCol + 1) Then

            '----- Current Price -------------------------------------------------
            wsFeedback.Cells(i, 3).Formula = _
                "=INDEX(Prices!" & _
                wsPrices.Cells(i, firstSupplierCol).Address(False, True) & ":" & _
                wsPrices.Cells(i, lastSupplierCol).Address(False, True) & _
                ", 1, MATCH(A2, Prices!" & _
                wsPrices.Cells(1, firstSupplierCol).Address(False, True) & ":" & _
                wsPrices.Cells(1, lastSupplierCol).Address(False, True) & ", 0))"
            wsFeedback.Cells(i, 3).HorizontalAlignment = xlCenter

            '----- % to Lower ----------------------------------------------------
            ' Build the MIN() once
            minimumFormula = "MIN(Prices!" & _
                             wsPrices.Range(wsPrices.Cells(i, firstSupplierCol), _
                                            wsPrices.Cells(i, lastSupplierCol)).Address(False, False) & ")"
            
            ' % difference expression reused
            Dim pctExpr As String
            pctExpr = "ABS(100*(" & minimumFormula & "-C" & i & ")/C" & i & ")"
            
            ' New formula with the 90–95% cap and +5% band
            wsFeedback.Cells(i, 4).Formula = _
                "=IF(OR(" & minimumFormula & "=0, C" & i & "=""" & "NA" & """), ""NA"", " & _
                "IF(ISNUMBER((" & minimumFormula & "-C" & i & ")/C" & i & "), " & _
                    "IF((C" & i & "-" & minimumFormula & ")<=0, ""Good"", " & _
                        "IF(" & pctExpr & ">90, ""90 - 95%"", " & _
                            "CEILING(" & pctExpr & ",5) & "" - "" & (CEILING(" & pctExpr & ",5)+5) & ""%"" " & _
                        ")" & _
                    "), " & _
                    "(" & minimumFormula & "-C" & i & ")/C" & i & _
                "))"
            
            wsFeedback.Cells(i, 4).HorizontalAlignment = xlCenter


        End If
    Next i

    '------------------------------------------------------------------
    ' 8. Formato final de columnas C y D (solo filas con fórmulas)
    '------------------------------------------------------------------
    For i = 2 To endRow - 1
        If wsFeedback.Cells(i, 3).HasFormula Then
            With wsFeedback.Range("C" & i)
                .NumberFormat = "$#,##0.00"
                .HorizontalAlignment = xlCenter
            End With
            wsFeedback.Columns("D").AutoFit
            With wsFeedback.Range("C" & i & ":D" & i)
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
            End With
        End If
    Next i

    '------------------------------------------------------------------
    ' 9. Formato condicional sobre columna D (solo filas con fórmulas)
    '------------------------------------------------------------------
    ' Primero borramos cualquier regla previa en D2:D(endRow-1)
    wsFeedback.Range("D2:D" & endRow - 1).FormatConditions.Delete
    ' Ahora, agregamos reglas fila por fila
    For i = 2 To endRow - 1
        If wsFeedback.Cells(i, 3).HasFormula Then
            Set cfRange = wsFeedback.Range("D" & i)
            cfRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NA"""
            cfRange.FormatConditions(1).Interior.Color = RGB(217, 217, 217)
            cfRange.FormatConditions(1).Font.Color = RGB(0, 0, 0)
            cfRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Good"""
            cfRange.FormatConditions(2).Interior.Color = RGB(198, 239, 206)
            cfRange.FormatConditions(2).Font.Color = RGB(0, 97, 0)
            cfRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, Formula1:="=""Good"""
            cfRange.FormatConditions(3).Interior.Color = RGB(255, 199, 206)
            cfRange.FormatConditions(3).Font.Color = RGB(156, 0, 6)
        End If
    Next i

    '------------------------------------------------------------------
    ' 10. Limpiar formatos de filas "Blank" (ni bordes ni formato condicional)
    '------------------------------------------------------------------
    For i = 2 To endRow - 1
        If Not wsFeedback.Cells(i, 3).HasFormula Then
            With wsFeedback.Range("A" & i & ":D" & i)
                .ClearFormats
                .FormatConditions.Delete
            End With
        End If
    Next i

End Sub







