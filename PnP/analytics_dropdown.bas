Sub Analytics_With_Baseline()
    
    Application.ScreenUpdating = False 'This part prevent the screen updating and makes it less probable it "crashes" due to lag or slow computers
    Application.DisplayAlerts = False
    
    Application.ErrorCheckingOptions.NumberAsText = False 'Testing errors

    Dim wsData As Worksheet
    Dim wsPrices As Worksheet
    Dim wsAnalysis As Worksheet
    Dim lastColumn As Long
    Dim supplierStart As Long
    Dim supplierEnd As Long
    Dim headerRow As Long
    Dim colOffset As Long
    Dim supplierCount As Long
    Dim i As Long, j As Long
    Dim currentCol As Long
    Dim currentRow As Long
    Dim supplierName As String
    Dim supplierHeaderRange As Range
    Dim subHeaderRange As Range
    Dim formulaCell As Range
    Dim formulaString As String
    Dim formulaLowPercent As String
    Dim supplierRangeAddress As String
    Dim lastRow As Long
    Dim exitCondition As Boolean

    ' Set the source data and Prices worksheet
    Set wsData = ActiveWorkbook.Sheets(1)
    Set wsPrices = ActiveWorkbook.Sheets("Prices")

    ' Add or clear Analysis worksheet
    On Error Resume Next
    Set wsAnalysis = ActiveWorkbook.Sheets("Analysis")
    On Error GoTo 0
    If wsAnalysis Is Nothing Then
        Set wsAnalysis = ActiveWorkbook.Sheets.Add
        wsAnalysis.Name = "Analysis"
    Else
        wsAnalysis.Cells.Clear
    End If
    

    ' Define header row
    headerRow = 1
    currentRow = 3

    ' Determine the range of suppliers
    supplierStart = 5
    lastColumn = wsData.Cells(headerRow, wsData.Columns.Count).End(xlToLeft).Column
    supplierEnd = lastColumn - 1
    lastRow = wsPrices.Cells(wsPrices.Rows.Count, 1).End(xlUp).Row

    ' Calculate the number of suppliers
    supplierCount = supplierEnd - supplierStart + 1
    
    
    '------
    ' Merge and format header cell for "Volume"
    currentCol = 1
    
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol))
        .Merge
        .Value = "Volume"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
    
    
    Dim rowNum As Long, srcVal As Variant
    
    For rowNum = 3 To lastRow
        srcVal = wsPrices.Cells(rowNum - 1, "C").Value        'read once, faster than repeated sheet calls volumn column
        
        With wsAnalysis.Cells(rowNum, currentCol)
            If srcVal = "Blank" Then
                .Clear                                         'remove content + all formatting
            Else
                .Formula = "=Prices!C" & (rowNum - 1)          'keep live link for every other case
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "#,##0"
            End If
        End With
    Next rowNum

    
    '------

    ' Merge and format header cell
    currentCol = 2
    BaselineColumn = 2
    
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol))
        .Merge
        .Value = "Baseline"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
    
    ' Insert values from Prices!G while clearing any cell whose source says "Blank"
    For rowNum = 3 To lastRow
        srcVal = wsPrices.Cells(rowNum - 1, "D").Value         'read once, avoid repeated sheet calls
        
        With wsAnalysis.Cells(rowNum, currentCol)
            If srcVal = "Blank" Then
                .Clear                                          'remove content and all formatting
            Else
                .Formula = "=Prices!D" & (rowNum - 1)           'keep live link for every other case
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End If
        End With
    Next rowNum
    '------
        

    ' Generate headers for each supplier in the Analysis sheet
    currentCol = 3
    
    Set UniqueValues = CreateObject("Scripting.Dictionary")
    UniqueValues.Add "All", 1 ' Add "All" as the first item
    
    ' Collect unique values from C2:lastRow
    For Each cell In wsPrices.Range("B2:B" & lastRow)
        If Not UniqueValues.exists(cell.Value) Then
            UniqueValues.Add cell.Value, 1
        End If
    Next cell
    
    ' Convert dictionary keys to an array
    ReDim arr(1 To UniqueValues.Count)
    i = 1
    For Each Key In UniqueValues.keys
        arr(i) = Key
        i = i + 1
    Next Key
    
    ' Merge and format header cell
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol))
        .Merge
        .Value = "CATEGORIZED"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
    
    ' Define and format dropdown cells
    Dim dropdownRange As Range
    ' Set dropdownRange to include all three cells
    Set dropdownRange = wsAnalysis.Range(wsAnalysis.Cells(headerRow + 2, currentCol), wsAnalysis.Cells(headerRow + 4, currentCol))
    
    With dropdownRange
        ' Apply formatting to all three cells
        .Interior.Color = RGB(202, 237, 251)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Add Data Validation to all three cells
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:=Join(arr, ",")
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = False
            .ShowError = False
        End With
    End With
    
    currentCol = 5
    For i = supplierStart To supplierEnd
        ' Get the supplier name
        supplierName = wsData.Cells(headerRow, i).Value

        ' Write the supplier name and merge cells across four columns
        Set supplierHeaderRange = wsAnalysis.Range(wsAnalysis.Cells(headerRow, currentCol), wsAnalysis.Cells(headerRow, currentCol + 3))
        supplierHeaderRange.Merge
        supplierHeaderRange.Value = supplierName
        supplierHeaderRange.HorizontalAlignment = xlCenter
        supplierHeaderRange.VerticalAlignment = xlCenter
        supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
        supplierHeaderRange.Borders.Weight = xlThin

        ' Write sub-headers
        Set subHeaderRange = wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol + 3))
        subHeaderRange.Cells(1, 1).Value = "Total Price"
        subHeaderRange.Cells(1, 2).Value = "Savings $"
        subHeaderRange.Cells(1, 3).Value = "Savings %"
        subHeaderRange.Cells(1, 4).Value = "Low %"
        subHeaderRange.Interior.Color = RGB(202, 237, 251)
        subHeaderRange.Borders.Weight = xlThin
        subHeaderRange.HorizontalAlignment = xlCenter
        subHeaderRange.VerticalAlignment = xlCenter

        ' Iterate over rows in Prices tab for calculations
        For j = 2 To lastRow
            ' Check for exit condition
            exitCondition = wsPrices.Cells(j, 1).Value = "end"
            If exitCondition Then Exit For

            ' Check if the cell contains "Blank" in the Total Price (I) or Savings % (J) column
            If wsPrices.Cells(j, i).Value = "Blank" Or wsPrices.Cells(j, i + 1).Value = "Blank" Then
                ' Leave the cell blank and continue to the next iteration
                wsAnalysis.Cells(currentRow, currentCol).ClearContents
                wsAnalysis.Cells(currentRow, currentCol + 1).ClearContents
                wsAnalysis.Cells(currentRow, currentCol + 2).ClearContents
                wsAnalysis.Cells(currentRow, currentCol + 3).ClearContents
            Else
                ' Total Price
                Set formulaCell = wsAnalysis.Cells(currentRow, currentCol)
                'formulaString = "='Prices'!" & wsPrices.Cells(j, i).Address(False, True) 'test filter
                
                'formulaString = "=IF(OR(" & dropdownRange.Address(False, False) & "=""All""," & dropdownRange.Address(False, False) & "=Prices!C" & j & "), 'Prices'!" & wsPrices.Cells(j, i).Address(False, True) & ", ""NA"")"
                formulaString = _
                "=IF(OR(" & _
                    "AND(" & _
                        "COUNTIF(" & dropdownRange.Address(False, False) & ",""All"")=0," & _
                        "COUNTIF(" & dropdownRange.Address(False, False) & ",Prices!B" & j & ")=0" & _
                    ")," & _
                    "'Prices'!C" & j & "=""NA""," & _
                    "'Prices'!" & wsPrices.Cells(j, i).Address(False, True) & "=""NA""" & _
                ")," & _
                """NA""," & _
                "'Prices'!C" & j & "*'Prices'!" & wsPrices.Cells(j, i).Address(False, True) & _
                ")"




                With formulaCell
                    .Formula = formulaString
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
                End With
                
                ' Savings $ (same as before)
                Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 1)
                ' Savings $ (column to the right of Total Price)
                
               formulaString = _
                "=IF(OR(" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 1).Address(False, True) & "=""NA""," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "=""NA""," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=""NA""," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 1).Address(False, True) & "=0," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "=0," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=0" & _
                "),""NA"",(" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "*" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 1).Address(False, True) & _
                ")-" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & _
                ")"

                

                With formulaCell
                    .Formula = formulaString
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
                    
                    ' Apply conditional formatting based on value
                    Dim formatConditions As FormatCondition
                    With .formatConditions
                        .Delete ' Clear existing conditions
                        
                        ' Format for values greater than 0
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                        formatConditions.Interior.Color = RGB(217, 217, 217) ' Light gray
                        formatConditions.Font.Color = RGB(0, 0, 0)         ' Black
                
                        ' Format for values greater than 0
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                        formatConditions.Interior.Color = RGB(198, 239, 206) ' Light green
                        formatConditions.Font.Color = RGB(0, 97, 0)         ' Dark green
                
                        ' Format for values less than 0
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                        formatConditions.Interior.Color = RGB(255, 199, 206) ' Light red
                        formatConditions.Font.Color = RGB(156, 0, 6)         ' Dark red
                
                        ' Format for values equal to 0
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                        formatConditions.Interior.Color = RGB(255, 235, 156) ' Yellow
                        formatConditions.Font.Color = RGB(156, 87, 0)        ' Dark yellow
                    End With
                            
                End With

                ' Savings % (Percentage Format)
                Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 2)
                formulaString = _
                "=IF(OR(" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 1).Address(False, True) & "=""NA""," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "=""NA""," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=""NA""," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 1).Address(False, True) & "=0," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "=0," & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=0" & _
                "),""NA"",(" & _
                    "(Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "*" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 1).Address(False, True) & ")" & _
                    "-" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & _
                ")/(" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "*" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 1).Address(False, True) & _
                "))"

                
                With formulaCell
                    .Formula = formulaString
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .NumberFormat = "0%" ' Formatting as percentage
                
                    With .formatConditions
                        .Delete ' Clear existing conditions
                
                        ' NA
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NA""")
                        formatConditions.Interior.Color = RGB(217, 217, 217)
                        formatConditions.Font.Color = RGB(0, 0, 0)
                
                        ' > 0
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                        formatConditions.Interior.Color = RGB(198, 239, 206)
                        formatConditions.Font.Color = RGB(0, 97, 0)
                
                        ' < 0
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                        formatConditions.Interior.Color = RGB(255, 199, 206)
                        formatConditions.Font.Color = RGB(156, 0, 6)
                
                        ' = 0
                        Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                        formatConditions.Interior.Color = RGB(255, 235, 156)
                        formatConditions.Font.Color = RGB(156, 87, 0)
                    End With
                End With


                ' Low % (Percentage Format)
                Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 3)
                
                ' Get supplier range address from Prices sheet
                supplierRangeAddress = "'Prices'!" & wsPrices.Range(wsPrices.Cells(j, supplierStart), _
                    wsPrices.Cells(j, supplierEnd)).Address(False, False)
                
                ' Construct the formula with better readability
                formulaLowPercent = "=IF(" & _
                    "Analysis!" & wsAnalysis.Cells(currentRow, currentCol).Address(False, False) & "=""NA"",""NA""," & _
                    "IF('Prices'!" & wsPrices.Cells(j, i).Address(False, True) & "=""NA"",""NA""," & _
                    "IFERROR((SMALL(" & supplierRangeAddress & "," & _
                    "@FREQUENCY(" & supplierRangeAddress & ",0)+" & _
                    wsAnalysis.Cells(1, (lastColumn - 5) * 5 + 11).Address(False, False) & ")-" & _
                    "'Prices'!" & wsPrices.Cells(j, i).Address(False, True) & ")/" & _
                    "'Prices'!" & wsPrices.Cells(j, i).Address(False, True) & ",""NA"")))"
                
                With formulaCell
                    .Formula = formulaLowPercent
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .NumberFormat = "0%"    ' Format as percentage
                
                    With .formatConditions
                        .Delete    ' Clear existing conditions
                
                        ' NA values - Light gray background, black text
                        Set formatConditions = .Add(Type:=xlCellValue, _
                                                  Operator:=xlEqual, _
                                                  Formula1:="=""NA""")
                        With formatConditions
                            .Interior.Color = RGB(217, 217, 217)
                            .Font.Color = RGB(0, 0, 0)
                        End With
                
                        ' Values > 0 - Light green background, dark green text
                        Set formatConditions = .Add(Type:=xlCellValue, _
                                                  Operator:=xlGreater, _
                                                  Formula1:="0")
                        With formatConditions
                            .Interior.Color = RGB(198, 239, 206)
                            .Font.Color = RGB(0, 97, 0)
                        End With
                
                        ' Values < 0 - Light red background, dark red text
                        Set formatConditions = .Add(Type:=xlCellValue, _
                                                  Operator:=xlLess, _
                                                  Formula1:="0")
                        With formatConditions
                            .Interior.Color = RGB(255, 199, 206)
                            .Font.Color = RGB(156, 0, 6)
                        End With
                
                        ' Values = 0 - Yellow background, dark yellow text
                        Set formatConditions = .Add(Type:=xlCellValue, _
                                                  Operator:=xlEqual, _
                                                  Formula1:="0")
                        With formatConditions
                            .Interior.Color = RGB(255, 235, 156)
                            .Font.Color = RGB(156, 87, 0)
                        End With
                    End With
                End With
                                
            End If

            ' Move to the next row in the Analysis sheet
            currentRow = currentRow + 1
        Next j
        
        currentRow = currentRow + 1 'just leaving one space btw the final item and the baseline header
       
        ' Write the supplier name and merge cells across four columns
        Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol)
        supplierHeaderRange.Value = "Normalized Total"
        supplierHeaderRange.HorizontalAlignment = xlCenter
        supplierHeaderRange.VerticalAlignment = xlCenter
        supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
        supplierHeaderRange.Borders.Weight = xlThin
        
        
        'wsAnalysis.Cells(currentRow + 1, currentCol).Formula = "=SUMIF('Prices'!" & wsPrices.Columns(supplierStart - 1).Address(False, True) & ", ""<>NA"", 'Prices'!" & wsPrices.Columns(i).Address(False, True) & ")"
        wsAnalysis.Cells(currentRow + 1, currentCol).Formula = _
        "=SUMPRODUCT(" & _
            "ISNUMBER(" & _
                wsAnalysis.Range(wsAnalysis.Cells(3, BaselineColumn), _
                                 wsAnalysis.Cells(lastRow, BaselineColumn)).Address(False, False) & _
            ")*ISNUMBER(" & _
                wsAnalysis.Range(wsAnalysis.Cells(3, currentCol), _
                                 wsAnalysis.Cells(lastRow, currentCol)).Address(False, False) & _
            ")*" & _
                wsAnalysis.Range(wsAnalysis.Cells(3, currentCol), _
                                 wsAnalysis.Cells(lastRow, currentCol)).Address(False, False) & _
        ")"




        
        Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol)
        With formulaCell
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormat = "$#,##0.00"
        End With
        
        
        
        'wsAnalysis.Cells(currentRow + 2, currentCol).Formula = "=SUMIF('Prices'!" & wsPrices.Columns(i).Address(False, True) & ", ""<>NA"", 'Prices'!" & wsPrices.Columns(supplierStart - 1).Address(False, True) & ")"
        wsAnalysis.Cells(currentRow + 2, currentCol).Formula = "=SUMIF(" & wsAnalysis.Cells(3, currentCol).Address(False, False) & ":" & wsAnalysis.Cells(lastRow, currentCol).Address(False, False) & ",""<>NA""," & wsAnalysis.Cells(3, BaselineColumn).Address(False, False) & ":" & wsAnalysis.Cells(lastRow, BaselineColumn).Address(False, False) & ")"

        
        Set formulaCell = wsAnalysis.Cells(currentRow + 2, currentCol)
        With formulaCell
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
        End With

        
       
        ' Write the supplier name and merge cells across four columns
        Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 1)
        supplierHeaderRange.Value = "Saving $"
        supplierHeaderRange.HorizontalAlignment = xlCenter
        supplierHeaderRange.VerticalAlignment = xlCenter
        supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
        supplierHeaderRange.Borders.Weight = xlThin
        
               
        ' Apply the formula to the target cell
        
        wsAnalysis.Cells(currentRow + 1, currentCol + 1).Formula = "=IF(OR(" & _
        wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=0," & _
        wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=""NA""," & _
        wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=0," & _
        wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=""NA""), " & _
        """NA"", " & _
        wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "-" & _
        wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & ")"

        
        ' Set a reference to the formula cell
        Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 1)
        
        ' Format the cell
        With formulaCell
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormat = "$#,##0.00" ' Formatting as currency
            
            With .formatConditions
                 .Delete ' Clear existing conditions
                
                 ' Format for values greater than 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                 formatConditions.Interior.Color = RGB(217, 217, 217) ' Light gray
                 formatConditions.Font.Color = RGB(0, 0, 0)         ' Black
        
                 ' Format for values greater than 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                 formatConditions.Interior.Color = RGB(198, 239, 206) ' Light green
                 formatConditions.Font.Color = RGB(0, 97, 0)         ' Dark green
        
                 ' Format for values less than 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                 formatConditions.Interior.Color = RGB(255, 199, 206) ' Light red
                 formatConditions.Font.Color = RGB(156, 0, 6)         ' Dark red
        
                 ' Format for values equal to 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                 formatConditions.Interior.Color = RGB(255, 235, 156) ' Yellow
                 formatConditions.Font.Color = RGB(156, 87, 0)        ' Dark yellow
                
             End With
        End With




        ' Write the supplier name and merge cells across four columns
        Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 2)
        supplierHeaderRange.Value = "Saving %"
        supplierHeaderRange.HorizontalAlignment = xlCenter
        supplierHeaderRange.VerticalAlignment = xlCenter
        supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
        supplierHeaderRange.Borders.Weight = xlThin
        
        wsAnalysis.Cells(currentRow + 1, currentCol + 2).Formula = "=IF(OR(" & _
        wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=0," & _
        wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=""NA""," & _
        wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=0," & _
        wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=""NA""), " & _
        """Check Values"", (" & _
        wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "-" & _
        wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & ")/" & _
        wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & ")"
        
        Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 2)
        
        ' Format the cell
        With formulaCell
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormat = "0%" ' Formatting as percentage
            
            With .formatConditions
                 .Delete ' Clear existing conditions
                
                 ' Format for values greater than 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                 formatConditions.Interior.Color = RGB(217, 217, 217) ' Light gray
                 formatConditions.Font.Color = RGB(0, 0, 0)         ' Black
        
                 ' Format for values greater than 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                 formatConditions.Interior.Color = RGB(198, 239, 206) ' Light green
                 formatConditions.Font.Color = RGB(0, 97, 0)         ' Dark green
        
                 ' Format for values less than 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                 formatConditions.Interior.Color = RGB(255, 199, 206) ' Light red
                 formatConditions.Font.Color = RGB(156, 0, 6)         ' Dark red
        
                 ' Format for values equal to 0
                 Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                 formatConditions.Interior.Color = RGB(255, 235, 156) ' Yellow
                 formatConditions.Font.Color = RGB(156, 87, 0)        ' Dark yellow
                
             End With
        End With
              

        ' Move to the next set of columns, leaving one blank
        currentCol = currentCol + 5
        currentRow = 3 ' Reset the row pointer for the next supplier
    Next i
    
   
    '-------- Incumbentt
    ' Merge and set header for "Incumbent Solution"
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow, currentCol), wsAnalysis.Cells(headerRow, currentCol + 5))
        .Merge
        .Value = "Incumbent Solution"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
    
    ' Add sub-headers for "Incumbent Solution"
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol + 5))
        .Value = Array("Supplier", "Unit Price", "Total Price", "Baseline", "Savings $", "Savings %")
        .Interior.Color = RGB(202, 237, 251)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Initialize current row for data entry
    currentRow = headerRow + 2
    
    ' Loop through rows to populate data
    For j = 2 To lastRow
        If wsPrices.Cells(j, 1).Value = "end" Then Exit For
    
        ' Skip rows where the supplier value is "Blank"
        If wsPrices.Cells(j, supplierStart).Value <> "Blank" Then
    
            ' ========================
            ' Supplier (text)
            ' ========================
            wsAnalysis.Cells(currentRow, currentCol).Formula = _
                "=Prices!" & wsPrices.Cells(j, 1).Address(False, False)
    
            With wsAnalysis.Cells(currentRow, currentCol)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ========================
            ' Unit Price (INDEX logic)
            ' ========================
           wsAnalysis.Cells(currentRow, currentCol + 1).Formula = _
            "=IF(OR(" & _
                "COUNTIF(" & dropdownRange.Address(False, False) & ",""All"")>0," & _
                "COUNTIF(" & dropdownRange.Address(False, False) & ",Prices!B" & j & ")>0)," & _
                "IFERROR(" & _
                    "INDEX(Prices!" & wsPrices.Cells(j, supplierStart - 1).Address(False, False) & ":" & _
                                wsPrices.Cells(j, supplierEnd).Address(False, False) & "," & _
                           "MATCH(" & wsAnalysis.Cells(currentRow, currentCol).Address(False, False) & "," & _
                                    "Prices!" & wsPrices.Cells(1, supplierStart - 1).Address(False, False) & ":" & _
                                              wsPrices.Cells(1, supplierEnd).Address(False, False) & ",0)" & _
                    ")," & """NA""" & _
                ")," & _
                """NA""" & _
            ")"

            With wsAnalysis.Cells(currentRow, currentCol + 1)
                .NumberFormat = "$#,##0.00"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ========================
            ' Total Price = Volume * Unit Price
            ' Volume is column 1 on Analysis
            ' ========================
            wsAnalysis.Cells(currentRow, currentCol + 2).Formula = _
                "=IF(OR(" & _
                    wsAnalysis.Cells(currentRow, 1).Address(False, False) & "=""NA""," & _
                    wsAnalysis.Cells(currentRow, 1).Address(False, False) & "=0," & _
                    wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & "=""NA"")," & _
                """NA""," & _
                    wsAnalysis.Cells(currentRow, 1).Address(False, False) & "*" & _
                    wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & ")"
    
            With wsAnalysis.Cells(currentRow, currentCol + 2)
                .NumberFormat = "$#,##0.00"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ========================
            ' Baseline total = BaselineUnit * Volume
            ' (same dropdown gating as before)
            ' ========================
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 3)
            formulaCell.Formula = _
                "=IF(OR(" & _
                    "COUNTIF(" & dropdownRange.Address(False, False) & ",""All"")>0," & _
                    "COUNTIF(" & dropdownRange.Address(False, False) & ",Prices!B" & j & ")>0)," & _
                    "IF(OR(" & _
                        "Prices!" & wsPrices.Cells(j, supplierStart - 1).Address(False, False) & "=0," & _
                        wsAnalysis.Cells(currentRow, 1).Address(False, False) & "=0," & _
                        wsAnalysis.Cells(currentRow, 1).Address(False, False) & "=""NA"")," & _
                    """NA""," & _
                        "Prices!" & wsPrices.Cells(j, supplierStart - 1).Address(False, False) & "*" & _
                        wsAnalysis.Cells(currentRow, 1).Address(False, False) & _
                    ")," & _
                """NA"")"
    
            With formulaCell
                .NumberFormat = "$#,##0.00"
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ========================
            ' Savings $ = Baseline - Total Price
            ' ========================
            wsAnalysis.Cells(currentRow, currentCol + 4).Formula = _
                "=IF(OR(" & _
                    wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=""NA""," & _
                    wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=0," & _
                    wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "=""NA"")," & _
                """NA""," & _
                    wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "-" & _
                    wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & ")"
    
            With wsAnalysis.Cells(currentRow, currentCol + 4)
                .NumberFormat = "$#,##0.00"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
    
                With .formatConditions
                    .Delete
    
                    ' "NA"
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NA""")
                    formatConditions.Interior.Color = RGB(217, 217, 217)
                    formatConditions.Font.Color = RGB(0, 0, 0)
    
                    ' > 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206)
                    formatConditions.Font.Color = RGB(0, 97, 0)
    
                    ' < 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206)
                    formatConditions.Font.Color = RGB(156, 0, 6)
    
                    ' = 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156)
                    formatConditions.Font.Color = RGB(156, 87, 0)
                End With
            End With
    
            ' ========================
            ' Savings % = Savings $ / Baseline
            ' ========================
            wsAnalysis.Cells(currentRow, currentCol + 5).Formula = _
                "=IF(OR(" & _
                    wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=0," & _
                    wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=""NA""," & _
                    wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "=""NA"")," & _
                """NA"",(" & _
                    wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "-" & _
                    wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & ")/" & _
                    wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & ")"
    
            With wsAnalysis.Cells(currentRow, currentCol + 5)
                .NumberFormat = "0%"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
    
                With .formatConditions
                    .Delete
    
                    ' "NA"
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NA""")
                    formatConditions.Interior.Color = RGB(217, 217, 217)
                    formatConditions.Font.Color = RGB(0, 0, 0)
    
                    ' > 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206)
                    formatConditions.Font.Color = RGB(0, 97, 0)
    
                    ' < 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206)
                    formatConditions.Font.Color = RGB(156, 0, 6)
    
                    ' = 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156)
                    formatConditions.Font.Color = RGB(156, 87, 0)
                End With
            End With
    
        End If
    
        ' Move to the next row
        currentRow = currentRow + 1
    Next j
    
    currentRow = currentRow + 1 ' just leaving one space btw the final item and the baseline header
    
    ' "Normalized Total" footer for Incumbent
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol)
    supplierHeaderRange.Value = "Normalized Total"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    ' Sum Total Price where Baseline <> NA
    wsAnalysis.Cells(currentRow + 1, currentCol).Formula2 = _
        "=SUMIF(" & _
            "Analysis!" & wsAnalysis.Range(wsAnalysis.Cells(1, currentCol + 3), wsAnalysis.Cells(lastRow, currentCol + 3)).Address(False, False) & _
            ",""<>NA""," & _
            "Analysis!" & wsAnalysis.Range(wsAnalysis.Cells(1, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & _
        ")"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00"
    End With
    
    ' Sum Baseline where Total Price <> NA  (BaselineUnit * Volume)
    wsAnalysis.Cells(currentRow + 2, currentCol).Formula2 = _
    "=SUMPRODUCT((" & _
        "Analysis!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & _
        "<>""NA"")*" & _
        "Analysis!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 3), wsAnalysis.Cells(lastRow, currentCol + 3)).Address(False, False) & _
    ")"

    
    Set formulaCell = wsAnalysis.Cells(currentRow + 2, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00"
    End With

    
    ' "Saving $" header
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 1)
    supplierHeaderRange.Value = "Saving $"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    ' Saving $ (footer)
    wsAnalysis.Cells(currentRow + 1, currentCol + 1).Formula = _
        "=IF(OR(" & _
            wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=0," & _
            wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=""NA""," & _
            wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=0," & _
            wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=""NA"")," & _
        """NA""," & _
            wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "-" & _
            wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & ")"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 1)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00"
    
        With .formatConditions
            .Delete
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NA""")
            formatConditions.Interior.Color = RGB(217, 217, 217)
            formatConditions.Font.Color = RGB(0, 0, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            formatConditions.Interior.Color = RGB(198, 239, 206)
            formatConditions.Font.Color = RGB(0, 97, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 199, 206)
            formatConditions.Font.Color = RGB(156, 0, 6)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 235, 156)
            formatConditions.Font.Color = RGB(156, 87, 0)
        End With
    End With
    
    ' "Saving %" header
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 2)
    supplierHeaderRange.Value = "Saving %"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    ' Saving % (footer)
    wsAnalysis.Cells(currentRow + 1, currentCol + 2).Formula = _
        "=IF(OR(" & _
            wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=0," & _
            wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=""NA""," & _
            wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=0," & _
            wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=""NA"")," & _
        """Check Values"",(" & _
            wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "-" & _
            wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & ")/" & _
            wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & ")"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 2)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0%"
    
        With .formatConditions
            .Delete
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NA""")
            formatConditions.Interior.Color = RGB(217, 217, 217)
            formatConditions.Font.Color = RGB(0, 0, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            formatConditions.Interior.Color = RGB(198, 239, 206)
            formatConditions.Font.Color = RGB(0, 97, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 199, 206)
            formatConditions.Font.Color = RGB(156, 0, 6)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 235, 156)
            formatConditions.Font.Color = RGB(156, 87, 0)
        End With
    End With
    
    ' Move to the next set of columns
    currentCol = currentCol + 7

    
    
    '-------- Incumbentt
   
   

   ' ------------------------------------------------ Lowestt
   
    ' -------------------------------------------------
    ' LOWEST SECTION (updated with Unit Price + Total)
    ' -------------------------------------------------
    
    ' Add "Lowest" header (covers 5 data columns to the right of the dropdown)
    ' Layout:
    '   [currentCol]       = dropdown (1..5)
    '   [currentCol+1..+5] = merged "Lowest"
    Set supplierHeaderRange = wsAnalysis.Range( _
        wsAnalysis.Cells(headerRow, currentCol + 1), _
        wsAnalysis.Cells(headerRow, currentCol + 5) _
    )
    
    With supplierHeaderRange
        .Merge
        .Value = "Lowest"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
    
    ' Dropdown (rank 1–5) in the cell to the left of "Lowest"
    Set dropdownRange = wsAnalysis.Range( _
        wsAnalysis.Cells(headerRow, currentCol), _
        wsAnalysis.Cells(headerRow, currentCol) _
    )
    
    With dropdownRange
        With .Validation
            .Delete
            .Add Type:=xlValidateList, _
                 AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, _
                 Formula1:="1,2,3,4,5"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Interior.Color = RGB(202, 237, 251)
    End With
    
    ' Sub-headers for "Lowest"
    '   currentCol      = Supplier
    '   currentCol + 1  = Unit Price   (k-th lowest unit)
    '   currentCol + 2  = Total Price  (Unit * Volume)
    '   currentCol + 3  = Baseline     (Baseline unit * Volume)
    '   currentCol + 4  = Savings $
    '   currentCol + 5  = Savings %
    Set subHeaderRange = wsAnalysis.Range( _
        wsAnalysis.Cells(headerRow + 1, currentCol), _
        wsAnalysis.Cells(headerRow + 1, currentCol + 5) _
    )
    
    subHeaderRange.Cells(1, 1).Value = "Supplier"
    subHeaderRange.Cells(1, 2).Value = "Unit Price"
    subHeaderRange.Cells(1, 3).Value = "Total Price"
    subHeaderRange.Cells(1, 4).Value = "Baseline"
    subHeaderRange.Cells(1, 5).Value = "Savings $"
    subHeaderRange.Cells(1, 6).Value = "Savings %"
    
    With subHeaderRange
        .Interior.Color = RGB(202, 237, 251)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' -------------------------------------------------
    ' Content rows
    ' -------------------------------------------------
    currentRow = headerRow + 2
    
    For j = 2 To lastRow
        If wsPrices.Cells(j, 1).Value = "end" Then Exit For
    
        ' Check if the row contains "Blank" in any supplier column
        containsBlank = False
        For Each cell In wsPrices.Range(wsPrices.Cells(j, supplierStart), wsPrices.Cells(j, supplierEnd))
            If cell.Value = "Blank" Then
                containsBlank = True
                Exit For
            End If
        Next cell
    
        If containsBlank Then
            ' Clear the row block in Analysis if "Blank" is found
            wsAnalysis.Range( _
                wsAnalysis.Cells(currentRow, currentCol), _
                wsAnalysis.Cells(currentRow, currentCol + 5) _
            ).Clear
        Else
            ' ==========================
            ' Unit Price (k-th lowest)
            ' ==========================
            ' ==========================
            ' Unit Price (k-th lowest)  (filtered by C3:C5 + k dropdown)
            ' ==========================
            Dim catRangeAddr As String
            
            Set compareRange = wsPrices.Range( _
                wsPrices.Cells(j, supplierStart), _
                wsPrices.Cells(j, supplierEnd) _
            )
            
            ' C3:C5 – the categorized dropdown range (All / categories)
            catRangeAddr = wsAnalysis.Range( _
                                wsAnalysis.Cells(headerRow + 2, 3), _
                                wsAnalysis.Cells(headerRow + 4, 3) _
                           ).Address(False, False)
            
            If Application.WorksheetFunction.CountIf(compareRange, "<>NA") = 0 Then
                wsAnalysis.Cells(currentRow, currentCol + 1).Value = "NA"
            Else
                wsAnalysis.Cells(currentRow, currentCol + 1).Formula = _
                    "=IF(OR(" & _
                        "COUNTIF(" & catRangeAddr & ",""All"")>0," & _
                        "COUNTIF(" & catRangeAddr & ",Prices!B" & j & ")>0)," & _
                        "IFERROR(SMALL(Prices!" & compareRange.Address(False, False) & "," & _
                            wsAnalysis.Cells(headerRow, currentCol).Address(False, False) & _
                        " ),""Not Found"")," & _
                    """NA"")"
            End If
            
            With wsAnalysis.Cells(currentRow, currentCol + 1)
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With


    
            ' ==========================
            ' Supplier (text)
            ' ==========================
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol)
            formulaCell.Formula = _
                "=IF(" & wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & _
                "=""Not Found"",""Not Found"",INDEX(Prices!$" & _
                Split(wsPrices.Cells(1, supplierStart).Address(True, True), "$")(1) & _
                "$1:$" & Split(wsPrices.Cells(1, supplierEnd).Address(True, True), "$")(1) & _
                "$1,MATCH(" & wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & _
                ",Prices!" & compareRange.Address(False, False) & ",0)))"
    
            With formulaCell
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ==========================
            ' Total Price (Unit * Volume)
            ' ==========================
            Dim uAddr As String, volAddr As String
            uAddr = wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False)
            volAddr = wsAnalysis.Cells(currentRow, 1).Address(False, False)   ' Volume is col 1
    
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 2)
            formulaCell.Formula = _
                "=IF(OR(" & uAddr & "=""""," & _
                       uAddr & "=""NA""," & _
                       uAddr & "=""Not Found""," & _
                       volAddr & "=0)," & _
                   """NA""," & uAddr & "*" & volAddr & ")"
    
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With
    
            ' ==========================
            ' Baseline (Baseline unit * Volume)
            ' ==========================
            Dim baseUnitAddr As String
            baseUnitAddr = wsPrices.Cells(j, supplierStart - 1).Address(False, False, xlA1, True)
    
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 3)
            formulaCell.Formula = _
                "=IF(OR(" & baseUnitAddr & "=0," & volAddr & "=0),""NA""," & _
                   baseUnitAddr & "*" & volAddr & ")"
    
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With
    
            ' ==========================
            ' Savings $ (BaselineTotal - TotalPrice)
            ' ==========================
            Dim bTotalAddr As String, tTotalAddr As String
            bTotalAddr = wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False)
            tTotalAddr = wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False)
    
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 4)
            formulaCell.Formula = _
                "=IF(OR(" & tTotalAddr & "=0," & _
                         tTotalAddr & "=""NA""," & _
                         tTotalAddr & "=""Not Found""," & _
                         bTotalAddr & "=0," & _
                         bTotalAddr & "=""NA""),""NA""," & _
                         bTotalAddr & "-" & tTotalAddr & ")"
    
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
    
                With .formatConditions
                    .Delete
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                    formatConditions.Interior.Color = RGB(217, 217, 217)
                    formatConditions.Font.Color = RGB(0, 0, 0)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206)
                    formatConditions.Font.Color = RGB(0, 97, 0)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206)
                    formatConditions.Font.Color = RGB(156, 0, 6)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156)
                    formatConditions.Font.Color = RGB(156, 87, 0)
                End With
            End With
    
            ' ==========================
            ' Savings % (on totals)
            ' ==========================
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 5)
            formulaCell.Formula = _
                "=IF(OR(" & bTotalAddr & "=0," & _
                         bTotalAddr & "=""NA""," & _
                         tTotalAddr & "=0," & _
                         tTotalAddr & "=""NA""," & _
                         tTotalAddr & "=""Not Found""),""NA"",(" & _
                         bTotalAddr & "-" & tTotalAddr & ")/" & bTotalAddr & ")"
    
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "0%"
    
                With .formatConditions
                    .Delete
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                    formatConditions.Interior.Color = RGB(217, 217, 217)
                    formatConditions.Font.Color = RGB(0, 0, 0)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206)
                    formatConditions.Font.Color = RGB(0, 97, 0)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206)
                    formatConditions.Font.Color = RGB(156, 0, 6)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156)
                    formatConditions.Font.Color = RGB(156, 87, 0)
                End With
            End With
        End If
    
        currentRow = currentRow + 1
    Next j
    
    currentRow = currentRow + 1    ' space before summary
    
    ' -------------------------------------------------
    ' Normalized Total / Savings summary for Lowest
    ' -------------------------------------------------
    
    ' Header cell
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol)
    supplierHeaderRange.Value = "Normalized Total"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    ' Ranges for totals (Total Price and Baseline)
    Dim rangeTotal As String, rangeBase As String
    rangeTotal = wsAnalysis.Range( _
        wsAnalysis.Cells(3, currentCol + 2), _
        wsAnalysis.Cells(lastRow, currentCol + 2) _
    ).Address(False, False)
    
    rangeBase = wsAnalysis.Range( _
        wsAnalysis.Cells(3, currentCol + 3), _
        wsAnalysis.Cells(lastRow, currentCol + 3) _
    ).Address(False, False)
    
    ' Sum of Total Price where both Total and Baseline are numeric
    wsAnalysis.Cells(currentRow + 1, currentCol).Formula2 = _
        "=SUM(IF(ISNUMBER(" & rangeTotal & ")*ISNUMBER(" & rangeBase & ")," & _
        rangeTotal & ",0))"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00"
    End With
    
    ' Sum of Baseline where both Total and Baseline are numeric
    wsAnalysis.Cells(currentRow + 2, currentCol).Formula2 = _
        "=SUM(IF(ISNUMBER(" & rangeTotal & ")*ISNUMBER(" & rangeBase & ")," & _
        rangeBase & ",0))"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 2, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00"
    End With
    
    ' Summary Saving $
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 1)
    supplierHeaderRange.Value = "Saving $"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    Dim sumTotalAddr As String, sumBaseAddr As String
    sumTotalAddr = wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False)
    sumBaseAddr = wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False)
    
    wsAnalysis.Cells(currentRow + 1, currentCol + 1).Formula = _
        "=IF(OR(" & sumBaseAddr & "=0," & _
                 sumBaseAddr & "=""NA""," & _
                 sumTotalAddr & "=0," & _
                 sumTotalAddr & "=""NA""),""NA""," & _
                 sumBaseAddr & "-" & sumTotalAddr & ")"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 1)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00"
    
        With .formatConditions
            .Delete
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
            formatConditions.Interior.Color = RGB(217, 217, 217)
            formatConditions.Font.Color = RGB(0, 0, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            formatConditions.Interior.Color = RGB(198, 239, 206)
            formatConditions.Font.Color = RGB(0, 97, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 199, 206)
            formatConditions.Font.Color = RGB(156, 0, 6)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 235, 156)
            formatConditions.Font.Color = RGB(156, 87, 0)
        End With
    End With
    
    ' Summary Saving %
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 2)
    supplierHeaderRange.Value = "Saving %"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    wsAnalysis.Cells(currentRow + 1, currentCol + 2).Formula = _
        "=IF(OR(" & sumBaseAddr & "=0," & _
                 sumBaseAddr & "=""NA""," & _
                 sumTotalAddr & "=0," & _
                 sumTotalAddr & "=""NA""),""Check Values"",(" & _
                 sumBaseAddr & "-" & sumTotalAddr & ")/" & sumBaseAddr & ")"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 2)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0%"
    
        With .formatConditions
            .Delete
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
            formatConditions.Interior.Color = RGB(217, 217, 217)
            formatConditions.Font.Color = RGB(0, 0, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            formatConditions.Interior.Color = RGB(198, 239, 206)
            formatConditions.Font.Color = RGB(0, 97, 0)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 199, 206)
            formatConditions.Font.Color = RGB(156, 0, 6)
    
            Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
            formatConditions.Interior.Color = RGB(255, 235, 156)
            formatConditions.Font.Color = RGB(156, 87, 0)
        End With
    End With

   
   
   ' ------------------------------------------------ Lowestt
   
   
    
    ' Add a blank space by shifting currentCol to the right for "LSI"
    currentCol = currentCol + 7  ' Add space between "Lowest" and "LSI" columns

    
    '--- Adjust merged header and sub-header definitions ---
    ' Now the merged header will cover 6 columns (from currentCol to currentCol+5)
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow, currentCol), wsAnalysis.Cells(headerRow, currentCol + 5))
        .Merge
        .Value = "LSI Solution"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
    
    ' Update sub-headers to include "Unit Price" between "Supplier" and "Total Price"
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol + 5))
        .Value = Array("Supplier", "Unit Price", "Total Price", "Baseline", "Savings $", "Savings %")
        .Interior.Color = RGB(202, 237, 251)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Initialize current row for data entry (data starts from headerRow + 2)
    currentRow = headerRow + 2
    
    ' Calculate the lowest price per supplier and add formulas for "LSI"
    For j = 2 To lastRow
        If wsPrices.Cells(j, 1).Value = "end" Then Exit For
    
        ' Check if the cell content in Prices is "Blank"
        If wsPrices.Cells(j, supplierStart).Value = "Blank" Then
            ' Clear the entire row block for this iteration
            wsAnalysis.Range(wsAnalysis.Cells(currentRow, currentCol), wsAnalysis.Cells(currentRow, currentCol + 4)).Clear
        Else
            ' ---------------------------
            ' Create a dropdown for selecting Supplier (stays the same)
            Dim supplierValues As String
            supplierValues = Join(Application.Index(wsPrices.Range(wsPrices.Cells(1, supplierStart - 1), wsPrices.Cells(1, supplierEnd)).Value, 1, 0), ",")
            
            With wsAnalysis.Cells(currentRow, currentCol)
                .Validation.Delete
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=supplierValues
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ---------------------------
            ' Insert the formula for Total Price in the new position (now at currentCol+2)
            wsAnalysis.Cells(currentRow, currentCol + 2).Formula = "=IFERROR(" & _
            "INDEX(Prices!" & wsPrices.Cells(j, supplierStart - 1).Address & ":" & wsPrices.Cells(j, supplierEnd).Address & _
            ", MATCH(" & wsAnalysis.Cells(currentRow, currentCol).Address(False, False) & ", Prices!" & wsPrices.Cells(1, supplierStart - 1).Address & ":" & wsPrices.Cells(1, supplierEnd).Address & ", 0))" & _
            ",""NA"")"
              
                
                
            With wsAnalysis.Cells(currentRow, currentCol + 2)
                .NumberFormat = "$#,##0.00"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ---------------------------
            ' Calculate Baseline (now placed at currentCol+3)
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 3)
            formulaCell.Formula = "=IF(" & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "=""NA"",""NA"",IF(Prices!" & _
                wsPrices.Cells(j, supplierStart - 1).Address & "=0,""NA"",Prices!" & _
                wsPrices.Cells(j, supplierStart - 1).Address & "))"

                
                
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With
    
            ' ---------------------------
            ' Insert the new Unit Price column at currentCol+1
            ' This divides Total Price (now in currentCol+2) by Baseline (in currentCol+3)
            wsAnalysis.Cells(currentRow, currentCol + 1).Formula = "=IFERROR(" & _
            wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "/" & _
            wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & ",""NA"")"

                
            With wsAnalysis.Cells(currentRow, currentCol + 1)
                .NumberFormat = "$#,##0.00"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' ---------------------------
            ' Calculate Savings $ (update references: now Baseline is col+3 and Total Price is col+2)
            wsAnalysis.Cells(currentRow, currentCol + 4).Formula = _
                "=IF(OR(" & wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=""NA"", " & _
                     wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=0, " & _
                     wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "=""NA""), ""NA"", " & _
                     wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & " - " & _
                     wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & ")"
                     
                     
            With wsAnalysis.Cells(currentRow, currentCol + 4)
                .NumberFormat = "$#,##0.00"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                With .formatConditions
                    .Delete ' Clear existing conditions
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                    formatConditions.Interior.Color = RGB(217, 217, 217) ' Light gray
                    formatConditions.Font.Color = RGB(0, 0, 0)           ' Black
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206) ' Light green
                    formatConditions.Font.Color = RGB(0, 97, 0)          ' Dark green
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206) ' Light red
                    formatConditions.Font.Color = RGB(156, 0, 6)          ' Dark red
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156) ' Yellow
                    formatConditions.Font.Color = RGB(156, 87, 0)        ' Dark yellow
                End With
            End With
    
            ' ---------------------------
            ' Calculate Savings % (update references: Baseline in col+3 and Total Price in col+2)
            wsAnalysis.Cells(currentRow, currentCol + 5).Formula = _
                "=IF(OR(" & wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=0, " & _
                     wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & "=""NA"", " & _
                     wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "=""NA""), " & _
                     """NA"", (" & wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & " - " & _
                     wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & ")/" & _
                     wsAnalysis.Cells(currentRow, currentCol + 3).Address(False, False) & ")"
            With wsAnalysis.Cells(currentRow, currentCol + 5)
                .NumberFormat = "0%"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
                With .formatConditions
                    .Delete
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                    formatConditions.Interior.Color = RGB(217, 217, 217)
                    formatConditions.Font.Color = RGB(0, 0, 0)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206)
                    formatConditions.Font.Color = RGB(0, 97, 0)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206)
                    formatConditions.Font.Color = RGB(156, 0, 6)
    
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156)
                    formatConditions.Font.Color = RGB(156, 87, 0)
                End With
            End With
        End If
    
        ' Move to the next row in the Analysis sheet
        currentRow = currentRow + 1
    Next j
    
    currentRow = currentRow + 1 'just leaving one space btw the final item and the baseline header
   
    ' Write the supplier name and merge cells across four columns
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol)
    supplierHeaderRange.Value = "Normalized Total"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
   
 
    wsAnalysis.Cells(currentRow + 1, currentCol).Formula2 = _
    "=SUM(IF(ISNUMBER('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & _
    ")*ISNUMBER('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 3), wsAnalysis.Cells(lastRow, currentCol + 3)).Address(False, False) & _
    "), 'Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & ", 0))"



    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
    End With
    
    
    wsAnalysis.Cells(currentRow + 2, currentCol).Formula2 = _
    "=SUM(IF(ISNUMBER('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & _
    ")*ISNUMBER('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 3), wsAnalysis.Cells(lastRow, currentCol + 3)).Address(False, False) & _
    "), 'Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 3), wsAnalysis.Cells(lastRow, currentCol + 3)).Address(False, False) & ", 0))"




    Set formulaCell = wsAnalysis.Cells(currentRow + 2, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
    End With

    
   
    ' Write the supplier name and merge cells across four columns
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 1)
    supplierHeaderRange.Value = "Saving $"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
           
    ' Apply the formula to the target cell
    
    wsAnalysis.Cells(currentRow + 1, currentCol + 1).Formula = "=IF(OR(" & _
    wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=0," & _
    wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=""NA""," & _
    wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=""NA""), " & _
    """NA"", " & _
    wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "-" & _
    wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & ")"

    
    ' Set a reference to the formula cell
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 1)
    
    ' Format the cell
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00" ' Formatting as currency
        
        With .formatConditions
             .Delete ' Clear existing conditions
            
             ' Format for values greater than 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
             formatConditions.Interior.Color = RGB(217, 217, 217) ' Light gray
             formatConditions.Font.Color = RGB(0, 0, 0)         ' Black
    
             ' Format for values greater than 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
             formatConditions.Interior.Color = RGB(198, 239, 206) ' Light green
             formatConditions.Font.Color = RGB(0, 97, 0)         ' Dark green
    
             ' Format for values less than 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
             formatConditions.Interior.Color = RGB(255, 199, 206) ' Light red
             formatConditions.Font.Color = RGB(156, 0, 6)         ' Dark red
    
             ' Format for values equal to 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
             formatConditions.Interior.Color = RGB(255, 235, 156) ' Yellow
             formatConditions.Font.Color = RGB(156, 87, 0)        ' Dark yellow
            
         End With
    End With




    ' Write the supplier name and merge cells across four columns
    Set supplierHeaderRange = wsAnalysis.Cells(currentRow, currentCol + 2)
    supplierHeaderRange.Value = "Saving %"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    wsAnalysis.Cells(currentRow + 1, currentCol + 2).Formula = "=IF(OR(" & _
    wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=0," & _
    wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "=""NA""," & _
    wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=0," & _
    wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & "=""NA""), " & _
    """Check Values"", (" & _
    wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & "-" & _
    wsAnalysis.Cells(currentRow + 1, currentCol).Address(False, False) & ")/" & _
    wsAnalysis.Cells(currentRow + 2, currentCol).Address(False, False) & ")"
    
    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol + 2)
    
    ' Format the cell
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "0%" ' Formatting as percentage
        
        With .formatConditions
             .Delete ' Clear existing conditions
            
             ' Format for values greater than 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
             formatConditions.Interior.Color = RGB(217, 217, 217) ' Light gray
             formatConditions.Font.Color = RGB(0, 0, 0)         ' Black
    
             ' Format for values greater than 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
             formatConditions.Interior.Color = RGB(198, 239, 206) ' Light green
             formatConditions.Font.Color = RGB(0, 97, 0)         ' Dark green
    
             ' Format for values less than 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
             formatConditions.Interior.Color = RGB(255, 199, 206) ' Light red
             formatConditions.Font.Color = RGB(156, 0, 6)         ' Dark red
    
             ' Format for values equal to 0
             Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
             formatConditions.Interior.Color = RGB(255, 235, 156) ' Yellow
             formatConditions.Font.Color = RGB(156, 87, 0)        ' Dark yellow
            
         End With
    End With
    
     
    ' Auto-fit columns
    Cells.EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.ErrorCheckingOptions.NumberAsText = True
    Application.EnableEvents = True


End Sub







