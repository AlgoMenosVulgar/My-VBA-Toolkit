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
    supplierStart = 8
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
    
    ' Insert values from Prices!D while clearing any cell whose source says "Blank"
    Dim rowNum As Long, srcVal As Variant
    
    For rowNum = 3 To lastRow
        srcVal = wsPrices.Cells(rowNum - 1, "F").Value        'read once, faster than repeated sheet calls volumn column
        
        With wsAnalysis.Cells(rowNum, currentCol)
            If srcVal = "Blank" Then
                .Clear                                         'remove content + all formatting
            Else
                .Formula = "=Prices!F" & (rowNum - 1)          'keep live link for every other case
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
        srcVal = wsPrices.Cells(rowNum - 1, "G").Value         'read once, avoid repeated sheet calls
        
        With wsAnalysis.Cells(rowNum, currentCol)
            If srcVal = "Blank" Then
                .Clear                                          'remove content and all formatting
            Else
                .Formula = "=Prices!G" & (rowNum - 1)           'keep live link for every other case
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
    For Each cell In wsPrices.Range("C2:C" & lastRow)
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
                formulaString = "=IF(OR(" & _
                "COUNTIF(" & dropdownRange.Address(False, False) & ", ""All"")>0," & _
                "COUNTIF(" & dropdownRange.Address(False, False) & ", Prices!C" & j & ")>0)," & _
                "'Prices'!" & wsPrices.Cells(j, i).Address(False, True) & ", ""NA"")"

                With formulaCell
                    .Formula = formulaString
                    .Borders.LineStyle = xlContinuous
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
                End With
                
                ' Savings $ (same as before)
                Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 1)
                formulaString = "=IF(OR(" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=0, " & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=""NA"", " & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "=""NA""), " & _
                    """NA"", " & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "-(Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "))"

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
                formulaString = "=IF(OR(" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=0, " & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & "=""NA"", " & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "=""NA""), " & _
                    """NA"", " & _
                    "(Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & "-Analysis!" & wsAnalysis.Cells(j + 1, currentCol).Address(False, True) & ")/" & _
                    "Analysis!" & wsAnalysis.Cells(j + 1, 2).Address(False, True) & ")"


                With formulaCell
                    .Formula = formulaString
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
                    wsAnalysis.Cells(1, (lastColumn - 5) * 5 + 10).Address(False, False) & ")-" & _
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
        wsAnalysis.Cells(currentRow + 1, currentCol).Formula = "=SUMIF(" & wsAnalysis.Cells(3, BaselineColumn).Address(False, False) & ":" & wsAnalysis.Cells(lastRow, BaselineColumn).Address(False, False) & ",""<>NA""," & wsAnalysis.Cells(3, currentCol).Address(False, False) & ":" & wsAnalysis.Cells(lastRow, currentCol).Address(False, False) & ")"

        Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol)
        With formulaCell
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
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
    
   '---start
   ' Merge and set header for "Incumbent Solution"
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow, currentCol), wsAnalysis.Cells(headerRow, currentCol + 4))
        .Merge
        .Value = "Incumbent Solution"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 192, 0)
        .Borders.Weight = xlThin
    End With
    
    ' Add sub-headers for "Incumbent Solution"
    With wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol + 4)) 'test
        .Value = Array("Supplier", "Total Price", "Baseline", "Savings $", "Savings %")
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
            ' Assign Supplier Name
            wsAnalysis.Cells(currentRow, currentCol).Formula = "=Prices!" & wsPrices.Cells(j, 1).Address(False, False)
            
            ' Format Total Price cell
            With wsAnalysis.Cells(currentRow, currentCol)
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' Total Price formula
            wsAnalysis.Cells(currentRow, currentCol + 1).Formula = _
            "=IF(OR(COUNTIF(" & dropdownRange.Address(False, False) & ", ""All"")>0, " & _
            "COUNTIF(" & dropdownRange.Address(False, False) & ", Prices!C" & j & ")>0), " & _
            "IFERROR(INDEX(Prices!" & wsPrices.Cells(j, supplierStart - 1).Address & ":" & _
            wsPrices.Cells(j, supplierEnd).Address & ", MATCH(" & _
            wsAnalysis.Cells(currentRow, currentCol).Address(False, False) & ", Prices!" & _
            wsPrices.Cells(1, supplierStart - 1).Address & ":" & wsPrices.Cells(1, supplierEnd).Address & ", 0)), ""NA""), ""NA"")"


    
            ' Format Total Price cell
            With wsAnalysis.Cells(currentRow, currentCol + 1)
                .NumberFormat = "$#,##0.00"
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
            End With
    
            ' Calculate Baseline
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 2)
            formulaCell.Formula = _
                "=IF(OR(COUNTIF(" & dropdownRange.Address(False, False) & ", ""All"")>0, " & _
                "COUNTIF(" & dropdownRange.Address(False, False) & ", Prices!C" & j & ")>0), " & _
                "IF(Prices!" & wsPrices.Cells(j, supplierStart - 1).Address & "=0, ""NA"", " & _
                "Prices!" & wsPrices.Cells(j, supplierStart - 1).Address & "), ""NA"")"

    
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With
    
            ' Calculate Savings $
            wsAnalysis.Cells(currentRow, currentCol + 3).Formula = _
                "=IF(OR(" & wsAnalysis.Cells(currentRow, currentCol + 2).Address & "=""NA"", " & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address & "=0, " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address & "=""NA""), ""NA"", " & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address & " - " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address & ")"
    
            ' Format Savings $ cell
            With wsAnalysis.Cells(currentRow, currentCol + 3)
                .NumberFormat = "$#,##0.00"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
    
                With .formatConditions
                    .Delete ' Clear existing conditions
    
                    ' Format for "NA"
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                    formatConditions.Interior.Color = RGB(217, 217, 217)
                    formatConditions.Font.Color = RGB(0, 0, 0)
    
                    ' Format for values greater than 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206)
                    formatConditions.Font.Color = RGB(0, 97, 0)
    
                    ' Format for values less than 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206)
                    formatConditions.Font.Color = RGB(156, 0, 6)
    
                    ' Format for values equal to 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156)
                    formatConditions.Font.Color = RGB(156, 87, 0)
                End With
            End With
    
            ' Calculate Savings %
            wsAnalysis.Cells(currentRow, currentCol + 4).Formula = _
                "=IF(OR(" & wsAnalysis.Cells(currentRow, currentCol + 2).Address & "=0, " & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address & "=""NA"", " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address & "=""NA""), ""NA"", (" & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address & " - " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address & ") / " & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address & ")"
    
            ' Format Savings % cell
            With wsAnalysis.Cells(currentRow, currentCol + 4)
                .NumberFormat = "0%"
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
                .Borders.LineStyle = xlContinuous
    
                With .formatConditions
                    .Delete ' Clear existing conditions
    
                    ' Format for "NA"
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="NA")
                    formatConditions.Interior.Color = RGB(217, 217, 217)
                    formatConditions.Font.Color = RGB(0, 0, 0)
    
                    ' Format for values greater than 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                    formatConditions.Interior.Color = RGB(198, 239, 206)
                    formatConditions.Font.Color = RGB(0, 97, 0)
    
                    ' Format for values less than 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 199, 206)
                    formatConditions.Font.Color = RGB(156, 0, 6)
    
                    ' Format for values equal to 0
                    Set formatConditions = .Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="0")
                    formatConditions.Interior.Color = RGB(255, 235, 156)
                    formatConditions.Font.Color = RGB(156, 87, 0)
                End With
            End With
        End If
    
        ' Move to the next row
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
    "=SUMIF('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(1, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & _
    ",""<>NA"",'Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(1, currentCol + 1), wsAnalysis.Cells(lastRow, currentCol + 1)).Address(False, False) & ")"


    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
    End With
    
    
    wsAnalysis.Cells(currentRow + 2, currentCol).Formula2 = _
    "=SUMIF('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(1, currentCol + 1), wsAnalysis.Cells(lastRow, currentCol + 1)).Address(False, False) & _
    ",""<>NA"",'Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(1, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & ")"


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
    
    ' Move to the next set of columns
    currentCol = currentCol + 6

   
   '--- End


   ' Add "Lowest" header
    Set supplierHeaderRange = wsAnalysis.Range(wsAnalysis.Cells(headerRow, currentCol + 1), wsAnalysis.Cells(headerRow, currentCol + 4))
       
    supplierHeaderRange.Merge
    supplierHeaderRange.Value = "Lowest"
    supplierHeaderRange.HorizontalAlignment = xlCenter
    supplierHeaderRange.VerticalAlignment = xlCenter
    supplierHeaderRange.Interior.Color = RGB(255, 192, 0)
    supplierHeaderRange.Borders.Weight = xlThin
    
    ' Add dropdown list with values 1, 2, 3, 4, 5
    'Dim dropdownRange As Range
    Set dropdownRange = wsAnalysis.Range(wsAnalysis.Cells(headerRow, currentCol), wsAnalysis.Cells(headerRow, currentCol)) ' Adjust range to where you want the dropdown
    
    With dropdownRange
    ' Apply data validation
    With .Validation
        .Delete ' Remove any existing validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="1,2,3,4,5" ' Add values 1, 2, 3, 4, 5 in the dropdown
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ' Center values horizontally and vertically
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter

    ' Add borders to all sides
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin

    ' Apply background fill color (RGB: 202, 237, 251)
    .Interior.Color = RGB(202, 237, 251)
    End With
    
    ' Add sub-headers for "Lowest"
    Set subHeaderRange = wsAnalysis.Range(wsAnalysis.Cells(headerRow + 1, currentCol), wsAnalysis.Cells(headerRow + 1, currentCol + 4))
    subHeaderRange.Cells(1, 1).Value = "Supplier"
    subHeaderRange.Cells(1, 2).Value = "Total Price"
    subHeaderRange.Cells(1, 3).Value = "Baseline"
    subHeaderRange.Cells(1, 4).Value = "Savings $"
    subHeaderRange.Cells(1, 5).Value = "Savings %"
    subHeaderRange.Interior.Color = RGB(202, 237, 251)
    subHeaderRange.Borders.Weight = xlThin
    subHeaderRange.HorizontalAlignment = xlCenter
    subHeaderRange.VerticalAlignment = xlCenter
    
    ' Calculate the lowest price per supplier and add formulas
    currentRow = headerRow + 2
    For j = 2 To lastRow
        If wsPrices.Cells(j, 1).Value = "end" Then Exit For
        
        ' Check if the row contains "Blank"
        containsBlank = False
        For Each cell In wsPrices.Range(wsPrices.Cells(j, supplierStart), wsPrices.Cells(j, supplierEnd))
            If cell.Value = "Blank" Then
                containsBlank = True
                Exit For
            End If
        Next cell
        
        If containsBlank Then
            ' Clear the row block in Analysis if "Blank" is found
            wsAnalysis.Range(wsAnalysis.Cells(currentRow, currentCol), wsAnalysis.Cells(currentRow, currentCol + 3)).Clear
        Else
            ' Calculate Supplier (Supplier Column)
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol)
            formulaCell.Formula = _
                "=IF(" & wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & "=""Not Found"", " & _
                """Not Found"", INDEX(Prices!$" & Split(wsPrices.Cells(1, supplierStart).Address(True, True), "$")(1) & _
                "$1:$" & Split(wsPrices.Cells(1, supplierEnd).Address(True, True), "$")(1) & _
                "$1, MATCH(" & wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & ", " & _
                "Prices!" & wsPrices.Range(wsPrices.Cells(j, supplierStart), wsPrices.Cells(j, supplierEnd)).Address(False, False) & ", 0)))"
            
            formulaCell.HorizontalAlignment = xlCenter
            formulaCell.VerticalAlignment = xlCenter
            formulaCell.Borders.LineStyle = xlContinuous

        
            ' Find the supplier column with the minimum price
            Dim minValue As Double
            Dim minColumn As Long
            
            ' … dentro de For j …
            Dim compareRange As Range
            Dim vPos       As Variant
        
            ' 1) Definir rango de comparación
            Set compareRange = wsPrices.Range( _
                  wsPrices.Cells(j, supplierStart), _
                  wsPrices.Cells(j, supplierEnd) _
            )
        
            ' 2) Si NO hay ningún precio distinto de "NA", salir con "NA"
            If Application.WorksheetFunction.CountIf(compareRange, "<>NA") = 0 Then
                wsAnalysis.Cells(currentRow, currentCol + 1).Value = "NA"
            Else
                ' 3) calcular el valor mínimo
                minValue = Application.WorksheetFunction.Min(compareRange)
                ' 4) ubicar posición con Application.Match (no WorksheetFunction.Match)
                vPos = Application.Match(minValue, compareRange, 0)
                If Not IsError(vPos) Then
                    minColumn = supplierStart + vPos - 1
                    ' … aquí vendría el resto de su lógica para escribir el SMALL(...) o INDEX(...)
                    wsAnalysis.Cells(currentRow, currentCol + 1).Formula = _
                        "=SMALL(Prices!" & compareRange.Address(False, False) & ", " & _
                        wsAnalysis.Cells(1, currentCol).Address(True, True) & ")"
                Else
                    wsAnalysis.Cells(currentRow, currentCol + 1).Value = "NA"
                End If
            End If
            ' … resto del bucle …

        
            ' Calculate Total Price (Total Price Column)
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 1)
            
            
            ' Use SMALL function to calculate the k-th smallest value
            formulaCell.Formula = "=IFERROR(SMALL(Prices!" & _
            wsPrices.Range(wsPrices.Cells(j, supplierStart), wsPrices.Cells(j, supplierEnd)).Address(False, False) & _
            ", " & wsAnalysis.Cells(1, currentCol).Address(True, True) & "),""Not Found"")"
            
            
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With

            
            ' Calculate Baseline
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 2)
            formulaCell.Formula = "=IF(Prices!" & wsPrices.Cells(j, supplierStart - 1).Address & "=0, ""NA"", Prices!" & wsPrices.Cells(j, supplierStart - 1).Address & ")"
            
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With
        
            ' Calculate Savings $ (Savings $ Column)
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 3)
            formulaCell.Formula = "=IF(OR(" & _
            wsAnalysis.Cells(currentRow, currentCol + 1).Address & "=0, " & _
            wsAnalysis.Cells(currentRow, currentCol + 1).Address & "=""NA"", " & _
            wsAnalysis.Cells(currentRow, currentCol + 1).Address & "=""Not Found"", " & _
            wsAnalysis.Cells(currentRow, currentCol + 2).Address & "=0, " & _
            wsAnalysis.Cells(currentRow, currentCol + 2).Address & "=""NA""), " & _
            """NA"", " & _
            wsAnalysis.Cells(currentRow, currentCol + 2).Address & " - " & _
            wsAnalysis.Cells(currentRow, currentCol + 1).Address & ")"

            
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
                
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
        
            ' Calculate Savings % (Savings % Column)
            Set formulaCell = wsAnalysis.Cells(currentRow, currentCol + 4)
            formulaCell.Formula = "=IF(OR(" & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "=0, " & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & "=""NA"", " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & "=0, " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & "=""NA"", " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & "=""Not Found""), " & _
                """NA"", (" & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & " - " & _
                wsAnalysis.Cells(currentRow, currentCol + 1).Address(False, False) & ")/" & _
                wsAnalysis.Cells(currentRow, currentCol + 2).Address(False, False) & ")"

            
            With formulaCell
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "0%"
                
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
        End If
        
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
    "=SUM(IF(ISNUMBER(Analysis!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 1), wsAnalysis.Cells(lastRow, currentCol + 1)).Address(False, False) & _
    ")*ISNUMBER(Analysis!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & _
    "), Analysis!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 1), wsAnalysis.Cells(lastRow, currentCol + 1)).Address(False, False) & ", 0))"


    Set formulaCell = wsAnalysis.Cells(currentRow + 1, currentCol)
    With formulaCell
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .NumberFormat = "$#,##0.00" ' Formatting as currency ----- here
    End With
    
    
    wsAnalysis.Cells(currentRow + 2, currentCol).Formula2 = _
    "=SUM(IF(ISNUMBER('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 1), wsAnalysis.Cells(lastRow, currentCol + 1)).Address(False, False) & _
    ")*ISNUMBER('Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & _
    "), 'Analysis'!" & wsAnalysis.Range(wsAnalysis.Cells(3, currentCol + 2), wsAnalysis.Cells(lastRow, currentCol + 2)).Address(False, False) & ", 0))"


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
    
    ' Add a blank space by shifting currentCol to the right for "LSI"
    currentCol = currentCol + 6  ' Add space between "Lowest" and "LSI" columns

    
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





