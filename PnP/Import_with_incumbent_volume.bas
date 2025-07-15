Attribute VB_Name = "Module1"
Option Explicit

Sub FindEndRowAndIterate()
    Dim ws As Worksheet
    Dim searchRange As Range
    Dim foundCell As Range
    Dim LastRowEnd As Long
    Dim ProjectRow As Long

    On Error GoTo ImportSheetNotFound
    Set ws = ThisWorkbook.Sheets("Import")
    On Error GoTo 0

    Set searchRange = ws.Range("A1:A" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
    Set foundCell = searchRange.Find(What:="end", LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        LastRowEnd = foundCell.Row
        For ProjectRow = 2 To LastRowEnd - 1
            CopyTabsFromMultipleWorkbooks_InPlace ProjectRow
        Next ProjectRow
        MsgBox "Macro execution complete for all project rows.", vbInformation, "Process Finished"
    Else
        MsgBox "The marker 'end' was not found in column A of the 'Import' sheet.", vbExclamation
    End If

    Exit Sub

ImportSheetNotFound:
    MsgBox "The 'Import' sheet was not found in this workbook.", vbCritical
End Sub

Sub CopyTabsFromMultipleWorkbooks_InPlace(ProjectRow As Long)
    Dim BaseFolderPath As String, FileName As String, sheetName As String
    Dim DestinationName As String, CurrentWorkbook As Workbook, SourceWorkbook As Workbook
    Dim ws As Worksheet, wsTab As Worksheet, existingSheet As Worksheet
    Dim findEnd As Range, ImportSheets As Variant
    Dim EndColumn As Long, i As Long, k As Long
    Dim Value1 As Long, Value2 As Long, rowNumber0 As Long
    Dim sourceColumnLetter As String, headerAddress As String
    Dim formulaString As String, formulaCell As Range
    Dim RandomColor As Long, Response As VbMsgBoxResult

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.ErrorCheckingOptions.NumberAsText = False
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic

    BaseFolderPath = ThisWorkbook.Path & "\"
    Set CurrentWorkbook = ThisWorkbook

    With ThisWorkbook.Sheets("Import").Rows(1)
        Set findEnd = .Find(What:="end", LookAt:=xlWhole, MatchCase:=False)
    End With

    If findEnd Is Nothing Then
        MsgBox "'end' column not found in 'Import' sheet row 1.", vbExclamation
        GoTo Cleanup
    End If

    EndColumn = findEnd.Column

    With ThisWorkbook.Sheets("Import")
        ImportSheets = .Range(.Cells(1, 1), .Cells(1, EndColumn - 1)).Value
    End With

    For i = 1 To UBound(ImportSheets, 2)
        FileName = ImportSheets(1, i) & ".xlsx"
        sheetName = ThisWorkbook.Sheets("Import").Cells(ProjectRow, i).Value

        If UCase(Trim(sheetName)) = "NA" Then GoTo SkipSheet

        DestinationName = ImportSheets(1, i)
        Response = MsgBox("Import '" & sheetName & "' from '" & FileName & "'?", vbYesNo + vbQuestion)
        If Response = vbNo Then GoTo SkipSheet

        On Error Resume Next
        Set existingSheet = CurrentWorkbook.Sheets(DestinationName)
        On Error GoTo 0
        If Not existingSheet Is Nothing Then
            Application.DisplayAlerts = False
            existingSheet.Delete
            Application.DisplayAlerts = True
        End If

        On Error Resume Next
        Set SourceWorkbook = Workbooks.Open(BaseFolderPath & FileName, ReadOnly:=True)
        On Error GoTo 0

        If Not SourceWorkbook Is Nothing Then
            On Error Resume Next
            Set ws = SourceWorkbook.Sheets(sheetName)
            On Error GoTo 0
            If Not ws Is Nothing Then
                ws.Copy After:=CurrentWorkbook.Sheets(CurrentWorkbook.Sheets.Count)
                CurrentWorkbook.Sheets(CurrentWorkbook.Sheets.Count).Name = DestinationName
            Else
                MsgBox "Sheet '" & sheetName & "' not found in '" & FileName & "'.", vbExclamation
            End If
            SourceWorkbook.Close SaveChanges:=False
        Else
            MsgBox "Could not open '" & FileName & "'.", vbExclamation
        End If

SkipSheet:
    Next i

    Dim importedCount As Long
    For Each wsTab In CurrentWorkbook.Sheets
        If wsTab.Name <> "Import" And wsTab.Name <> "Sheet1" Then importedCount = importedCount + 1
    Next

    If importedCount = 0 Then GoTo Cleanup

    Set ws = CurrentWorkbook.Sheets.Add(Before:=CurrentWorkbook.Sheets(1))
    ws.Name = "Prices"
    
    'Set the first column header
    ws.Cells(1, 1).Value = "Incumbent"
    'Set the second column header
    ws.Cells(1, 2).Value = "Volumen"

    'Copy the remaining headers from the Import sheet
    With ThisWorkbook.Sheets("Import")
        'Start copying from the first column of ImportSheets, as the first two headers are already set
        .Range(.Cells(1, 1), .Cells(1, EndColumn - 1)).Offset(0, 0).Copy Destination:=ws.Cells(1, 3)
    End With

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, EndColumn + 1)) 'Adjust EndColumn to include the new "Incumbent" and "Volumen" columns
        .Interior.Color = RGB(255, 192, 0)
        .Font.Name = "Aptos Narrow"
        .Font.Size = 12
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    Value1 = ThisWorkbook.Sheets("Import").Cells(ProjectRow, EndColumn + 2).Value
    Value2 = ThisWorkbook.Sheets("Import").Cells(ProjectRow, EndColumn + 3).Value
    sourceColumnLetter = ThisWorkbook.Sheets("Import").Cells(ProjectRow, EndColumn + 1).Text

    For k = 2 To Value2 - Value1 + 2
        rowNumber0 = Value1 + (k - 2)
        'Loop from 3 to EndColumn + 1 to account for the new "Incumbent" and "Volumen" columns
        For i = 3 To EndColumn + 1
            Set formulaCell = ws.Cells(k, i)
            headerAddress = ws.Cells(1, i).Address(RowAbsolute:=True, ColumnAbsolute:=False)

            formulaString = "=IFERROR(INDIRECT(" & _
                            Chr(34) & "'" & Chr(34) & " & " & headerAddress & " & " & _
                            Chr(34) & "'!" & sourceColumnLetter & Chr(34) & " & " & _
                            "ROW(" & sourceColumnLetter & rowNumber0 & ")), " & _
                            Chr(34) & "NA" & Chr(34) & ")"

            With formulaCell
                .Formula = formulaString
                .Borders.LineStyle = xlContinuous
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .NumberFormat = "$#,##0.00"
            End With
        Next i
    Next k

    Dim Links As Variant, Lnk As Variant
    Links = CurrentWorkbook.LinkSources(xlLinkTypeExcelLinks)
    If Not IsEmpty(Links) Then
        For Each Lnk In Links
            On Error Resume Next
            CurrentWorkbook.BreakLink Name:=Lnk, Type:=xlLinkTypeExcelLinks
            On Error GoTo 0
        Next Lnk
    End If

    Links = CurrentWorkbook.LinkSources(xlOLELinks)
    If Not IsEmpty(Links) Then
        For Each Lnk In Links
            On Error Resume Next
            CurrentWorkbook.BreakLink Name:=Lnk, Type:=xlOLELinks
            On Error GoTo 0
        Next Lnk
    End If

    ws.Tab.Color = RGB(128, 0, 128)
    RandomColor = RandomPaleColor()
    For Each wsTab In CurrentWorkbook.Sheets
        If wsTab.Name <> "Prices" And wsTab.Name <> "Import" Then
            wsTab.Tab.Color = RandomColor
        End If
    Next

    With CurrentWorkbook.Sheets("Prices")
        .Activate
        .Columns.AutoFit
        With ActiveWindow
            .Zoom = 85
            .ScrollRow = 1
            .ScrollColumn = 1
        End With
        .Cells(1, 1).Select
    End With

    On Error Resume Next
    Set wsTab = CurrentWorkbook.Sheets("Sheet1")
    If Not wsTab Is Nothing Then
        If wsTab.usedRange.Cells.Count = 1 And IsEmpty(wsTab.Cells(1, 1).Value) Then
            Application.DisplayAlerts = False
            wsTab.Delete
            Application.DisplayAlerts = True
        End If
    End If
    On Error GoTo 0

Cleanup:
    Application.CalculateFullRebuild
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.ErrorCheckingOptions.NumberAsText = True
    Application.CutCopyMode = False
End Sub

Function RandomPaleColor() As Long
    Randomize
    RandomPaleColor = RGB( _
        Int((200 - 180 + 1) * Rnd + 180), _
        Int((200 - 180 + 1) * Rnd + 180), _
        Int((200 - 180 + 1) * Rnd + 180))
End Function



