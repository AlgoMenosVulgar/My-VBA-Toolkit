Attribute VB_Name = "Module1"
Option Explicit

Sub LockAllExceptColoredCells()
    Dim ws As Worksheet
    Dim cell As Range, usedRange As Range
    Dim color1 As Long, color2 As Long
    Const SHEET_PASSWORD As String = "ES"

    ' Define target fill colors
    color1 = RGB(197, 217, 241) ' Light blue
    color2 = RGB(255, 255, 153) ' Light yellow

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=SHEET_PASSWORD
        On Error GoTo 0

        ' Lock everything first
        ws.Cells.Locked = True

        ' Work on the used range only
        Set usedRange = ws.usedRange

        For Each cell In usedRange
            ' Use MergeArea to unlock entire merged block if needed
            If cell.Interior.Color = color1 Or cell.Interior.Color = color2 Then
                If cell.MergeCells Then
                    cell.MergeArea.Locked = False
                Else
                    cell.Locked = False
                End If
            End If
        Next cell

        ' Re-protect the sheet
        ws.Protect Password:=SHEET_PASSWORD, _
                   AllowFormattingCells:=True, _
                   AllowSorting:=True, _
                   AllowFiltering:=True
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Done: Only light green and yellow cells (even merged ones) are editable.", vbInformation
End Sub




