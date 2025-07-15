Attribute VB_Name = "Module1"
Option Explicit

Sub UnlockAllSheetsAndCells()
    Dim ws As Worksheet
    Const SHEET_PASSWORD As String = "CB"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=SHEET_PASSWORD
        On Error GoTo 0

        ' Unlock every cell on the sheet
        ws.Cells.Locked = False
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "All sheets are now unprotected and fully editable.", vbInformation
End Sub


