Option Explicit

Sub DeleteAllSheetsExceptImport()
    Dim ws As Worksheet
    Dim hasImport As Boolean
    
    ' Check that "Import" sheet exists
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Import" Then
            hasImport = True
            Exit For
        End If
    Next ws
    
    If Not hasImport Then
        MsgBox "No existe una hoja llamada ""Import"". No se eliminaron hojas.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Delete every sheet except "Import"
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Import" Then
            ws.Delete
        End If
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Se han eliminado todas las hojas excepto ""Import"".", vbInformation
End Sub

