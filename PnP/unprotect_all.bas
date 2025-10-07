Option Explicit

'==================== CONFIG ====================
Private Const PASSWORD As String = "CB"   ' <-- set your password here
Private Const INCLUDE_SUBFOLDERS As Boolean = False
'================================================

Public Sub Unprotect_All_Excel_InFolder()
    Dim app As Application
    Dim wbHost As Workbook
    Dim folderPath As String

    Set app = Application
    Set wbHost = ThisWorkbook ' the workbook that contains this macro
    folderPath = GetFolderPathOfWorkbook(wbHost)

    If Len(folderPath) = 0 Then
        MsgBox "Could not resolve folder of the host workbook.", vbExclamation
        Exit Sub
    End If

    app.ScreenUpdating = False
    app.DisplayAlerts = False
    app.EnableEvents = False

    Dim okCount As Long, failCount As Long, skipCount As Long
    okCount = 0: failCount = 0: skipCount = 0

    ProcessFolder folderPath, okCount, failCount, skipCount

    app.ScreenUpdating = True
    app.DisplayAlerts = True
    app.EnableEvents = True

    MsgBox "Done." & vbCrLf & _
           "Unprotected & saved: " & okCount & vbCrLf & _
           "Failed:              " & failCount & vbCrLf & _
           "Skipped:             " & skipCount, vbInformation
End Sub

Private Function GetFolderPathOfWorkbook(wb As Workbook) As String
    On Error GoTo EH
    Dim p As String
    p = wb.Path
    If Len(p) > 0 Then
        If Right$(p, 1) <> "\" Then p = p & "\"
        GetFolderPathOfWorkbook = p
    Else
        GetFolderPathOfWorkbook = ""
    End If
    Exit Function
EH:
    GetFolderPathOfWorkbook = ""
End Function

Private Sub ProcessFolder(ByVal folderPath As String, _
                          ByRef okCount As Long, ByRef failCount As Long, ByRef skipCount As Long)
    Dim fso As Object, folder As Object, file As Object, subf As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set folder = fso.GetFolder(folderPath)

    ' Files in folder
    For Each file In folder.Files
        If IsExcelCandidate(file) Then
            If LCase(file.Path) = LCase(ThisWorkbook.FullName) Then
                skipCount = skipCount + 1 ' do not process the host workbook
            Else
                ProcessWorkbook file.Path, okCount, failCount
            End If
        End If
    Next

    ' Optional recursion
    If INCLUDE_SUBFOLDERS Then
        For Each subf In folder.SubFolders
            ProcessFolder subf.Path, okCount, failCount, skipCount
        Next
    End If
End Sub

Private Function IsExcelCandidate(fileObj As Object) As Boolean
    Dim name As String, ext As String
    name = LCase(fileObj.Name)
    If Left$(name, 2) = "~$" Then
        IsExcelCandidate = False ' temp/lock file
        Exit Function
    End If
    ext = LCase$(fileObj.Name)
    ext = LCase$(Right$(ext, Len(ext) - InStrRev(ext, ".")))
    Select Case ext
        Case "xlsx", "xlsm", "xlsb", "xls"
            IsExcelCandidate = True
        Case Else
            IsExcelCandidate = False
    End Select
End Function

Private Sub ProcessWorkbook(ByVal filePath As String, _
                            ByRef okCount As Long, ByRef failCount As Long)
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Application.Workbooks.Open(Filename:=filePath, UpdateLinks:=0, ReadOnly:=False)
    If Err.Number <> 0 Or wb Is Nothing Then
        Debug.Print "OPEN FAIL: "; filePath; " | "; Err.Description
        Err.Clear
        failCount = failCount + 1
        Exit Sub
    End If
    On Error GoTo 0

    ' Try to unprotect workbook structure (if protected)
    On Error Resume Next
    wb.Unprotect PASSWORD
    Err.Clear

    ' Unprotect all worksheets
    Dim ws As Worksheet, sheetErr As Boolean
    sheetErr = False
    For Each ws In wb.Worksheets
        ws.Unprotect PASSWORD
        If Err.Number <> 0 Then
            Debug.Print "Sheet unprotect failed: [" & wb.Name & "] '" & ws.Name & "' | "; Err.Description
            Err.Clear
            sheetErr = True
        End If
    Next ws

    ' Save
    On Error Resume Next
    wb.Save
    If Err.Number <> 0 Then
        Debug.Print "SAVE FAIL: "; filePath; " | "; Err.Description
        Err.Clear
        sheetErr = True
    End If

    wb.Close SaveChanges:=False
    On Error GoTo 0

    If sheetErr Then failCount = failCount + 1 Else okCount = okCount + 1
End Sub
