Option Explicit

Sub UnmergeVertical_CopyDown_ActiveSheet()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim c As Range, addr As Variant, rng As Range
    Dim tr As Long, br As Long, lc As Long, rc As Long, r As Long, v

    Application.ScreenUpdating = False
    On Error Resume Next
    For Each c In ws.UsedRange
        If c.MergeCells Then
            If c.MergeArea.Rows.Count > 1 Then dict(c.MergeArea.Address(False, False)) = 1
        End If
    Next c
    On Error GoTo 0

    For Each addr In dict.Keys
        Set rng = ws.Range(CStr(addr))
        tr = rng.Row: lc = rng.Column
        br = tr + rng.Rows.Count - 1
        rc = lc + rng.Columns.Count - 1
        v = rng.Cells(1, 1).Value
        rng.UnMerge
        For r = tr To br
            If rc > lc Then ws.Range(ws.Cells(r, lc), ws.Cells(r, rc)).Merge
            ws.Cells(r, lc).Value = v
        Next r
    Next addr

    Application.ScreenUpdating = True
End Sub
