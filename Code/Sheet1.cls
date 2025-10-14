VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    If Target Is Nothing Then Exit Sub
    If Target.CountLarge > 1 Then Exit Sub

    Dim i As Long
    Dim dvType As Long
    Dim newVal As String, oldVal As String
    Dim items As Variant, out As String, exists As Boolean, itm As String
    Dim sel As String, arr As Variant, nm As String
    Dim r As Long, c As Long
    Dim ws As Worksheet

    '========================
    ' 1) MULTI-SELECT in N:N
    '========================
    If Not Intersect(Target, Me.Columns("N")) Is Nothing Then
        On Error Resume Next
        dvType = Target.Validation.Type
        On Error GoTo CleanExit
        If dvType = 3 Then
            Application.EnableEvents = False
            out = ""        ' reset helpers
            exists = False

            newVal = Target.Value
            If Len(newVal) > 0 Then
                Application.Undo
                oldVal = Target.Value

                If Len(Trim$(oldVal)) > 0 Then
                    items = Split(oldVal, ",")
                    For i = LBound(items) To UBound(items)
                        itm = Trim$(CStr(items(i)))
                        If StrComp(itm, newVal, vbTextCompare) = 0 Then
                            exists = True
                        Else
                            If Len(out) = 0 Then
                                out = itm
                            Else
                                out = out & ", " & itm
                            End If
                        End If
                    Next i
                    If Not exists Then
                        If Len(out) = 0 Then out = newVal Else out = out & ", " & newVal
                    End If
                Else
                    out = newVal
                End If
                Target.Value = out
            End If
            Application.EnableEvents = True
        End If
    End If

    '=====================================================
    ' 2) PUSH HEADERS (when row 2 changes now, not row 1)
    '=====================================================
    If Not Intersect(Target, Me.Range("A2:K2")) Is Nothing Then
        Dim hdrRng As Range, hdrCell As Range
        Set hdrRng = Intersect(Target, Me.Range("A2:K2"))
        If Not hdrRng Is Nothing Then
            Application.EnableEvents = False
            sel = Me.Cells(2, "N").Value   ' sheet list in N2
            If Len(Trim$(sel)) > 0 Then
                arr = Split(sel, ",")
                For i = LBound(arr) To UBound(arr)
                    nm = Trim$(CStr(arr(i)))
                    If Len(nm) > 0 And StrComp(nm, Me.Name, vbTextCompare) <> 0 Then
                        On Error Resume Next
                        Set ws = ThisWorkbook.Worksheets(nm)
                        If Not ws Is Nothing Then
                            For Each hdrCell In hdrRng.Cells
                                ws.Cells(2, hdrCell.Column).Value = hdrCell.Value  ' Row 2 now
                            Next hdrCell
                        End If
                        Set ws = Nothing
                        On Error GoTo 0
                    End If
                Next i
            End If
            Application.EnableEvents = True
        End If
    End If

    '=====================================================
    ' 3) PUSH ROW DATA on edits (except headers)
    '=====================================================
    If Not Intersect(Target, Me.Range("A:K")) Is Nothing And Target.Row > 1 Then
        Application.EnableEvents = False
        r = Target.Row
        c = Target.Column

        sel = Me.Cells(r, "N").Value
        If Len(Trim$(sel)) > 0 Then
            arr = Split(sel, ",")
            For i = LBound(arr) To UBound(arr)
                nm = Trim$(CStr(arr(i)))
                If Len(nm) > 0 And StrComp(nm, Me.Name, vbTextCompare) <> 0 Then
                    On Error Resume Next
                    Set ws = ThisWorkbook.Worksheets(nm)
                    If Not ws Is Nothing Then
                        ws.Cells(r, c).Value = Target.Value
                    End If
                    Set ws = Nothing
                    On Error GoTo 0
                End If
            Next i
        End If
        Application.EnableEvents = True
    End If

CleanExit:
    If Application.EnableEvents = False Then Application.EnableEvents = True
End Sub
