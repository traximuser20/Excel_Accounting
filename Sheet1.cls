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
    ' 1) MULTI-SELECT in L:L
    '========================
    If Not Intersect(Target, Me.Columns("L")) Is Nothing Then
        On Error Resume Next
        dvType = Target.Validation.Type
        On Error GoTo CleanExit
        If dvType = 3 Then
            Application.EnableEvents = False
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
    ' 2) PUSH HEADERS (when L2 changes)
    '=====================================================
    If Not Intersect(Target, Me.Range("L2")) Is Nothing Then
        Application.EnableEvents = False
        sel = Me.Cells(2, "L").Value
        If Len(Trim$(sel)) > 0 Then
            arr = Split(sel, ",")
            For i = LBound(arr) To UBound(arr)
                nm = Trim$(CStr(arr(i)))
                If Len(nm) > 0 And StrComp(nm, Me.Name, vbTextCompare) <> 0 Then
                    On Error Resume Next
                    Set ws = ThisWorkbook.Worksheets(nm)
                    If Not ws Is Nothing Then
                        ' Copy headers A1:J1 into each selected sheet
                        Me.Range("A1:J1").Copy ws.Range("A1")
                    End If
                    Set ws = Nothing
                    On Error GoTo 0
                End If
            Next i
        End If
        Application.EnableEvents = True
    End If

    '=====================================================
    ' 3) PUSH ROW DATA on edits (except headers)
    '=====================================================
    If Not Intersect(Target, Me.Range("A:J")) Is Nothing And Target.Row > 1 Then
        Application.EnableEvents = False
        r = Target.Row
        c = Target.Column

        sel = Me.Cells(r, "L").Value
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

    '=====================================================
    ' 4) LOG CREDIT/DEBIT TO BALANCE SHEET
    '=====================================================
    ' ?? Adjust these columns to match your Master sheet
    Const COL_CREDIT As Long = 9   ' Column I
    Const COL_DEBIT As Long = 10   ' Column J

    If (Target.Column = COL_CREDIT Or Target.Column = COL_DEBIT) And Target.Row > 1 Then
        Application.EnableEvents = False

        Dim bs As Worksheet
        Dim lastRowBS As Long
        Dim sheetList As Variant
        Dim entryType As String
        Dim f As Range

        Set bs = ThisWorkbook.Worksheets("Balance Sheet")

        If Target.Column = COL_CREDIT Then
            entryType = "Credit"
        ElseIf Target.Column = COL_DEBIT Then
            entryType = "Debit"
        End If

        sel = Me.Cells(Target.Row, "L").Value
        If Len(Trim$(sel)) > 0 Then
            sheetList = Split(sel, ",")
            For i = LBound(sheetList) To UBound(sheetList)
                nm = Trim$(CStr(sheetList(i)))
                If Len(nm) > 0 Then
                    ' Look for existing row for this sheet
                    Set f = bs.Range("A:A").Find(What:=nm, LookAt:=xlWhole, MatchCase:=False)
                    If f Is Nothing Then
                        ' Add new row
                        lastRowBS = bs.Cells(bs.Rows.Count, "A").End(xlUp).Row + 1
                        bs.Cells(lastRowBS, "A").Value = nm
                        If entryType = "Credit" Then
                            bs.Cells(lastRowBS, "C").Value = Target.Value
                        ElseIf entryType = "Debit" Then
                            bs.Cells(lastRowBS, "D").Value = Target.Value
                        End If
                        ' Calculate balance
                        bs.Cells(lastRowBS, "B").Value = Val(bs.Cells(lastRowBS, "C").Value) - Val(bs.Cells(lastRowBS, "D").Value)
                    Else
                        ' Update existing row
                        If entryType = "Credit" Then
                            bs.Cells(f.Row, "C").Value = Target.Value
                        ElseIf entryType = "Debit" Then
                            bs.Cells(f.Row, "D").Value = Target.Value
                        End If
                        ' Recalculate balance
                        bs.Cells(f.Row, "B").Value = Val(bs.Cells(f.Row, "C").Value) - Val(bs.Cells(f.Row, "D").Value)
                    End If
                End If
            Next i
        End If
        Application.EnableEvents = True
    End If

CleanExit:
    If Application.EnableEvents = False Then Application.EnableEvents = True
End Sub

