Attribute VB_Name = "Module1"
Option Explicit

Sub RefreshSheetList()
    Dim ws As Worksheet, r As Long, lastRow As Long
    With Worksheets("Master")
        .Range("K2:K9999").ClearContents
        r = 2
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name <> .Name Then
                .Cells(r, "K").Value = ws.Name
                r = r + 1
            End If
        Next ws

        ' Create/refresh a named range "SheetNames" that points to E2:Elast
        lastRow = .Cells(.Rows.Count, "K").End(xlUp).Row
        If lastRow < 2 Then lastRow = 2
        On Error Resume Next
        ThisWorkbook.Names("SheetNames").Delete
        On Error GoTo 0
        ThisWorkbook.Names.Add Name:="SheetNames", _
            RefersTo:="=" & .Name & "!$K$2:$K$" & lastRow, Visible:=True
    End With
    MsgBox "Sheet list refreshed.", vbInformation
End Sub

