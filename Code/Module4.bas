Attribute VB_Name = "Module4"
Sub CollectStockBalances()
    Dim ws As Worksheet, summaryWs As Worksheet
    Dim headerCell As Range, balanceCell As Range
    Dim lastRow As Long, col As Long

    ' Create or reference Stock Balance sheet
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Worksheets("Stock Balance")
    On Error GoTo 0
    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Worksheets.Add
        summaryWs.Name = "Stock Balance"
    End If

    ' Clear previous summary
    summaryWs.Cells.Clear
    summaryWs.Range("A1").Value = "Sheet Name"
    summaryWs.Range("B1").Value = "Total Stock Balance"

    lastRow = 2

    ' Loop through each sheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> summaryWs.Name And ws.Name <> "Master" And ws.Name <> "Ledger Balance" Then
            
            Set headerCell = ws.Rows(2).Find(What:="Stock Balance", LookIn:=xlValues, LookAt:=xlWhole)
            If Not headerCell Is Nothing Then
                col = headerCell.Column
                Set balanceCell = ws.Cells(1, col)
                If balanceCell.Value <> "" Then
                    summaryWs.Cells(lastRow, 1).Value = ws.Name
                    summaryWs.Cells(lastRow, 2).Value = balanceCell.Value
                    lastRow = lastRow + 1
                End If
            End If
        End If
    Next ws

    MsgBox "Stock Balance summary updated!", vbInformation
End Sub
