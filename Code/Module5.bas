Attribute VB_Name = "Module5"
Sub SetupStockBalanceSummaryButton()
    Dim ws As Worksheet, btn As Button
    Dim btnExists As Boolean

    ' Create or reference Stock Balance sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Stock Balance")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Stock Balance"
    End If

    ' Check if button exists
    btnExists = False
    For Each btn In ws.Buttons
        If btn.Caption = "Refresh Stock Balances" Then
            btnExists = True
            Exit For
        End If
    Next btn

    ' Create button if not exists
    If Not btnExists Then
        Set btn = ws.Buttons.Add(10, 10, 160, 30)
        btn.Caption = "Refresh Stock Balances"
        btn.OnAction = "CollectStockBalances"
    End If

    MsgBox "Setup Complete! Click the button to refresh stock balances.", vbInformation
End Sub
