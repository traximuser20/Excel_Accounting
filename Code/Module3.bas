Attribute VB_Name = "Module3"
Sub SetupLedgerBalanceSummaryButton()
    Dim ws As Worksheet, btn As Button
    Dim btnExists As Boolean
    
    ' Create or reference Ledger Balance sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Ledger Balance")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Ledger Balance"
    End If

    ' Check if button already exists
    btnExists = False
    For Each btn In ws.Buttons
        If btn.Caption = "Refresh Ledger Balances" Then
            btnExists = True
            Exit For
        End If
    Next btn

    ' Create button if not exists
    If Not btnExists Then
        Set btn = ws.Buttons.Add(10, 10, 150, 30)
        btn.Caption = "Refresh Ledger Balances"
        btn.OnAction = "CollectLedgerBalances"
    End If

    MsgBox "Setup Complete! Click the button to refresh ledger balances.", vbInformation
End Sub
