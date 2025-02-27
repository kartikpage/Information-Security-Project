Attribute VB_Name = "Module1"
Sub Highlight_Overdue_Risks()
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim i As Integer
    
    ' Check if the worksheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("security_risk_data") ' Change to your actual sheet name
    On Error GoTo 0
    
    ' If worksheet is not found, show an error message and exit
    If ws Is Nothing Then
        MsgBox "Worksheet 'RiskData' not found! Please check the sheet name.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Determine the last row with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through each row and highlight overdue risks
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, 6).Value) Then ' Ensure the value is a valid date
            If ws.Cells(i, 6).Value < Date And ws.Cells(i, 5).Value <> "Closed" Then
                ws.Cells(i, 6).Interior.Color = RGB(255, 0, 0) ' Red for overdue
            ElseIf ws.Cells(i, 5).Value = "Closed" Then
                ws.Cells(i, 6).Interior.Color = RGB(0, 255, 0) ' Green for closed
            End If
        End If
    Next i
    
    MsgBox "Overdue risks have been highlighted successfully!", vbInformation, "Process Complete"
End Sub

