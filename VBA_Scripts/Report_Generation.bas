Attribute VB_Name = "Module2"
Sub GenerateStructuredSecurityReport()
    Dim WordApp As Object
    Dim WordDoc As Object
    Dim ws As Worksheet
    Dim lastRow As Integer
    Dim i As Integer
    Dim filePath As String
    Dim totalRisks As Integer, overdueRisks As Integer, highRisks As Integer, criticalRisks As Integer, closedRisks As Integer
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("security_risk_data") ' Change to your actual sheet name
    
    ' Find last row with data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Initialize risk counters
    totalRisks = lastRow - 1 ' Excluding header row
    overdueRisks = 0
    highRisks = 0
    criticalRisks = 0
    closedRisks = 0

    ' Count risks based on their status and severity
    For i = 2 To lastRow
        If ws.Cells(i, 5).Value = "Overdue" Then overdueRisks = overdueRisks + 1
        If ws.Cells(i, 3).Value = "High" Then highRisks = highRisks + 1
        If ws.Cells(i, 3).Value = "Critical" Then criticalRisks = criticalRisks + 1
        If ws.Cells(i, 5).Value = "Closed" Then closedRisks = closedRisks + 1
    Next i

    ' Create Word application
    Set WordApp = CreateObject("Word.Application")
    WordApp.Visible = True ' Set to False if you don't want to see the document

    ' Create a new Word document
    Set WordDoc = WordApp.Documents.Add

    ' Add Title
    With WordDoc.Range
        .Text = "Information Security Risk & Governance Summary Report" & vbNewLine & _
                "Date: " & Format(Date, "MMMM DD, YYYY") & vbNewLine & vbNewLine
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' Add Key Metrics Section
    WordDoc.Content.InsertAfter vbNewLine & "Key Risk Metrics:" & vbNewLine & _
        "-----------------------------------------------------" & vbNewLine & _
        "Total Security Risks: " & totalRisks & vbNewLine & _
        "Overdue Risks: " & overdueRisks & vbNewLine & _
        "High-Risk Findings: " & highRisks & vbNewLine & _
        "Critical Risks: " & criticalRisks & vbNewLine & _
        "Closed Risks: " & closedRisks & vbNewLine & vbNewLine

    ' Add Summary Table Headers
    WordDoc.Content.InsertAfter "Security Risk Overview:" & vbNewLine & _
        "-----------------------------------------------------" & vbNewLine & _
        "Finding ID  |  Security Risk  |  Risk Level  |  Status  |  Due Date" & vbNewLine

    ' Loop through risk data and append to the report
    For i = 2 To lastRow
        WordDoc.Content.InsertAfter ws.Cells(i, 1).Value & " | " & ws.Cells(i, 2).Value & " | " & _
                     ws.Cells(i, 3).Value & " | " & ws.Cells(i, 5).Value & " | " & _
                     ws.Cells(i, 6).Value & vbNewLine
    Next i

    ' Add Actionable Recommendations Section
    WordDoc.Content.InsertAfter vbNewLine & "Actionable Recommendations:" & vbNewLine & _
        "-----------------------------------------------------" & vbNewLine & _
        "1. Implement automated tracking to reduce overdue risks." & vbNewLine & _
        "2. Prioritize resolution of critical security risks within 7 days." & vbNewLine & _
        "3. Conduct quarterly compliance audits to assess risk mitigation progress." & vbNewLine & _
        "4. Increase security training for departments with repeated overdue risks." & vbNewLine

    ' Save the report
    filePath = ThisWorkbook.Path & "\SecurityRiskSummaryReport.docx"
    WordDoc.SaveAs filePath
    
    ' Save as PDF
    WordDoc.SaveAs2 ThisWorkbook.Path & "\SecurityRiskReport.pdf", 17 ' PDF format

    ' Cleanup
    WordDoc.Close
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing

    MsgBox "Security Risk Summary Report generated successfully!", vbInformation, "Report Complete"
End Sub


