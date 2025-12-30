' ---------------- USERFORM CODE ----------------

' Initialize dropdowns
Private Sub UserForm_Initialize()
    cmbDept.AddItem "IT"
    cmbDept.AddItem "HR"
    cmbDept.AddItem "Finance"
    cmbDept.AddItem "Operations"
    
    cmbCategory.AddItem "Hardware"
    cmbCategory.AddItem "Software"
    cmbCategory.AddItem "Network"
    cmbCategory.AddItem "Account"
    
    cmbPriority.AddItem "Low"
    cmbPriority.AddItem "Medium"
    cmbPriority.AddItem "High"
End Sub

' Submit ticket
Private Sub btnSubmit_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tickets")
    
    ' Validate fields
    If txtName.Value = "" Or cmbDept.Value = "" Or cmbCategory.Value = "" Or cmbPriority.Value = "" Or txtDescription.Value = "" Then
        MsgBox "Please fill in all fields.", vbExclamation
        Exit Sub
    End If
    
    ' Find next empty row
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Generate Ticket ID
    Dim ticketID As String
    ticketID = "TICK" & Format(nextRow - 1, "0000")
    
    ' Write to Tickets sheet
    ws.Cells(nextRow, 1).Value = ticketID
    ws.Cells(nextRow, 2).Value = Now()
    ws.Cells(nextRow, 3).Value = txtName.Value
    ws.Cells(nextRow, 4).Value = cmbDept.Value
    ws.Cells(nextRow, 5).Value = cmbCategory.Value
    ws.Cells(nextRow, 6).Value = txtDescription.Value
    ws.Cells(nextRow, 7).Value = cmbPriority.Value
    ws.Cells(nextRow, 8).Value = "Open"
    
    ' Clear form
    txtName.Value = ""
    cmbDept.Value = ""
    cmbCategory.Value = ""
    txtDescription.Value = ""
    cmbPriority.Value = ""
    
    MsgBox "Ticket submitted successfully! ID: " & ticketID, vbInformation
    
    ' Update SLA highlight and dashboard
    Call HighlightOverdue
    Call UpdateDashboard
End Sub

