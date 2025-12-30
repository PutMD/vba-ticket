Sub SearchTicket()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tickets")
    
    Dim ticketID As String
    ticketID = InputBox("Enter Ticket ID to search:", "Search Ticket")
    
    If ticketID = "" Then Exit Sub
    
    Dim found As Range
    Set found = ws.Columns("A").Find(ticketID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not found Is Nothing Then
        MsgBox "Ticket ID: " & ws.Cells(found.Row, 1).Value & vbCrLf & _
               "Requester: " & ws.Cells(found.Row, 3).Value & vbCrLf & _
               "Department: " & ws.Cells(found.Row, 4).Value & vbCrLf & _
               "Category: " & ws.Cells(found.Row, 5).Value & vbCrLf & _
               "Priority: " & ws.Cells(found.Row, 7).Value & vbCrLf & _
               "Status: " & ws.Cells(found.Row, 8).Value & vbCrLf & _
               "SLA Hours: " & ws.Cells(found.Row, 10).Value, vbInformation
    Else
        MsgBox "Ticket ID not found.", vbExclamation
    End If
End Sub
