' ---------------- MODULE 1: STANDARD MACROS ----------------

' Open the ticket submission form
Public Sub OpenTicketForm()
    frmNewTicket.Show
End Sub

' Highlight overdue tickets
Public Sub HighlightOverdue()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tickets")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 8).Value <> "Closed" Then
            Dim hoursOpen As Double
            hoursOpen = (Now() - ws.Cells(i, 2).Value) * 24
            If hoursOpen > 24 Then
                ws.Rows(i).Interior.Color = RGB(255, 200, 200) ' light red
            Else
                ws.Rows(i).Interior.ColorIndex = 0
            End If
        Else
            ws.Rows(i).Interior.ColorIndex = 0
        End If
    Next i
End Sub

' Update dashboard metrics
Public Sub UpdateDashboard()
    Dim wsTickets As Worksheet
    Dim wsDash As Worksheet
    Set wsTickets = ThisWorkbook.Sheets("Tickets")
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    Dim lastRow As Long
    lastRow = wsTickets.Cells(wsTickets.Rows.Count, "A").End(xlUp).Row
    
    Dim totalTickets As Long, openTickets As Long, closedTickets As Long, overdueTickets As Long
    totalTickets = lastRow - 1
    openTickets = 0
    closedTickets = 0
    overdueTickets = 0
    
    Dim i As Long
    For i = 2 To lastRow
        Select Case wsTickets.Cells(i, 8).Value
            Case "Open", "In Progress"
                openTickets = openTickets + 1
                If (Now() - wsTickets.Cells(i, 2).Value) * 24 > 24 Then
                    overdueTickets = overdueTickets + 1
                End If
            Case "Closed"
                closedTickets = closedTickets + 1
        End Select
    Next i
    
    ' Update Dashboard sheet
    wsDash.Cells(1, 2).Value = totalTickets
    wsDash.Cells(2, 2).Value = openTickets
    wsDash.Cells(3, 2).Value = closedTickets
    wsDash.Cells(4, 2).Value = overdueTickets
End Sub

' Close a ticket by ID
Public Sub CloseTicket()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tickets")
    
    Dim ticketID As String
    ticketID = InputBox("Enter Ticket ID to close:", "Close Ticket")
    
    If ticketID = "" Then Exit Sub
    
    Dim found As Range
    Set found = ws.Columns("A").Find(ticketID, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not found Is Nothing Then
        ws.Cells(found.Row, 8).Value = "Closed"
        ws.Cells(found.Row, 9).Value = Now() ' DateClosed
        ws.Cells(found.Row, 10).Value = Round((ws.Cells(found.Row, 9).Value - ws.Cells(found.Row, 2).Value) * 24, 2)
        MsgBox "Ticket " & ticketID & " closed successfully!", vbInformation
        Call HighlightOverdue
        Call UpdateDashboard
    Else
        MsgBox "Ticket ID not found.", vbExclamation
    End If
End Sub


