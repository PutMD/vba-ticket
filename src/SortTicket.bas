Sub SortTickets()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tickets")
    
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range("G2:G" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ws.Sort.SetRange ws.Range("A1:J" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    ws.Sort.Header = xlYes
    ws.Sort.Apply
End Sub

